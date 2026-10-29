# -*- coding: utf-8 -*-
"""
每日库存预警 → 钉钉群
================================
读 inventory_data.json，推送两类预警（只含常规备货分类，不含定制）：
  1. 缺货    ：与看板页面同源 —— level == "缺货"（在售 且 可发 ≤ 0）
  2. 急需关注：与看板页面同源 —— level == "急需关注"（在售 且 周转 < urgent_days）
消息按 旅行箱 / 包袋 / 旅行生活 分层列出，钉钉群自定义机器人（加签）markdown 推送。

【v0.13.9 口径修复 2026-10-29】
  旧版缺货段用 kf < 0（可发严格小于 0），且分类写死 旅行箱/包袋，导致：
    · 4 条「停采」尾货（可发 -1 / -21）被当成缺货推送，而页面缺货为 0 条；
    · 页面把停采/未上市剔除在缺货之外（build_inventory.py 方案A），钉钉却照推。
  现统一为读 level 字段（页面同一判据），分类与阈值一律从 data["config"] 取，
  钉钉与页面永远同源，不再各自维护一套规则。

配置：dingtalk_config.txt（两行，不入库）
  第一行：webhook 完整地址
  第二行：加签 secret（SEC 开头；无加签则留空）

用法：
  py send_dingtalk.py            # 推当前 inventory_data.json
"""
import os
import sys
import json
import time
import hmac
import hashlib
import base64
import urllib.request
import urllib.parse
import datetime

OUT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(OUT_DIR, "dingtalk_config.txt")
DATA_PATH = os.path.join(OUT_DIR, "inventory_data.json")

WATCH_CATS = {"旅行箱", "包袋"}        # 兜底：data.config.stocked 缺失时用（v0.13.9 起以 config 为准）
TIER_ORDER = ["旅行箱", "包袋", "旅行生活"]   # 分层顺序（页面 stocked 三类）
FALLBACK_URGENT_DAYS = 7               # 兜底：data.config.urgent_days 缺失时用
MAX_PER_TIER = 15                      # 每层最多列出的条数（防刷屏）


def page_scope(data):
    """分类范围 / 分层顺序 / 紧迫阈值 —— 全部取自看板自己写进 JSON 的 config。

    v0.13.9：钉钉不再自己写死分类与阈值。页面改口径（如新增 stocked 分类、
    调整 urgent_days）时，钉钉自动跟随，避免两份规则各说各话。
    """
    cfg = data.get("config") or {}
    stocked = cfg.get("stocked") or sorted(WATCH_CATS)
    order = [c for c in TIER_ORDER if c in stocked]
    order += [c for c in stocked if c not in order]      # config 新增的分类追加到末尾，保证不漏推
    urgent = cfg.get("urgent_days")
    if not isinstance(urgent, (int, float)):
        urgent = FALLBACK_URGENT_DAYS
    return set(stocked), order, urgent


def is_stockout(s):
    """缺货判定 —— 与看板页面 level 同源（build_inventory.py：在售 且 可发 ≤ 0）。

    v0.13.9 前的旧判据 kf < 0 是错的：停采尾货可发为负，被当成缺货刷屏，
    而页面缺货为 0 条。此处绝不再回到 kf < 0。
    仅在数据缺 level 字段（老结构）时回退公式，且同样排除停采/未上市/定制。
    """
    lv = s.get("level")
    if lv:
        return lv == "缺货"
    life = s.get("lifecycle")
    return life in (None, "", "在售") and float(s.get("kf") or 0) <= 0


def is_urgent(s, urgent_days):
    """急需关注判定 —— 与看板页面 level 同源（在售 且 0 < 可发 且 周转 < urgent_days）。"""
    lv = s.get("level")
    if lv:
        return lv == "急需关注"
    life = s.get("lifecycle")
    to = s.get("turnover")
    return (life in (None, "", "在售") and to is not None
            and to < urgent_days and float(s.get("kf") or 0) > 0)


def load_config(path=CONFIG):
    if not os.path.exists(path):
        print("[warn] 未找到 dingtalk_config.txt，跳过钉钉推送", file=sys.stderr)
        return None
    lines = [l.strip() for l in open(path, encoding="utf-8")
             if l.strip() and not l.startswith("#")]
    webhook = lines[0] if len(lines) > 0 else ""
    secret = lines[1] if len(lines) > 1 else ""
    if not webhook.startswith("http"):
        print("[warn] dingtalk_config.txt 格式错误（第1行 webhook、第2行 secret）", file=sys.stderr)
        return None
    return {"webhook": webhook, "secret": secret}


def build_sign(secret, ts):
    """钉钉加签：HMAC-SHA256(timestamp+'\n'+secret) → base64 → urlencode"""
    string_to_sign = f"{ts}\n{secret}"
    hmac_code = hmac.new(secret.encode("utf-8"), string_to_sign.encode("utf-8"),
                         digestmod=hashlib.sha256).digest()
    return urllib.parse.quote_plus(base64.b64encode(hmac_code))


def send(cfg, payload):
    ts = str(round(time.time() * 1000))
    sep = "&" if "?" in cfg["webhook"] else "?"
    url = f"{cfg['webhook']}{sep}timestamp={ts}&sign={build_sign(cfg['secret'], ts)}"
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.loads(r.read().decode("utf-8"))


def short_name(s):
    """用 系列+颜色+尺寸 精简展示（无则回退名称去 ITO 前缀）"""
    parts = [s.get("series") or "", s.get("color") or "", s.get("size") or ""]
    short = " ".join(p for p in parts if p)
    if short:
        return short
    return (s.get("name") or "").replace("ITO ", "").strip()


def build_message(data):
    d = datetime.date.fromisoformat(data["source_date"])
    cats, tier_order, urgent_days = page_scope(data)
    skus = [s for s in data["skus"] if s["cat"] in cats]
    out = [f"## 📦 ITO 库存预警 · {d.month}月{d.day}日", f"> 数据快照：{data['source_date']}（品牌 ITO · 在售）"]

    # ===== 1. 缺货（与页面同源：level=="缺货"，即在售且可发≤0）=====
    stk = sorted([s for s in skus if is_stockout(s)],
                 key=lambda s: float(s.get("kf") or 0))
    out.append("")
    out.append("**🔴 缺货（在售且可发 ≤ 0）**")
    if not stk:
        out.append("✅ 无")
    for tier in tier_order:
        items = [s for s in stk if s["cat"] == tier]
        if not items:
            continue
        out.append(f"**【{tier}】**")
        for s in items:
            line = f"- {short_name(s)}：可发 {int(s['kf'])}，在途 {int(s.get('in_transit') or 0)}"
            if s.get("arrival_qty"):
                eta = (s.get("arrival_eta") or "")[5:].replace("-", "/")
                qty = int(s["arrival_qty"])
                nbat = len(s.get("arrival_batches") or [])
                # 首到日 + 起N批共X；单批不写「起1批」（与看板单批不显示批数标签一致）
                if nbat >= 2:
                    line += f"，预计到货 {eta}起{nbat}批共{qty}"
                else:
                    line += f"，预计到货 {eta}共{qty}"
            out.append(line)

    # ===== 2. 急需关注（与页面同源：level=="急需关注"，周转<urgent_days）=====
    urg = sorted([s for s in skus if is_urgent(s, urgent_days)],
                 key=lambda s: float(s.get("turnover") or 0))
    out.append("")
    out.append(f"**🟠 急需关注（周转 < {urgent_days:g} 天）**")
    if not urg:
        out.append("✅ 无")
    for tier in tier_order:
        items = [s for s in urg if s["cat"] == tier]
        if not items:
            continue
        out.append(f"**【{tier}】**")
        for s in items[:MAX_PER_TIER]:
            out.append(f"- {short_name(s)}：可发 {int(s['kf'])}，周转 {s['turnover']:.1f}天")
        rest = len(items) - MAX_PER_TIER
        if rest > 0:
            out.append(f"- ……其余 {rest} 条见看板")

    out.append("")
    out.append("> 口径与看板一致：缺货=在售且可发≤0；停采 / 未上市 / 定制不计入")
    out.append("📊 完整看板：https://ito-inventory-dashboard.vercel.app")
    return "\n".join(out)


def caliber_check(data):
    """口径自检：钉钉各段条数 vs 看板 KPI，不一致就告警（只告警，不阻断推送）。

    起因（2026-10-29）：页面缺货 0 条，钉钉却推 4 条停采尾货，两边静默分歧了很久。
    把比对做成每日固定输出，下次再漂移当天就能看见。
    """
    cats, _order, urgent_days = page_scope(data)
    skus = [s for s in data["skus"] if s["cat"] in cats]
    n_stk = sum(1 for s in skus if is_stockout(s))
    n_urg = sum(1 for s in skus if is_urgent(s, urgent_days))
    kpi = data.get("kpi") or {}
    k_stk = kpi.get("out_of_stock")
    k_warn = kpi.get("warning")
    msg = f"[口径] 钉钉缺货 {n_stk} 条 / 页面缺货 {k_stk} 条；钉钉急需关注 {n_urg} 条 / 页面预警 {k_warn} 条（含缺货风险，可不等于）"
    warn = (k_stk is not None and n_stk != k_stk)
    if warn:
        msg += f"  ⚠ 缺货口径不一致（页面 level 与钉钉判定已漂移）"
    return msg, warn


def main():
    data = json.load(open(DATA_PATH, encoding="utf-8"))
    cfg = load_config()
    if not cfg:
        return
    check_msg, drift = caliber_check(data)
    print(check_msg, file=sys.stderr if drift else sys.stdout)
    text = build_message(data)
    payload = {
        "msgtype": "markdown",
        "markdown": {"title": f"ITO 库存预警 {data['source_date']}", "text": text},
    }
    resp = send(cfg, payload)
    if resp.get("errcode") == 0:
        print(f"[ok] 钉钉推送成功（{data['source_date']}，缺货/急需关注）")
    else:
        print(f"[错误] 钉钉推送失败: {resp}", file=sys.stderr)


if __name__ == "__main__":
    main()
