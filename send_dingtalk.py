# -*- coding: utf-8 -*-
"""
每日库存预警 → 钉钉群
================================
读 inventory_data.json，推送两类预警（只含 旅行箱 / 包袋，不含定制）：
  1. 缺货：可发库存 < 0
  2. 急需关注：周转天数 < 7
消息按 旅行箱 / 包袋 分层列出，钉钉群自定义机器人（加签）markdown 推送。

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

WATCH_CATS = {"旅行箱", "包袋"}        # 只推这两个分类（在售，品牌已过滤 ITO）
TIER_ORDER = ["旅行箱", "包袋"]        # 分层顺序
MAX_PER_TIER = 15                      # 每层最多列出的条数（防刷屏）


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
    skus = [s for s in data["skus"] if s["cat"] in WATCH_CATS]
    out = [f"## 📦 ITO 库存预警 · {d.month}月{d.day}日", f"> 数据快照：{data['source_date']}（品牌 ITO · 在售）"]

    # ===== 1. 缺货（可发<0）=====
    stk = sorted([s for s in skus if s["kf"] < 0], key=lambda s: s["kf"])
    out.append("")
    out.append("**🔴 缺货（可发 < 0）**")
    if not stk:
        out.append("✅ 无")
    for tier in TIER_ORDER:
        items = [s for s in stk if s["cat"] == tier]
        if not items:
            continue
        out.append(f"**【{tier}】**")
        for s in items:
            line = f"- {short_name(s)}：可发 {int(s['kf'])}，在途 {int(s.get('in_transit') or 0)}"
            if s.get("arrival_qty"):
                eta = (s.get("arrival_eta") or "")[5:].replace("-", "/")
                line += f"，预计到货 {eta}({int(s['arrival_qty'])})"
            out.append(line)

    # ===== 2. 急需关注（周转<7天，仅有效周转：可发>0 的 SKU，避免与缺货模块重复）=====
    urg = sorted([s for s in skus if s.get("turnover") is not None and s["turnover"] < 7 and s["kf"] > 0],
                 key=lambda s: s["turnover"])
    out.append("")
    out.append("**🟠 急需关注（周转 < 7 天）**")
    if not urg:
        out.append("✅ 无")
    for tier in TIER_ORDER:
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
    out.append("📊 完整看板：https://ito-inventory-dashboard.vercel.app")
    return "\n".join(out)


def main():
    data = json.load(open(DATA_PATH, encoding="utf-8"))
    cfg = load_config()
    if not cfg:
        return
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
