# -*- coding: utf-8 -*-
"""
Vercel 部署脚本（2026-08-19 新增，替代 bat 内 for 循环，解决 errorlevel 误判）
- 读取 vercel_token.txt
- 调用 vercel CLI 部署 deploy/ 目录
- 失败自动重试 3 次（每次间隔 5 秒）
- 完整输出写入 deploy_vercel.log，失败副本 deploy_vercel_fail_N.log
- 返回码：0=成功 1=失败
"""
import os
import subprocess
import sys
import time
import datetime

BASE = r'D:\E盘文件\ITO库存看板系统'
TOKEN_FILE = os.path.join(BASE, 'vercel_token.txt')
LOG_FILE = os.path.join(BASE, 'deploy_vercel.log')
RUN_LOG = os.path.join(BASE, 'run.log')

def log_run(msg):
    try:
        with open(RUN_LOG, 'a', encoding='utf-8') as f:
            f.write(f'[{datetime.datetime.now().strftime("%Y/%m/%d %a %H:%M:%S.%f")[:-3]}] {msg}\n')
    except Exception:
        pass

def verify_online():
    """部署后自检：线上 inventory_data.json 的 generated_at 与本地一致 + 页面含更新时间轴标记。
    容忍 Vercel CDN 缓存延迟（最多重试 5 次 × 5 秒）。
    2026-08-19 新增：防止"部署成功但线上实际没更新/页面缺时间轴"的同类问题再次发生。"""
    import json as _json
    import urllib.request
    base_url = 'https://ito-inventory-dashboard.vercel.app'
    local_path = os.path.join(BASE, 'deploy', 'inventory_data.json')
    try:
        with open(local_path, encoding='utf-8') as f:
            local_ts = _json.load(f).get('generated_at', '')
    except Exception as e:
        print(f'        [VERIFY] 本地数据读取失败: {e}')
        return False
    # 1) 线上数据时间戳对比（容忍缓存，重试 5 次）
    online_ts = None
    for i in range(5):
        try:
            with urllib.request.urlopen(base_url + '/inventory_data.json', timeout=30) as r:
                online_ts = _json.load(r).get('generated_at', '')
            if online_ts == local_ts:
                break
        except Exception:
            online_ts = None
        if i < 4:
            time.sleep(5)
    # 2) 页面关键标记检查（更新时间轴）
    has_tl = False
    try:
        with urllib.request.urlopen(base_url + '/', timeout=30) as r:
            html = r.read().decode('utf-8', errors='replace')
            has_tl = '更新时间轴' in html
    except Exception:
        has_tl = False
    ok = (online_ts == local_ts) and has_tl
    print(f'        [VERIFY] 线上 generated_at={online_ts} | 本地={local_ts} | 时间轴标记={has_tl} -> {"OK" if ok else "FAIL"}')
    log_run(f'deploy VERIFY {"OK" if ok else "FAIL"} (ts={online_ts}, tl={has_tl})')
    return ok


def _find_node_bin():
    """动态探测可用的 Node.js bin 目录（2026-09-08 修复：不再硬编码单一版本路径）。
    根因：WorkBuddy 托管的 versions/ 目录会在平台更新时被删除重装/改名
    （例：22.22.2 -> 22.22.2-2），硬编码路径失效 -> node 注入失败 ->
    vercel.cmd 报 '"node" 不是内部或外部命令'，此类 deploy FAIL 反复复发。
    探测顺序：current 指针（平台自动维护，最权威）-> 已知版本 -> 目录扫描兜底。
    """
    versions_root = r'C:\Users\QwQ\.workbuddy\binaries\node\versions'
    candidates = []
    # 1) current 指针文件（平台自动更新指向当前可用版本）
    current_file = os.path.join(versions_root, 'current')
    try:
        with open(current_file, encoding='utf-8') as f:
            cur = f.read().strip()
        if cur and not cur.startswith('.'):
            candidates.append(os.path.join(versions_root, cur))
    except Exception:
        pass
    # 2) 已知版本目录
    candidates += [
        os.path.join(versions_root, '22.22.2-2'),
        os.path.join(versions_root, '24.14.0'),
        os.path.join(versions_root, '22.22.2'),
    ]
    # 3) 目录扫描兜底（跳过 .deleting/.locks 等临时目录）
    try:
        candidates += [os.path.join(versions_root, d) for d in os.listdir(versions_root)
                       if not d.startswith('.') and os.path.isdir(os.path.join(versions_root, d))]
    except Exception:
        pass
    seen = set()
    for cand in candidates:
        if cand in seen:
            continue
        seen.add(cand)
        if os.path.isfile(os.path.join(cand, 'node.exe')):
            return cand
    return None


def _notify_fail(stage, detail):
    """部署失败 / 线上自检失败时推钉钉告警（2026-09-08 新增：失败主动告警，
    无人值守时不再"停更几天没人知道"）。复用 send_dingtalk.py 的配置与发送通道。"""
    try:
        import send_dingtalk
        cfg = send_dingtalk.load_config()
        if not cfg:
            print('        [notify] 无 dingtalk_config.txt，跳过失败告警')
            return
        text = ('### ITO 库存看板 · 网页部署' + stage + '\n\n'
                '> ' + detail + '\n\n'
                '日志：' + BASE + '\\deploy_vercel_fail_*.log\n\n'
                '线上：https://ito-inventory-dashboard.vercel.app')
        payload = {"msgtype": "markdown",
                   "markdown": {"title": "ITO 库存看板部署" + stage, "text": text}}
        resp = send_dingtalk.send(cfg, payload)
        print(f'        [notify] 钉钉告警已发送: {resp.get("errmsg", resp)}')
    except Exception as e:
        print(f'        [notify] 钉钉告警发送失败(不阻断): {e}')


def main():
    # 读取 token
    if not os.path.exists(TOKEN_FILE):
        print('[ERROR] 未找到 vercel_token.txt')
        log_run('deploy FAIL (no token file)')
        return 1
    with open(TOKEN_FILE, 'r', encoding='utf-8') as f:
        token = f.read().strip()
    if not token:
        print('[ERROR] vercel_token.txt 为空')
        log_run('deploy FAIL (empty token)')
        return 1

    # 清空 NODE_OPTIONS（避免 WorkBuddy safe-delete shim 干扰）
    env = os.environ.copy()
    env['NODE_OPTIONS'] = ''
    # 动态探测 Node.js bin 并注入 PATH（2026-08-20 首次修复：cmd 环境 PATH 不含 node；
    # 2026-09-08 根治：不再硬编码版本路径，平台重装/改名后自动适配）
    node_bin = _find_node_bin()
    if node_bin:
        env['PATH'] = node_bin + os.pathsep + env.get('PATH', '')
        env['NODE'] = os.path.join(node_bin, 'node.exe')
        print(f"        [fix] 已注入 Node.js bin: {node_bin}")
    else:
        print('        [warn] 未探测到可用 Node.js 目录，将退回 vercel.cmd 依赖系统 PATH')

    deploy_dir = os.path.join(BASE, 'deploy')

    # 2026-09-08：优先用探测到的 node.exe 直调 vercel 入口 vc.js，
    # 彻底绕开 vercel.cmd 对 PATH/PATHEXT 的依赖（.cmd 本质就是 node vc.js %*）
    vc_js = os.path.join(BASE, 'node_modules', 'vercel', 'dist', 'vc.js')
    if node_bin and os.path.isfile(vc_js):
        node_exe = os.path.join(node_bin, 'node.exe')
        cmd = [node_exe, vc_js, 'deploy', deploy_dir, '--prod', '--yes', '--name',
               'ito-inventory-dashboard', '--token', token]
    else:
        # 兜底：退回 vercel.cmd
        vercel_cmd = os.path.join(BASE, 'node_modules', '.bin', 'vercel.cmd')
        if not os.path.exists(vercel_cmd):
            print('[ERROR] 未找到可用的 node 与 vercel.cmd，无法部署')
            log_run('deploy FAIL (node & vercel.cmd not found)')
            return 1
        cmd = [vercel_cmd, 'deploy', deploy_dir, '--prod', '--yes', '--name',
               'ito-inventory-dashboard', '--token', token]

    max_attempts = 3
    for attempt in range(1, max_attempts + 1):
        print(f'        尝试 {attempt}/{max_attempts} ...')
        try:
            r = subprocess.run(
                cmd,
                env=env,
                capture_output=True,
                text=True,
                encoding='gbk',
                errors='replace',
                timeout=180,
            )
            combined = r.stdout + r.stderr
            # 写入日志
            with open(LOG_FILE, 'w', encoding='utf-8', errors='replace') as f:
                f.write(combined)
            # 判断成功：退出码 0 且包含 Ready/Aliased
            success = (r.returncode == 0) and ('Ready' in combined or 'Aliased' in combined)
            if success:
                print(f'        [OK] 部署成功 (attempt {attempt})')
                log_run('deploy OK')
                # 部署后自检：线上数据与本地一致 + 页面含更新时间轴（防止"部署成功但线上没生效"）
                try:
                    if not verify_online():
                        print('        [VERIFY FAIL] 线上内容与本地不一致或页面缺少更新时间轴，请人工检查！')
                        _notify_fail('异常', '部署命令成功但线上自检未通过（数据时间戳不一致或页面缺时间轴）')
                except Exception as e:
                    print(f'        [VERIFY] 自检异常（不阻断）: {e}')
                return 0
            else:
                print(f'        [{attempt}/{max_attempts}] 部署失败 (exit={r.returncode})')
                # 保存失败日志副本
                fail_log = os.path.join(BASE, f'deploy_vercel_fail_{attempt}.log')
                try:
                    with open(fail_log, 'w', encoding='utf-8', errors='replace') as f:
                        f.write(combined)
                except Exception:
                    pass
                if attempt < max_attempts:
                    print('        5 秒后重试...')
                    log_run(f'deploy retry {attempt}/{max_attempts} FAIL (exit={r.returncode})')
                    time.sleep(5)
        except subprocess.TimeoutExpired:
            print(f'        [{attempt}/{max_attempts}] 部署超时 (180s)')
            log_run(f'deploy retry {attempt}/{max_attempts} FAIL (timeout)')
            if attempt < max_attempts:
                time.sleep(5)

    # 3 次尝试均失败 → 主动钉钉告警（2026-09-08：失败不再等人发现）
    try:
        fail_log = os.path.join(BASE, 'deploy_vercel_fail_1.log')
        head = '未知错误（无失败日志）'
        if os.path.exists(fail_log):
            with open(fail_log, encoding='utf-8', errors='replace') as f:
                head = f.read(300).strip().replace('\r', ' ').replace('\n', ' ')[:200]
        _notify_fail('失败', head)
    except Exception:
        pass
    print('[ERROR] 3 次尝试均失败，详见 deploy_vercel_fail_*.log')
    log_run('deploy FAIL (3 attempts)')
    return 1

if __name__ == '__main__':
    sys.exit(main())