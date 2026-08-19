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

    # vercel.cmd 路径
    vercel_cmd = os.path.join(BASE, 'node_modules', '.bin', 'vercel.cmd')
    if not os.path.exists(vercel_cmd):
        print('[ERROR] 未找到 node_modules/.bin/vercel.cmd')
        log_run('deploy FAIL (vercel.cmd not found)')
        return 1

    deploy_dir = os.path.join(BASE, 'deploy')
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

    print('[ERROR] 3 次尝试均失败，详见 deploy_vercel_fail_*.log')
    log_run('deploy FAIL (3 attempts)')
    return 1

if __name__ == '__main__':
    sys.exit(main())