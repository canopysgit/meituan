#!/usr/bin/env python3
"""
美团GEO监测看板 - 自动化构建流水线
功能：从飞书群下载Excel → 解析 → 生成JSON → git push → Vercel自动部署

用法:
  python3 build_pipeline.py                    # 从飞书群拉最新Excel并处理
  python3 build_pipeline.py --file <path.xlsx>  # 直接处理指定文件
  python3 build_pipeline.py --backfill          # 回填历史数据
  python3 build_pipeline.py --deploy-only       # 仅git push（不拉新数据）
"""

import os
import sys
import json
import subprocess
import tempfile
from pathlib import Path
from datetime import datetime

SCRIPT_DIR = Path(__file__).parent
SITE_DIR = SCRIPT_DIR.parent
DATA_DIR = SITE_DIR / "data"
PARSER = SCRIPT_DIR / "parse_excel.py"

# 飞书群配置
CHAT_ID = "oc_41bc8e496a0d9bb93498a6a6cf6f30c1"


def run_cmd(cmd: str, check=True, capture=True) -> subprocess.CompletedProcess:
    """执行shell命令"""
    result = subprocess.run(
        cmd, shell=True, cwd=str(SITE_DIR),
        capture_output=capture, text=True, timeout=120
    )
    if check and result.returncode != 0:
        print(f"  ❌ 命令失败: {cmd}")
        if result.stderr:
            print(f"     {result.stderr[:200]}")
    return result


def check_auth() -> bool:
    """检查飞书授权状态"""
    result = run_cmd('lark-cli auth check --scope "im:message.group_msg:get_as_user" 2>&1')
    return '"ok": true' in result.stdout


def fetch_group_messages() -> list:
    """从飞书群拉取最近消息"""
    print("📡 正在从飞书群拉取消息...")
    result = run_cmd(
        f'lark-cli im +chat-messages-list --chat-id "{CHAT_ID}" --page-size 50 --order desc --as user 2>&1'
    )
    if result.returncode != 0:
        print(f"  ❌ 拉取消息失败: {result.stdout[:200]}")
        return []

    try:
        # 提取JSON部分（跳过可能的warning行）
        raw = result.stdout
        json_start = raw.find('{')
        if json_start < 0:
            print("  ❌ 无JSON输出")
            return []
        data = json.loads(raw[json_start:])
        if not data.get('ok'):
            print(f"  ❌ API错误: {data.get('error', {}).get('message', 'unknown')}")
            return []
        return data.get('data', {}).get('items', [])
    except json.JSONDecodeError as e:
        print(f"  ❌ JSON解析失败: {e}")
        return []


def find_excel_attachments(messages: list) -> list:
    """从消息列表中找到Excel附件"""
    excels = []
    for msg in messages:
        msg_type = msg.get('msg_type', '')
        if msg_type != 'file':
            continue

        body = msg.get('body', {})
        content_str = body.get('content', '{}')
        try:
            content = json.loads(content_str)
        except:
            continue

        filename = content.get('file_name', '')
        file_key = content.get('file_key', '')
        if not filename.endswith('.xlsx'):
            continue

        # 解析日期
        create_time = msg.get('create_time', '0')
        ts = int(create_time) / 1000 if len(create_time) > 10 else int(create_time)
        msg_time = datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M')

        excels.append({
            'filename': filename,
            'file_key': file_key,
            'msg_time': msg_time,
            'sender': msg.get('sender', {}).get('id', 'unknown'),
        })

    return excels


def download_file(file_key: str, filename: str) -> Path:
    """从飞书下载文件"""
    tmp_dir = Path(tempfile.mkdtemp(prefix='meituan_'))
    output = tmp_dir / filename
    result = run_cmd(
        f'lark-cli im +file-download --file-key "{file_key}" --output "{output}" 2>&1'
    )
    if result.returncode != 0 or not output.exists():
        print(f"  ❌ 下载失败: {filename}")
        return None
    return output


def parse_files(file_paths: list):
    """调用解析脚本处理Excel文件"""
    if not file_paths:
        return
    paths_str = ' '.join(f'"{p}"' for p in file_paths)
    result = run_cmd(f'python3 "{PARSER}" {paths_str} 2>&1', check=False)
    print(result.stdout)
    if result.stderr:
        print(result.stderr[:500])


def git_push():
    """Git commit + push"""
    print("\n📦 正在推送到 GitHub...")
    run_cmd('git add -A')
    
    # 检查是否有变更
    result = run_cmd('git status --porcelain', check=False)
    if not result.stdout.strip():
        print("  ℹ️ 无变更，跳过push")
        return True

    date_str = datetime.now().strftime('%Y-%m-%d %H:%M')
    run_cmd(f'git commit -m "📊 看板数据更新 {date_str}" 2>&1')
    result = run_cmd('git push 2>&1', check=False)
    if result.returncode != 0:
        # 尝试强制IPv4
        print("  ⚠️ 尝试IPv4 push...")
        result = run_cmd('git push -4 2>&1', check=False)
    
    if result.returncode == 0:
        print("  ✅ Push成功，Vercel将自动部署")
        return True
    else:
        print(f"  ❌ Push失败: {result.stdout[:200]}")
        return False


def pipeline_from_group():
    """完整流水线：从飞书群拉数据 → 解析 → 部署"""
    print("🚀 美团GEO看板自动化流水线启动")
    print(f"   时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()

    # 1. 检查授权
    if not check_auth():
        print("❌ 飞书授权已过期，请运行:")
        print('   lark-cli auth login --scope "im:message.group_msg:get_as_user"')
        return False
    print("✅ 飞书授权正常")

    # 2. 拉取群消息
    messages = fetch_group_messages()
    if not messages:
        print("⚠️ 群内无消息")
        return True
    print(f"   找到 {len(messages)} 条消息")

    # 3. 找Excel附件
    excels = find_excel_attachments(messages)
    if not excels:
        print("⚠️ 未找到Excel附件")
        return True
    print(f"   找到 {len(excels)} 个Excel文件:")
    for e in excels:
        print(f"     📄 {e['filename']} ({e['msg_time']})")

    # 4. 下载Excel
    print("\n📥 正在下载Excel文件...")
    downloaded = []
    for e in excels:
        path = download_file(e['file_key'], e['filename'])
        if path:
            downloaded.append(path)
            print(f"  ✅ {e['filename']}")

    if not downloaded:
        print("⚠️ 无文件下载成功")
        return True

    # 5. 解析
    print(f"\n📊 正在解析 {len(downloaded)} 个文件...")
    parse_files(downloaded)

    # 6. 部署
    return git_push()


def pipeline_from_file(filepath: str):
    """从本地文件处理"""
    print(f"📄 处理本地文件: {filepath}")
    parse_files([filepath])
    return git_push()


def pipeline_deploy_only():
    """仅部署"""
    return git_push()


def pipeline_backfill():
    """回填历史数据"""
    print("🔄 回填历史Excel数据...")
    
    # 扫描所有目录下的xlsx
    search_dirs = [
        SITE_DIR.parent / "美团医美看板" / "data",
        SITE_DIR.parent / "美团美眼",
        SITE_DIR.parent / "美团祛痘",
    ]
    
    files = []
    for d in search_dirs:
        if d.exists():
            files.extend(d.glob("*.xlsx"))
    
    # 去重
    unique = {f.name: f for f in files}
    files = sorted(unique.values())
    
    print(f"   找到 {len(files)} 个历史Excel文件")
    for f in files:
        print(f"   📄 {f.name}")
    
    parse_files([str(f) for f in files])
    return git_push()


def main():
    if "--file" in sys.argv:
        idx = sys.argv.index("--file")
        if idx + 1 < len(sys.argv):
            pipeline_from_file(sys.argv[idx + 1])
        else:
            print("❌ --file 需要指定文件路径")
    elif "--deploy-only" in sys.argv:
        pipeline_deploy_only()
    elif "--backfill" in sys.argv:
        pipeline_backfill()
    else:
        pipeline_from_group()


if __name__ == "__main__":
    main()
