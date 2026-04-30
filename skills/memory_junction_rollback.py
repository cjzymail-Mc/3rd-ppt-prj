"""3 账号 auto-memory junction 回滚（保险脚本）

每个账号：
1. 用 reparse point 属性检测 memory/ 是否是 junction
2. 是则用 os.rmdir 删除 junction（**绝对不能用 shutil.rmtree，会穿透杀 D 盘本体**）
3. 从 .pre-junction-backup/<account>-<最新编号>/ 用 shutil.copytree 恢复

使用：
    python skills/memory_junction_rollback.py              # dry-run
    python skills/memory_junction_rollback.py --apply      # 实际执行
    python skills/memory_junction_rollback.py --apply --account mc   # 仅回滚 mc
"""
import argparse
import ctypes
import os
import re
import shutil
import stat
import sys

USERPROFILE = os.path.expandvars("%USERPROFILE%")

ACCOUNTS = {
    "mc": os.path.join(USERPROFILE, ".claude-mc"),
    "yk": os.path.join(USERPROFILE, ".claude"),
    "xh": os.path.join(USERPROFILE, ".claude-xh"),
}

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
TARGET_DIR = os.path.join(REPO_ROOT, ".claude", "auto-memory")
BACKUP_BASE = os.path.join(TARGET_DIR, ".pre-junction-backup")


def detect_project_name():
    name = REPO_ROOT.replace(":", "-").replace("\\", "-").replace("/", "-").replace(" ", "-")
    return name.rstrip("-")


def is_junction(path):
    if not os.path.isdir(path):
        return False
    attrs = ctypes.windll.kernel32.GetFileAttributesW(path)
    if attrs == 0xFFFFFFFF:
        return False
    return bool(attrs & stat.FILE_ATTRIBUTE_REPARSE_POINT)


def find_latest_backup(account_key):
    if not os.path.isdir(BACKUP_BASE):
        return None
    pattern = re.compile(rf"^{re.escape(account_key)}-(\d+)$")
    candidates = []
    for name in os.listdir(BACKUP_BASE):
        m = pattern.match(name)
        if m:
            candidates.append((int(m.group(1)), os.path.join(BACKUP_BASE, name)))
    if not candidates:
        return None
    return max(candidates, key=lambda x: x[0])[1]


def rollback_one(account_key, project_name, dry_run):
    src = os.path.join(ACCOUNTS[account_key], "projects", project_name, "memory")
    print(f"\n--- [{account_key}] {src}")

    if not os.path.exists(src):
        print(f"  [skip] 路径不存在")
        return True

    if not is_junction(src):
        print(f"  [skip] 不是 junction（可能未 setup 或已 rollback）")
        return True

    print(f"  [1/2] os.rmdir 删除 junction（不动 D 盘本体）")
    if not dry_run:
        os.rmdir(src)

    backup_path = find_latest_backup(account_key)
    if backup_path:
        print(f"  [2/2] 从备份恢复：{backup_path}")
        if not dry_run:
            shutil.copytree(backup_path, src)
        print(f"  [OK  ] 已恢复")
    else:
        print(f"  [warn] 没找到备份，账号 memory 目录现在是空的（你可以手工再跑 union merge）")
    return True


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--apply", action="store_true")
    parser.add_argument("--account", choices=["mc", "yk", "xh"], help="只回滚指定账号")
    args = parser.parse_args()

    project_name = detect_project_name()
    print(f"项目名：{project_name}")
    print(f"备份目录：{BACKUP_BASE}")
    print(f"\n模式：{'APPLY (实际执行)' if args.apply else 'DRY-RUN'}")
    if not args.apply:
        print(f"(dry-run，不会修改任何文件；加 --apply 实际执行)")

    accounts = [args.account] if args.account else ["mc", "yk", "xh"]
    for account_key in accounts:
        try:
            rollback_one(account_key, project_name, dry_run=not args.apply)
        except Exception as e:
            print(f"  [EXCEPT] {e}")
            return 1

    print(f"\n完成（{len(accounts)} 个账号）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
