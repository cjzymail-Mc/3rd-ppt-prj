"""3 账号 auto-memory 目录 junction 化（一次性架构修复）

前置条件：先跑 memory_union_merge.py --apply，确保 .claude/auto-memory/ 已有 union 后的文件。

执行流程（每个账号）：
1. 备份原 memory/ 到 .claude/auto-memory/.pre-junction-backup/<account>-<num>/
2. shutil.rmtree 原 memory/ 目录（此时它还是真实目录，不是 junction，安全）
3. mklink /J <account_path> <repo_target> 建立 NTFS 目录联接
4. 用 reparse point 属性验证 junction 生效

使用：
    python skills/memory_junction_setup.py              # dry-run
    python skills/memory_junction_setup.py --apply      # 实际执行
"""
import argparse
import ctypes
import os
import re
import shutil
import stat
import subprocess
import sys
import time

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


def next_backup_number(account_key):
    if not os.path.isdir(BACKUP_BASE):
        return 1
    nums = []
    pattern = re.compile(rf"^{re.escape(account_key)}-(\d+)$")
    for name in os.listdir(BACKUP_BASE):
        m = pattern.match(name)
        if m:
            nums.append(int(m.group(1)))
    return max(nums) + 1 if nums else 1


def account_memory_path(account_key, project_name):
    return os.path.join(ACCOUNTS[account_key], "projects", project_name, "memory")


def validate_target():
    if not os.path.isdir(TARGET_DIR):
        return False, f"目标目录不存在：{TARGET_DIR}（请先跑 memory_union_merge.py --apply）"
    md_files = [f for f in os.listdir(TARGET_DIR) if f.endswith(".md")]
    if not md_files:
        return False, f"目标目录为空：{TARGET_DIR}（请先跑 memory_union_merge.py --apply）"
    return True, f"目标目录有 {len(md_files)} 个 .md 文件，OK"


def setup_one(account_key, project_name, dry_run):
    src = account_memory_path(account_key, project_name)
    print(f"\n--- [{account_key}] {src}")

    if not os.path.isdir(src):
        print(f"  [skip] 源目录不存在，跳过")
        return True

    if is_junction(src):
        print(f"  [skip] 已经是 junction，跳过")
        return True

    num = next_backup_number(account_key)
    backup_path = os.path.join(BACKUP_BASE, f"{account_key}-{num:03d}")
    print(f"  [1/3] 备份 → {backup_path}")
    if not dry_run:
        os.makedirs(BACKUP_BASE, exist_ok=True)
        shutil.copytree(src, backup_path)

    print(f"  [2/3] rmtree 原目录")
    if not dry_run:
        shutil.rmtree(src)

    print(f'  [3/3] mklink /J "{src}" "{TARGET_DIR}"')
    if not dry_run:
        result = subprocess.run(
            ["cmd", "/c", "mklink", "/J", src, TARGET_DIR],
            capture_output=True, text=True
        )
        if result.returncode != 0:
            print(f"  [FAIL] mklink 失败：{result.stderr}")
            return False
        if not is_junction(src):
            print(f"  [FAIL] junction 验证失败")
            return False
        print(f"  [OK  ] junction 创建并验证")
    return True


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args()

    project_name = detect_project_name()
    print(f"项目名：{project_name}")
    print(f"目标物理目录：{TARGET_DIR}")

    ok, msg = validate_target()
    print(f"目标校验：{msg}")
    if not ok:
        return 1

    print(f"\n模式：{'APPLY (实际执行)' if args.apply else 'DRY-RUN'}")
    if not args.apply:
        print(f"(dry-run，不会修改任何文件；加 --apply 实际执行)")

    failures = []
    for account_key in ["mc", "yk", "xh"]:
        try:
            ok = setup_one(account_key, project_name, dry_run=not args.apply)
            if not ok:
                failures.append(account_key)
        except Exception as e:
            print(f"  [EXCEPT] {e}")
            failures.append(account_key)

    print()
    if failures:
        print(f"完成（部分失败）：{failures}。请检查上述日志，必要时跑 memory_junction_rollback.py")
        return 1
    print(f"完成（全部成功）：3 个账号已 junction 到 {TARGET_DIR}")
    if args.apply:
        print(f"\n下一步：跑 python skills/memory_junction_verify.py 端到端验证")
    return 0


if __name__ == "__main__":
    sys.exit(main())
