"""3 账号 auto-memory junction 端到端验证

测试流程：
1. 在 mc 账号 memory/ 写一个 _test_<timestamp>.md 时间戳文件
2. 验证 yk、xh、repo 物理目录三处都能立即看到（同一物理文件）
3. 删除该文件
4. 验证四处同步消失

也兼具检测函数：报告 3 账号 memory 目录的 reparse point 状态。

使用：
    python skills/memory_junction_verify.py
"""
import ctypes
import os
import stat
import sys
import time

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

USERPROFILE = os.path.expandvars("%USERPROFILE%")

ACCOUNTS = {
    "mc": os.path.join(USERPROFILE, ".claude-mc"),
    "yk": os.path.join(USERPROFILE, ".claude"),
    "xh": os.path.join(USERPROFILE, ".claude-xh"),
}

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
TARGET_DIR = os.path.join(REPO_ROOT, ".claude", "auto-memory")


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


def memory_path(account_key, project_name):
    return os.path.join(ACCOUNTS[account_key], "projects", project_name, "memory")


def report_status(project_name):
    print("【账号 memory 目录 junction 状态】")
    for k in ["mc", "yk", "xh"]:
        p = memory_path(k, project_name)
        if not os.path.exists(p):
            status = "不存在"
        elif is_junction(p):
            status = "✓ junction"
        else:
            status = "✗ 真实目录（未 junction）"
        print(f"  {k:<3} {status:<22} {p}")
    print()


def run_propagation_test(project_name):
    print("【写入传播测试】")
    test_name = f"_test_{int(time.time())}.md"
    mc_path = os.path.join(memory_path("mc", project_name), test_name)
    yk_path = os.path.join(memory_path("yk", project_name), test_name)
    xh_path = os.path.join(memory_path("xh", project_name), test_name)
    repo_path = os.path.join(TARGET_DIR, test_name)

    test_content = f"junction-verify-{time.strftime('%Y%m%d-%H%M%S')}\n"

    print(f"  [1/4] 在 mc 账号写入 {test_name}")
    try:
        with open(mc_path, "w", encoding="utf-8") as f:
            f.write(test_content)
    except Exception as e:
        print(f"  [FAIL] 写入失败：{e}")
        return False

    print(f"  [2/4] 验证 yk / xh / repo 三处同步可见")
    failures = []
    for label, p in [("yk", yk_path), ("xh", xh_path), ("repo", repo_path)]:
        if os.path.isfile(p):
            with open(p, "r", encoding="utf-8") as f:
                content = f.read()
            if content == test_content:
                print(f"        ✓ {label} 可见且内容一致")
            else:
                print(f"        ✗ {label} 可见但内容不同（不是同一物理文件！）")
                failures.append(label)
        else:
            print(f"        ✗ {label} 看不到该文件")
            failures.append(label)

    print(f"  [3/4] 从 mc 账号删除")
    try:
        os.remove(mc_path)
    except Exception as e:
        print(f"  [FAIL] 删除失败：{e}")
        return False

    print(f"  [4/4] 验证四处同步消失")
    for label, p in [("mc", mc_path), ("yk", yk_path), ("xh", xh_path), ("repo", repo_path)]:
        if os.path.isfile(p):
            print(f"        ✗ {label} 还存在（同步失败！）")
            failures.append(f"{label}-after-delete")
        else:
            print(f"        ✓ {label} 已消失")

    return not failures


def main():
    project_name = detect_project_name()
    print(f"项目名：{project_name}")
    print(f"REPO_ROOT：{REPO_ROOT}")
    print(f"物理目录：{TARGET_DIR}")
    print()

    report_status(project_name)

    junctions = sum(1 for k in ACCOUNTS if is_junction(memory_path(k, project_name)))
    if junctions < 3:
        print(f"[abort] 只有 {junctions}/3 账号是 junction，先跑 memory_junction_setup.py --apply")
        return 1

    print()
    ok = run_propagation_test(project_name)
    print()
    if ok:
        print("✓✓✓ 端到端验证通过：3 账号 memory 物理合一，写入传播无延迟")
        return 0
    else:
        print("✗✗✗ 端到端验证失败，详见上方日志")
        return 1


if __name__ == "__main__":
    sys.exit(main())
