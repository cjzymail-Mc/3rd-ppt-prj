"""3 账号 auto-memory union merge 工具

扫描 mc / yk / xh 三个账号下 projects/<proj>/memory/ 目录，求 union 写入 repo 内
.claude/auto-memory/。MEMORY.md 按文件名 key 合并 bullet 行；其他文件按文件名 union，
同名取最新 mtime。

使用：
    python tools/memory_union_merge.py              # dry-run，输出三栏决策表
    python tools/memory_union_merge.py --apply      # 执行 merge
"""
import argparse
import os
import re
import shutil
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


def detect_project_name():
    name = REPO_ROOT.replace(":", "-").replace("\\", "-").replace("/", "-").replace(" ", "-")
    return name.rstrip("-")


def list_memory_files(account_key, project_name):
    memory_dir = os.path.join(ACCOUNTS[account_key], "projects", project_name, "memory")
    if not os.path.isdir(memory_dir):
        return {}
    out = {}
    for fname in os.listdir(memory_dir):
        full = os.path.join(memory_dir, fname)
        if os.path.isfile(full) and fname.endswith(".md"):
            out[fname] = (full, os.path.getmtime(full))
    return out


def parse_memory_index_bullets(path):
    """解析 MEMORY.md 的 - [Title](file.md) — desc 行，返回 {file_key: original_line}。"""
    bullets = {}
    header = []
    with open(path, "r", encoding="utf-8") as f:
        lines = f.read().splitlines()
    for line in lines:
        m = re.match(r"^\s*-\s*\[[^\]]+\]\(([^)]+)\)", line)
        if m:
            bullets[m.group(1)] = line
        elif not bullets:
            header.append(line)
    return header, bullets


def build_decision_table(project_name):
    """返回 [(filename, source_account, decision_text, content_path)]。"""
    per_account = {k: list_memory_files(k, project_name) for k in ACCOUNTS}
    all_filenames = set()
    for files in per_account.values():
        all_filenames.update(files.keys())

    decisions = []
    for fname in sorted(all_filenames):
        if fname == "MEMORY.md":
            sources = [k for k in ACCOUNTS if fname in per_account[k]]
            decisions.append((fname, "+".join(sources), f"merge MEMORY.md from {len(sources)} accounts (union of bullets)", None))
            continue
        accounts_with = [(k, per_account[k][fname][1]) for k in ACCOUNTS if fname in per_account[k]]
        if len(accounts_with) == 1:
            k, _ = accounts_with[0]
            decisions.append((fname, k, f"new from {k}-only", per_account[k][fname][0]))
        else:
            best = max(accounts_with, key=lambda x: x[1])
            mtimes = [time.strftime("%Y-%m-%d", time.localtime(m)) for _, m in accounts_with]
            same_mtime = len(set(mtimes)) == 1
            if same_mtime:
                k, _ = best
                decisions.append((fname, "+".join([k for k, _ in accounts_with]), f"all identical mtime ({mtimes[0]}), pick from {k}", per_account[k][fname][0]))
            else:
                k, _ = best
                decisions.append((fname, "+".join([k for k, _ in accounts_with]), f"newest from {k} ({time.strftime('%Y-%m-%d', time.localtime(best[1]))})", per_account[k][fname][0]))
    return decisions, per_account


def render_dry_run(decisions):
    print(f"{'文件名':<40} {'来源':<12} {'决策'}")
    print("-" * 100)
    for fname, source, decision, _ in decisions:
        print(f"{fname:<40} {source:<12} {decision}")
    print("-" * 100)
    print(f"共 {len(decisions)} 个文件 → {TARGET_DIR}")


def do_apply(decisions, per_account):
    if os.path.isdir(TARGET_DIR):
        existing = [f for f in os.listdir(TARGET_DIR) if f.endswith(".md")]
        if existing:
            print(f"[abort] 目标目录已有文件：{TARGET_DIR}")
            print(f"        请先清空（或备份后清空），再跑 --apply。")
            return 1
    os.makedirs(TARGET_DIR, exist_ok=True)

    for fname, source, decision, content_path in decisions:
        target_path = os.path.join(TARGET_DIR, fname)
        if fname == "MEMORY.md":
            merged_header, merged_bullets = None, {}
            for k in ACCOUNTS:
                if fname in per_account[k]:
                    p = per_account[k][fname][0]
                    h, b = parse_memory_index_bullets(p)
                    if merged_header is None:
                        merged_header = h
                    for key, line in b.items():
                        merged_bullets[key] = line
            with open(target_path, "w", encoding="utf-8") as f:
                f.write("\n".join(merged_header).rstrip() + "\n")
                if merged_bullets:
                    f.write("\n")
                    for key in sorted(merged_bullets.keys()):
                        f.write(merged_bullets[key] + "\n")
            print(f"[merge ] {fname}  ←  {len(merged_bullets)} bullets union")
        else:
            shutil.copy2(content_path, target_path)
            print(f"[copy  ] {fname}  ←  {source}")
    print(f"\n完成：{len(decisions)} 个文件已写入 {TARGET_DIR}")
    return 0


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--apply", action="store_true", help="执行 merge（默认 dry-run）")
    args = parser.parse_args()

    project_name = detect_project_name()
    print(f"项目名（自动探测）：{project_name}")
    print(f"3 账号根目录：")
    for k, root in ACCOUNTS.items():
        print(f"  {k}: {root}")
    print(f"目标目录：{TARGET_DIR}")
    print()

    decisions, per_account = build_decision_table(project_name)
    if not decisions:
        print("[warn] 3 个账号 memory 目录都为空或不存在，没有可 merge 的内容。")
        return 1

    render_dry_run(decisions)
    if args.apply:
        print()
        return do_apply(decisions, per_account)
    print()
    print("(dry-run，未写入任何文件。加 --apply 实际执行。)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
