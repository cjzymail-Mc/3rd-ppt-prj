import os
import re
import shutil
import time
import tkinter as tk
from tkinter import ttk, messagebox


# ============================================================
# 账户配置
# ============================================================
USERPROFILE = os.path.expandvars("%USERPROFILE%")

ACCOUNTS = {
    "mc": {
        "label": "账户1：mc",
        "dir": os.path.join(USERPROFILE, ".claude-mc"),
    },
    "yk": {
        "label": "账户2：yk",
        "dir": os.path.join(USERPROFILE, ".claude"),
    },
    "xh": {
        "label": "账户3：xh",
        "dir": os.path.join(USERPROFILE, ".claude-xh"),
    },
}


# ============================================================
# 工具函数
# ============================================================
def detect_project_name():
    """根据脚本所在目录自动推断 Claude 项目文件夹名称"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    name = script_dir.replace(":", "-").replace("\\", "-").replace("/", "-").replace(" ", "-")
    return name.rstrip("-")


def project_exists(account_key, project_name):
    """检查项目文件夹在指定账户中是否存在"""
    path = os.path.join(ACCOUNTS[account_key]["dir"], "projects", project_name)
    return os.path.isdir(path)


def get_project_mtime(account_key, project_name):
    """获取项目文件夹中最新 .jsonl 会话文件的修改时间，返回 (时间戳, 格式化字符串)；不存在则返回 (0, None)"""
    project_dir = os.path.join(ACCOUNTS[account_key]["dir"], "projects", project_name)
    if not os.path.isdir(project_dir):
        return 0, None
    latest_ts = 0
    for f in os.listdir(project_dir):
        if f.endswith(".jsonl"):
            ts = os.path.getmtime(os.path.join(project_dir, f))
            if ts > latest_ts:
                latest_ts = ts
    if latest_ts == 0:
        return 0, None
    return latest_ts, time.strftime("%Y/%m/%d %H:%M", time.localtime(latest_ts))


def next_backup_number(source_dir):
    """扫描 back-up 目录，返回下一个可用的备份编号"""
    backup_base = os.path.join(source_dir, "back-up")
    if not os.path.isdir(backup_base):
        return 1
    nums = []
    for name in os.listdir(backup_base):
        m = re.match(r"backup#(\d+)$", name)
        if m:
            nums.append(int(m.group(1)))
    return max(nums) + 1 if nums else 1


def run_migration(source_key, target_key, project_name):
    """
    执行迁移：备份源 → 清理目标 → 复制到目标
    返回 (success: bool, message: str)
    """
    src_dir = ACCOUNTS[source_key]["dir"]
    tgt_dir = ACCOUNTS[target_key]["dir"]

    src_project = os.path.join(src_dir, "projects", project_name)
    src_history = os.path.join(src_dir, "history.jsonl")
    src_settings = os.path.join(src_dir, "settings.json")

    # 前置校验
    if not os.path.isdir(src_project):
        return False, f"源账户中不存在该项目文件夹：\n{src_project}"

    # ---------- 第1步：确定备份路径 ----------
    num = next_backup_number(src_dir)
    backup_dir = os.path.join(src_dir, "back-up", f"backup#{num:03d}")

    try:
        # ---------- 第2步：备份源账户数据 ----------
        os.makedirs(backup_dir, exist_ok=True)
        shutil.copytree(src_project, os.path.join(backup_dir, project_name))
        if os.path.isfile(src_history):
            shutil.copy2(src_history, os.path.join(backup_dir, "history.jsonl"))
        if os.path.isfile(src_settings):
            shutil.copy2(src_settings, os.path.join(backup_dir, "settings.json"))

        # ---------- 第3步：清理目标账户旧数据 ----------
        tgt_project = os.path.join(tgt_dir, "projects", project_name)
        tgt_history = os.path.join(tgt_dir, "history.jsonl")
        tgt_settings = os.path.join(tgt_dir, "settings.json")

        if os.path.isdir(tgt_project):
            shutil.rmtree(tgt_project)
        if os.path.isfile(tgt_history):
            os.remove(tgt_history)
        if os.path.isfile(tgt_settings):
            os.remove(tgt_settings)

        # ---------- 第4步：复制到目标账户 ----------
        os.makedirs(os.path.join(tgt_dir, "projects"), exist_ok=True)
        shutil.copytree(src_project, tgt_project)
        if os.path.isfile(src_history):
            shutil.copy2(src_history, tgt_history)
        if os.path.isfile(src_settings):
            shutil.copy2(src_settings, tgt_settings)

        return True, f"备份已保存至：\n{backup_dir}"

    except Exception as e:
        return False, f"迁移过程中出错：\n{e}"


# ============================================================
# GUI
# ============================================================
class MigrateApp:

    def __init__(self):
        self.project_name = detect_project_name()

        self.root = tk.Tk()
        self.root.title("Claude 账户会话迁移工具")
        self.root.minsize(700, 0)
        self.root.resizable(False, False)
        self._build_main()
        self.root.mainloop()

    # ----------------------------------------------------------
    # 第一个窗口：主界面
    # ----------------------------------------------------------
    def _build_main(self):
        for w in self.root.winfo_children():
            w.destroy()

        pad_x = 25
        font_title = ("Microsoft YaHei", 15, "bold")
        font_normal = ("Microsoft YaHei", 9)
        font_bold = ("Microsoft YaHei", 9, "bold")
        font_btn = ("Microsoft YaHei", 10)

        # 1. 标题
        tk.Label(self.root, text="Claude 账户会话迁移工具", font=font_title).pack(
            pady=(18, 6)
        )

        # 2. 当前项目
        f_proj = tk.Frame(self.root)
        f_proj.pack(fill="x", padx=pad_x)
        tk.Label(f_proj, text="当前项目：", font=font_normal).pack(side="left")
        tk.Label(
            f_proj, text=self.project_name, font=font_bold, fg="#1565C0"
        ).pack(side="left")

        # 3-4. 项目状态 + 最新会话（表格式对齐）
        keys = ["mc", "yk", "xh"]
        col_width = 18  # 每列固定字符宽度

        # 收集数据
        exists_map = {k: project_exists(k, self.project_name) for k in keys}
        mtime_ts_map = {}
        mtime_str_map = {}
        for k in keys:
            ts, s = get_project_mtime(k, self.project_name)
            mtime_ts_map[k] = ts
            mtime_str_map[k] = s
        # 找出最新时间的账户
        latest_key = None
        latest_ts = 0
        for k in keys:
            if mtime_ts_map[k] > latest_ts:
                latest_ts = mtime_ts_map[k]
                latest_key = k

        # 第一行：项目状态
        f_status = tk.Frame(self.root)
        f_status.pack(fill="x", padx=pad_x, pady=(6, 0))
        tk.Label(f_status, text="项目状态：", font=font_normal, width=9, anchor="e").pack(side="left")
        for i, key in enumerate(keys):
            if i > 0:
                tk.Label(f_status, text="|", font=font_normal, fg="gray").pack(side="left")
            mark = "✓ 存在" if exists_map[key] else "✗ 不存在"
            color = "#2E7D32" if exists_map[key] else "#C62828"
            tk.Label(
                f_status, text=f"{key}: {mark}", font=font_normal, fg=color,
                width=col_width, anchor="center"
            ).pack(side="left")

        # 第二行：最新会话时间
        f_time = tk.Frame(self.root)
        f_time.pack(fill="x", padx=pad_x, pady=(2, 0))
        tk.Label(f_time, text="最新会话：", font=font_normal, width=9, anchor="e").pack(side="left")
        for i, key in enumerate(keys):
            if i > 0:
                tk.Label(f_time, text="|", font=font_normal, fg="gray").pack(side="left")
            display = mtime_str_map[key] if mtime_str_map[key] else "-"
            is_latest = (key == latest_key)
            lbl = tk.Label(
                f_time, text=f"{key}: {display}",
                font=font_bold if is_latest else font_normal,
                fg="#333" if is_latest else "#555",
                bg="#FFEB3B" if is_latest else self.root.cget("bg"),
                width=col_width, anchor="center"
            )
            lbl.pack(side="left")

        # 5. 账户路径信息
        f_paths = tk.Frame(self.root)
        f_paths.pack(fill="x", padx=pad_x, pady=(6, 0))
        tk.Label(f_paths, text="账户路径：", font=font_bold).pack(anchor="w")
        for key in ["mc", "yk", "xh"]:
            info = ACCOUNTS[key]
            tk.Label(
                f_paths,
                text=f"  {info['label']}  ——  {info['dir']}",
                font=font_normal,
                fg="#555",
            ).pack(anchor="w")

        # 5. 分隔线
        ttk.Separator(self.root, orient="horizontal").pack(
            fill="x", padx=pad_x, pady=(12, 8)
        )

        # 6. 迁移按钮（6 个，分 3 组）
        tk.Label(self.root, text="选择迁移方向：", font=font_bold).pack(
            padx=pad_x, anchor="w"
        )

        groups = [
            [("mc", "yk"), ("mc", "xh")],
            [("yk", "mc"), ("yk", "xh")],
            [("xh", "mc"), ("xh", "yk")],
        ]

        f_btns = tk.Frame(self.root)
        f_btns.pack(fill="x", padx=pad_x, pady=(4, 15))

        for g_idx, group in enumerate(groups):
            if g_idx > 0:
                ttk.Separator(f_btns, orient="horizontal").pack(fill="x", pady=4)
            row = tk.Frame(f_btns)
            row.pack(fill="x", pady=2)
            for src, tgt in group:
                text = f"【{ACCOUNTS[src]['label']}】→【{ACCOUNTS[tgt]['label']}】"
                btn = tk.Button(
                    row,
                    text=text,
                    font=font_btn,
                    command=lambda s=src, t=tgt: self._show_confirm(s, t),
                )
                btn.pack(side="left", expand=True, fill="x", padx=4)

    # ----------------------------------------------------------
    # 第二个窗口：确认对话框
    # ----------------------------------------------------------
    def _show_confirm(self, source_key, target_key):
        src = ACCOUNTS[source_key]
        tgt = ACCOUNTS[target_key]
        num = next_backup_number(src["dir"])
        backup_path = os.path.join(src["dir"], "back-up", f"backup#{num:03d}")

        dlg = tk.Toplevel(self.root)
        dlg.title("确认迁移操作")
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.focus_set()

        font_normal = ("Microsoft YaHei", 9)
        font_bold = ("Microsoft YaHei", 9, "bold")
        font_btn = ("Microsoft YaHei", 10)
        pad_x = 25

        # 迁移方向标题
        direction = f"【{src['label']}】 → 【{tgt['label']}】"
        tk.Label(
            dlg, text=direction, font=("Microsoft YaHei", 13, "bold"), fg="#1565C0"
        ).pack(pady=(15, 10))

        ttk.Separator(dlg, orient="horizontal").pack(fill="x", padx=pad_x)

        # 详细信息
        f_info = tk.Frame(dlg)
        f_info.pack(fill="x", padx=pad_x, pady=10)

        rows = [
            ("源账户：", src["label"]),
            ("源路径：", src["dir"]),
            ("", ""),
            ("目标账户：", tgt["label"]),
            ("目标路径：", tgt["dir"]),
            ("", ""),
            ("项目文件夹：", self.project_name),
            ("备份位置：", backup_path),
            ("", ""),
            ("迁移文件：", "① projects/<项目文件夹>"),
            ("", "② history.jsonl"),
            ("", "③ settings.json"),
        ]

        for label, value in rows:
            if not label and not value:
                tk.Frame(f_info, height=4).pack()
                continue
            r = tk.Frame(f_info)
            r.pack(fill="x", pady=1)
            tk.Label(r, text=label, font=font_bold, width=11, anchor="e").pack(
                side="left"
            )
            tk.Label(r, text=value, font=font_normal, anchor="w").pack(side="left")

        ttk.Separator(dlg, orient="horizontal").pack(fill="x", padx=pad_x, pady=(5, 0))

        # 按钮
        f_btn = tk.Frame(dlg)
        f_btn.pack(pady=15)

        def on_confirm():
            dlg.destroy()
            ok, msg = run_migration(source_key, target_key, self.project_name)
            if ok:
                messagebox.showinfo(
                    "迁移完成",
                    f"迁移成功！\n\n"
                    f"{direction}\n\n"
                    f"{msg}\n\n"
                    f"请在目标账户中执行：\nclaude --resume",
                )
            else:
                messagebox.showerror("迁移失败", msg)
            # 刷新主界面状态
            self._build_main()

        tk.Button(
            f_btn,
            text="确认执行",
            font=font_btn,
            bg="#4CAF50",
            fg="white",
            activebackground="#388E3C",
            activeforeground="white",
            width=12,
            command=on_confirm,
        ).pack(side="left", padx=10)

        tk.Button(
            f_btn,
            text="返回",
            font=font_btn,
            width=12,
            command=dlg.destroy,
        ).pack(side="left", padx=10)


# ============================================================
# 入口
# ============================================================
if __name__ == "__main__":
    MigrateApp()
