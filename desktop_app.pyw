from __future__ import annotations

import os
import shutil
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
SCRIPTS = ROOT / "scripts"
CONFIG = ROOT / "config"
DATA = ROOT / "data" / "每日任务"
IMPORTS = ROOT / "imports"
LOGS = ROOT / "logs"


def python_exe() -> str:
    candidate = Path(sys.executable)
    if candidate.name.lower() == "pythonw.exe":
        candidate = candidate.with_name("python.exe")
    if candidate.exists():
        return str(candidate)
    return shutil.which("python.exe") or "python.exe"


def powershell_exe() -> str:
    return r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CyberPlan Boot Mailer - 每日计划与邮件")
        self.geometry("980x720")
        self.minsize(860, 620)
        self.today = datetime.now().strftime("%Y-%m-%d")
        self.task_dir = DATA / self.today
        self.status_var = tk.StringVar(value="就绪")
        self.auto_status_var = tk.StringVar()
        self.email_status_var = tk.StringVar()
        self.plan_preview = None
        self.external_text = None
        self._build_style()
        self._build_ui()
        self.refresh_status()
        self.load_plan()

    def _build_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("微软雅黑", 18, "bold"), foreground="#1F4E79")
        style.configure("Subtitle.TLabel", font=("微软雅黑", 10), foreground="#666666")
        style.configure("Section.TLabel", font=("微软雅黑", 12, "bold"), foreground="#1F4E79")
        style.configure("Action.TButton", font=("微软雅黑", 10), padding=(12, 7))
        style.configure("TNotebook.Tab", font=("微软雅黑", 10), padding=(14, 8))
        style.configure("TLabel", font=("微软雅黑", 10))
        self.configure(bg="#F4F7FA")

    def _build_ui(self):
        header = ttk.Frame(self, padding=(22, 18, 22, 8))
        header.pack(fill="x")
        ttk.Label(header, text="CyberPlan Boot Mailer", style="Title.TLabel").pack(anchor="w")
        ttk.Label(header, text="开机自动生成计划、发送邮件、导入其他 AI 的任务安排", style="Subtitle.TLabel").pack(anchor="w", pady=(4, 0))

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=18, pady=8)
        self.today_tab = ttk.Frame(self.notebook, padding=16)
        self.external_tab = ttk.Frame(self.notebook, padding=16)
        self.auto_tab = ttk.Frame(self.notebook, padding=16)
        self.tutorial_tab = ttk.Frame(self.notebook, padding=16)
        self.notebook.add(self.today_tab, text="今日计划")
        self.notebook.add(self.external_tab, text="外部 AI 计划")
        self.notebook.add(self.auto_tab, text="自动运行与邮件")
        self.notebook.add(self.tutorial_tab, text="使用教程")
        self._build_today_tab()
        self._build_external_tab()
        self._build_auto_tab()
        self._build_tutorial_tab()

        status = ttk.Frame(self, padding=(18, 4, 18, 12))
        status.pack(fill="x")
        ttk.Label(status, textvariable=self.status_var, foreground="#555555").pack(anchor="w")

    def _button_row(self, parent):
        return ttk.Frame(parent)

    def _build_today_tab(self):
        top = ttk.Frame(self.today_tab)
        top.pack(fill="x")
        ttk.Label(top, text=f"今天：{self.today}", style="Section.TLabel").pack(side="left")
        ttk.Button(top, text="刷新", style="Action.TButton", command=self.load_plan).pack(side="right")

        actions = ttk.Frame(self.today_tab)
        actions.pack(fill="x", pady=12)
        for text, command in [
            ("生成/刷新今日计划", self.generate_today),
            ("立即发送今日邮件", self.send_today_email),
            ("测试桌面通知", self.test_popup),
            ("打开 Word", self.open_today_word),
            ("打开 Markdown", self.open_today_markdown),
            ("打开任务文件夹", lambda: self.open_path(self.task_dir)),
        ]:
            ttk.Button(actions, text=text, style="Action.TButton", command=command).pack(side="left", padx=(0, 8), pady=4)

        ttk.Label(self.today_tab, text="今日计划预览", style="Section.TLabel").pack(anchor="w", pady=(8, 4))
        self.plan_preview = tk.Text(self.today_tab, wrap="word", font=("微软雅黑", 10), bg="white", relief="solid", borderwidth=1)
        self.plan_preview.pack(fill="both", expand=True)

    def _build_external_tab(self):
        ttk.Label(self.external_tab, text="把其他 AI 生成的计划放进每天任务", style="Section.TLabel").pack(anchor="w")
        info = (
            "支持两种方式：\n"
            "1. 把其他 AI 的内容粘贴到下面，点击“保存并合并”。\n"
            "2. 让其他 AI 直接把 Markdown 或文本保存到 imports\\YYYY-MM-DD.md，开机生成时会自动合并。\n"
            "合并后的内容会加入今日 Markdown 和 Word，并随邮件发送。"
        )
        ttk.Label(self.external_tab, text=info, justify="left", foreground="#555555").pack(anchor="w", pady=(6, 10))

        actions = ttk.Frame(self.external_tab)
        actions.pack(fill="x", pady=(0, 8))
        for text, command in [
            ("从剪贴板导入", self.paste_external),
            ("选择 Markdown/TXT 文件", self.choose_external_file),
            ("保存并合并", self.merge_external),
            ("刷新 imports 并合并", self.refresh_external),
            ("打开 imports 文件夹", lambda: self.open_path(IMPORTS)),
            ("清空编辑区", self.clear_external),
        ]:
            ttk.Button(actions, text=text, style="Action.TButton", command=command).pack(side="left", padx=(0, 8), pady=4)

        self.external_text = tk.Text(self.external_tab, wrap="word", font=("微软雅黑", 10), bg="white", relief="solid", borderwidth=1)
        self.external_text.pack(fill="both", expand=True)
        self.external_text.insert("1.0", self._read_external())

    def _build_auto_tab(self):
        ttk.Label(self.auto_tab, text="自动运行与邮件状态", style="Section.TLabel").pack(anchor="w")
        ttk.Label(self.auto_tab, textvariable=self.auto_status_var, foreground="#555555").pack(anchor="w", pady=(8, 2))
        ttk.Label(self.auto_tab, textvariable=self.email_status_var, foreground="#555555").pack(anchor="w", pady=(2, 14))

        row1 = ttk.Frame(self.auto_tab)
        row1.pack(fill="x", pady=5)
        ttk.Button(row1, text="安装/修复开机自启", style="Action.TButton", command=self.install_autostart).pack(side="left", padx=(0, 8))
        ttk.Button(row1, text="卸载开机自启", style="Action.TButton", command=self.uninstall_autostart).pack(side="left", padx=(0, 8))
        ttk.Button(row1, text="配置邮件", style="Action.TButton", command=self.configure_email).pack(side="left")

        row2 = ttk.Frame(self.auto_tab)
        row2.pack(fill="x", pady=16)
        ttk.Button(row2, text="打开日志文件夹", style="Action.TButton", command=lambda: self.open_path(LOGS)).pack(side="left", padx=(0, 8))
        ttk.Button(row2, text="打开项目目录", style="Action.TButton", command=lambda: self.open_path(ROOT)).pack(side="left", padx=(0, 8))
        ttk.Button(row2, text="打开邮件日志", style="Action.TButton", command=lambda: self.open_file(LOGS / "email.log")).pack(side="left")

        ttk.Separator(self.auto_tab).pack(fill="x", pady=15)
        ttk.Label(self.auto_tab, text="运行规则", style="Section.TLabel").pack(anchor="w")
        ttk.Label(
            self.auto_tab,
            text="开机并登录 Windows 后 20 秒：生成今日任务、发送邮件、弹出桌面通知。\n"
                 "周二至周日 21:00：发送复盘提醒邮件并弹出通知。\n"
                 "周一：只提示休息，不发送任务邮件。\n"
                 "Codex 关闭时仍会运行。",
            justify="left",
            foreground="#555555",
        ).pack(anchor="w", pady=(6, 0))

    def _build_tutorial_tab(self):
        text = tk.Text(self.tutorial_tab, wrap="word", font=("微软雅黑", 10), bg="white", relief="solid", borderwidth=1)
        text.pack(fill="both", expand=True)
        text.insert("1.0", self.tutorial_content())
        text.configure(state="disabled")

    def _read_external(self) -> str:
        IMPORTS.mkdir(parents=True, exist_ok=True)
        for suffix in (".md", ".txt"):
            path = IMPORTS / f"{self.today}{suffix}"
            if path.exists():
                return path.read_text(encoding="utf-8")
        return ""

    def load_plan(self):
        self.task_dir.mkdir(parents=True, exist_ok=True)
        md = self.task_dir / f"{self.today}_今日任务.md"
        if md.exists() and self.plan_preview is not None:
            self.plan_preview.delete("1.0", "end")
            self.plan_preview.insert("1.0", md.read_text(encoding="utf-8"))
        self.refresh_status()

    def refresh_status(self):
        self.auto_status_var.set("开机自启：" + ("已安装" if self.task_exists("CyberPlanBootMailer-Logon") else "未安装"))
        self.email_status_var.set("邮件配置：" + ("已配置" if (CONFIG / "email_config.json").exists() and (CONFIG / "email_secret.txt").exists() else "未配置"))

    def task_exists(self, name: str) -> bool:
        try:
            result = subprocess.run(["schtasks.exe", "/Query", "/TN", name], capture_output=True, text=True, timeout=10)
            return result.returncode == 0
        except Exception:
            return False

    def set_status(self, text: str):
        self.after(0, lambda: self.status_var.set(text))

    def run_async(self, label: str, func):
        def worker():
            self.set_status(label)
            try:
                func()
                self.set_status(label + "：完成")
            except Exception as exc:
                self.set_status(label + "：失败")
                self.after(0, lambda: messagebox.showerror("操作失败", str(exc)))
        threading.Thread(target=worker, daemon=True).start()

    def run_process(self, command, timeout=120):
        result = subprocess.run(command, cwd=str(ROOT), capture_output=True, text=True, timeout=timeout)
        if result.returncode != 0:
            detail = (result.stderr or result.stdout or "").strip()
            raise RuntimeError(detail or f"命令失败，退出码 {result.returncode}")
        return result.stdout

    def run_generator(self):
        self.run_process([python_exe(), str(SRC / "generate_daily_task.py"), "--mode", "morning"])

    def run_sender(self, mode="Morning"):
        self.run_process([powershell_exe(), "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(SCRIPTS / "send_email.ps1"), "-Mode", mode, "-Day", self.today], timeout=180)

    def generate_today(self):
        def work():
            self.run_generator()
            self.after(0, self.load_plan)
        self.run_async("生成今日计划", work)

    def send_today_email(self):
        def work():
            self.run_generator()
            self.run_sender("Morning")
        self.run_async("发送今日邮件", work)

    def test_popup(self):
        def work():
            self.run_process([powershell_exe(), "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(SCRIPTS / "entrypoint.ps1"), "-Mode", "Morning", "-Force", "-PopupTimeout", "5"], timeout=60)
        self.run_async("测试桌面通知", work)

    def open_path(self, path: Path):
        path.mkdir(parents=True, exist_ok=True)
        os.startfile(str(path))

    def open_file(self, path: Path):
        if path.exists():
            os.startfile(str(path))
        else:
            messagebox.showinfo("文件不存在", str(path))

    def open_today_word(self):
        base = self.task_dir / f"{self.today}_今日任务.docx"
        pointer = self.task_dir / "current_docx.txt"
        if pointer.exists():
            candidate = self.task_dir / pointer.read_text(encoding="utf-8").strip()
            if candidate.exists():
                base = candidate
        self.open_file(base)

    def open_today_markdown(self):
        self.open_file(self.task_dir / f"{self.today}_今日任务.md")

    def paste_external(self):
        try:
            text = self.clipboard_get()
        except tk.TclError:
            text = ""
        if text:
            self.external_text.delete("1.0", "end")
            self.external_text.insert("1.0", text)
            self.set_status("已从剪贴板导入，点击“保存并合并”生效")

    def choose_external_file(self):
        path = filedialog.askopenfilename(title="选择其他 AI 生成的计划", filetypes=[("Markdown/Text", "*.md *.txt"), ("所有文件", "*.*")])
        if path:
            text = Path(path).read_text(encoding="utf-8", errors="replace")
            self.external_text.delete("1.0", "end")
            self.external_text.insert("1.0", text)

    def clear_external(self):
        self.external_text.delete("1.0", "end")

    def refresh_external(self):
        text = self._read_external()
        if not text:
            messagebox.showinfo("没有外部计划", f"请先把文件放到：{IMPORTS}\\{self.today}.md")
            return
        self.external_text.delete("1.0", "end")
        self.external_text.insert("1.0", text)
        self.merge_external()

    def merge_external(self):
        text = self.external_text.get("1.0", "end").strip()
        if not text:
            messagebox.showwarning("内容为空", "请先粘贴其他 AI 的计划。")
            return
        def work():
            IMPORTS.mkdir(parents=True, exist_ok=True)
            path = IMPORTS / f"{self.today}.md"
            path.write_text(text, encoding="utf-8")
            self.run_generator()
            self.after(0, self.load_plan)
        self.run_async("合并外部 AI 计划", work)

    def install_autostart(self):
        def work():
            self.run_process([powershell_exe(), "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ROOT / "install.ps1")], timeout=180)
            self.after(0, self.refresh_status)
        self.run_async("安装/修复开机自启", work)

    def uninstall_autostart(self):
        if not messagebox.askyesno("确认卸载", "确定要移除开机自启吗？邮件配置和日志会保留。"):
            return
        def work():
            self.run_process([powershell_exe(), "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ROOT / "uninstall.ps1")], timeout=60)
            self.after(0, self.refresh_status)
        self.run_async("卸载开机自启", work)

    def configure_email(self):
        subprocess.Popen([powershell_exe(), "-NoProfile", "-ExecutionPolicy", "Bypass", "-NoExit", "-File", str(SCRIPTS / "configure_email.ps1")], cwd=str(ROOT))

    def tutorial_content(self) -> str:
        return """CyberPlan Boot Mailer 详细教程

一、这个应用做什么
1. 电脑开机并登录 Windows 后，自动生成当天计划。
2. 自动弹出桌面通知，提醒你今天先做什么。
3. 自动把今日任务和 Word 附件发送到你的邮箱。
4. 每天 21:00 自动发送复盘提醒邮件并弹出通知。
5. 支持把其他 AI 生成的计划合并进每天的任务。

二、第一次使用
1. 双击“配置邮件.lnk”。
2. 选择邮箱类型，填写发件邮箱和 SMTP 授权码。
3. 程序会自动发送一封测试邮件。
4. 双击“安装开机自启.lnk”注册 Windows 任务计划。
5. 以后正常开机即可，不需要打开 Codex。

三、每天的开机流程
1. 开机并登录 Windows。
2. 等待约 20 秒。
3. 应用自动生成 data\\每日任务\\日期 文件夹。
4. 同时生成 Markdown 和 Word 文件。
5. 自动发送今日任务邮件。
6. 弹出桌面通知，点击 是 可以直接打开 Word。

四、晚间复盘
1. 周二至周日 21:00 自动弹出复盘提醒。
2. 同时发送复盘提醒邮件。
3. 填写“今日完成、今日产出、最大卡点、明天三件事、得分”。
4. 周一不安排学习任务，不发送任务邮件。

五、导入其他 AI 的计划
方式一：在桌面应用中导入
1. 打开“外部 AI 计划”标签。
2. 粘贴其他 AI 的输出，或点击“从剪贴板导入”。
3. 点击“保存并合并”。
4. 程序会重新生成今日 Markdown 和 Word，并包含外部 AI 补充。
5. 下次发送邮件时会自动带上合并后的内容。

方式二：文件夹自动接收
1. 其他 AI 把计划保存为 imports\\YYYY-MM-DD.md。
2. 例如今天是 2026-09-15，就保存为 imports\\2026-09-15.md。
3. 开机生成任务时，程序会自动读取并合并。
4. 支持 .md 和 .txt 文件。

六、桌面应用按钮说明
- 生成/刷新今日计划：重新生成今天的计划和文件。
- 立即发送今日邮件：生成并发送今天的任务邮件。
- 测试桌面通知：弹出一个短通知，检查桌面提醒是否正常。
- 打开 Word：打开今天带排版的 Word 任务表。
- 打开 Markdown：打开今天可编辑的任务文本。
- 外部 AI 计划：粘贴、导入、合并其他 AI 的输出。
- 自动运行与邮件：安装、修复或卸载开机自启，配置邮箱和查看日志。
- 使用教程：查看这份说明。

七、故障排查
1. 没有生成计划：打开“自动运行与邮件”，点击“安装/修复开机自启”，然后重新登录 Windows。
2. 没有收到邮件：先点“立即发送今日邮件”，再检查垃圾箱和授权码。
3. Word 文件打不开：关闭已经打开的 Word 文件后，再点击“生成/刷新今日计划”。
4. 其他 AI 内容没出现：确认文件在 imports 文件夹，文件名严格为 YYYY-MM-DD.md 或 .txt。
5. 想停止自动运行：点击“卸载开机自启”。
6. 更换邮箱或授权码：点击“配置邮件”，重新填写并发送测试邮件。

八、安全说明
- 授权码由 Windows 加密保存，不要发给别人。
- 不要测试未授权资产，所有安全实验只能在自己虚拟机或授权靶场进行。
- 本项目只负责生成计划、桌面通知和发送邮件，不扫描网络、不攻击任何目标。
"""


if __name__ == "__main__":
    app = App()
    app.mainloop()


