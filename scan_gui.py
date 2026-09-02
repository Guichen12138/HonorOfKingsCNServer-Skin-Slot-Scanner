# -*- coding: utf-8 -*-
"""
王者荣耀皮肤扫描工具 GUI

双击「皮肤扫描工具.exe」启动（开发时也可以 python scan_gui.py）。

设计要点：
- hero_scan.py / skin_scan.py 以「子进程」方式运行，
  永远执行目录里的最新脚本，改完脚本不用重打 exe
- 脚本 print 的内容实时进输出栏，
  「扫描进度：x/999」和「[x/n]」两种格式自动转成进度条
- skin_scan 里的 input() 等待 → 点「继续」按钮相当于按回车
- Git 区块：提交信息输入框 + 一键 add / commit / push，
  推送默认走 Clash 代理（这台机器直连 GitHub 不通）

重新打包 exe（PyInstaller 在 F:\\conda3 里）：

    F:\\conda3\\python.exe -m PyInstaller --noconfirm --clean \
        --onefile --noconsole --name "皮肤扫描工具" \
        --distpath "E:\\feverapps\\123" \
        --workpath <临时目录> --specpath <临时目录> \
        scan_gui.py
"""

import os
import re
import shutil
import subprocess
import sys

from PyQt5.QtCore import (
    Qt,
    QProcess,
    QThread,
    pyqtSignal,
)
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QApplication,
    QCheckBox,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


# ============================================================
# 路径
# ============================================================

# 打包成 exe 后，工作目录取 exe 所在目录；
# 开发时取本文件所在目录
if getattr(sys, "frozen", False):

    APP_DIR = os.path.dirname(
        os.path.abspath(sys.executable)
    )

else:

    APP_DIR = os.path.dirname(
        os.path.abspath(__file__)
    )


EXCEL_NAME = "hero_skin_scan.xlsx"
HERO_SCRIPT_NAME = "hero_scan.py"
SKIN_SCRIPT_NAME = "skin_scan.py"

# Clash Verge 本机代理端口
PROXY_URL = "http://127.0.0.1:7897"


# ============================================================
# 输出解析（不依赖 Qt，方便单独测试）
# ============================================================

# 两种进度格式：
#   hero_scan：扫描进度：x/999
#   skin_scan：[x/n] 编号 名字
_PROG_HERO = re.compile(
    r"扫描进度：\s*(\d+)\s*/\s*(\d+)"
)
_PROG_SKIN = re.compile(
    r"^\s*\[(\d+)\s*/\s*(\d+)\]"
)


def classify_line(line):

    """是进度行就返回 (done, total)，否则返回 None"""

    match = (
        _PROG_HERO.search(line)
        or _PROG_SKIN.search(line)
    )

    if not match:
        return None

    return int(match.group(1)), int(match.group(2))


def feed_lines(buffer, text):

    """
    子进程输出按行切开。

    进度行用 \\r 覆盖、不带换行，所以 \\r / \\n 都算分隔；
    最后一段可能不完整，留进缓冲等下一段。
    """

    buffer += text

    parts = re.split(r"[\r\n]", buffer)

    return parts.pop(), parts


# ============================================================
# 找到能跑脚本的 Python 解释器
# ============================================================

def find_python():

    # 开发模式：直接用当前解释器
    if not getattr(sys, "frozen", False):
        return sys.executable

    # 打包后：sys.executable 是 exe 自己，
    # 需要找系统里的 Python
    candidates = [
        r"F:\conda3\python.exe",
        shutil.which("python"),
        shutil.which("python3"),
    ]

    for candidate in candidates:

        if candidate and os.path.exists(candidate):
            return candidate

    # 实在找不到，交给系统 PATH 碰运气
    return "python"


# ============================================================
# Git 工作线程
# ============================================================

class GitWorker(QThread):

    output = pyqtSignal(str)
    done = pyqtSignal(int)

    def __init__(self, steps, workdir, parent=None):

        super().__init__(parent)

        # steps: [{"argv": [...], "stdin": str|None, "timeout": 秒}]
        self.steps = steps
        self.workdir = workdir

    def run(self):

        final_rc = 0

        for step in self.steps:

            argv = step["argv"]

            self.output.emit(
                "\n$ " + " ".join(argv) + "\n"
            )

            try:

                proc = subprocess.run(
                    argv,
                    input=step.get("stdin"),
                    cwd=self.workdir,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    timeout=step.get("timeout", 120),
                )

            except FileNotFoundError:

                self.output.emit(
                    "找不到 git 命令，请确认已安装 Git。\n"
                )
                self.done.emit(1)
                return

            except subprocess.TimeoutExpired:

                self.output.emit(
                    "命令超时（网络可能不通）。\n"
                )
                self.done.emit(1)
                return

            if proc.stdout:
                self.output.emit(proc.stdout)

            if proc.stderr:
                self.output.emit(proc.stderr)

            if proc.returncode != 0:

                self.output.emit(
                    f"（返回码 {proc.returncode}，"
                    f"后续步骤已停止）\n"
                )
                final_rc = proc.returncode
                self.done.emit(final_rc)
                return

        self.done.emit(final_rc)


# ============================================================
# 主窗口
# ============================================================

class MainWindow(QMainWindow):

    def __init__(self, workdir=APP_DIR, python_exe=None):

        super().__init__()

        self.workdir = workdir
        self.python_exe = python_exe or find_python()

        self.excel_path = os.path.join(
            workdir, EXCEL_NAME
        )

        # 当前正在跑的脚本子进程
        self.process = None

        # 子进程输出的行缓冲（进度行没有换行符，需要自己拼）
        self._line_buffer = ""

        # git 线程
        self.git_worker = None

        self.setWindowTitle("王者荣耀皮肤扫描工具")
        self.resize(860, 640)

        self._build_ui()
        self._apply_style()

        # 启动时刷新一次 git 状态
        self.refresh_git_status()

    # ========================================================
    # 界面
    # ========================================================

    def _build_ui(self):

        central = QWidget()
        self.setCentralWidget(central)

        layout = QVBoxLayout(central)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        # ----------------------------------------------------
        # 顶部：两个运行按钮 + 打开 Excel
        # ----------------------------------------------------

        top_row = QHBoxLayout()
        top_row.setSpacing(10)

        self.btn_hero = QPushButton(
            "运行 hero_scan\n增量扫描皮肤格子"
        )
        self.btn_hero.clicked.connect(
            lambda: self.start_scan("hero")
        )

        self.btn_skin = QPushButton(
            "运行 skin_scan\n打开未命名海报"
        )
        self.btn_skin.clicked.connect(
            lambda: self.start_scan("skin")
        )

        self.btn_excel = QPushButton("打开 Excel")
        self.btn_excel.clicked.connect(
            self.open_excel
        )

        top_row.addWidget(self.btn_hero, 3)
        top_row.addWidget(self.btn_skin, 3)
        top_row.addWidget(self.btn_excel, 2)

        layout.addLayout(top_row)

        # ----------------------------------------------------
        # 进度条 + 状态
        # ----------------------------------------------------

        self.status_label = QLabel("就绪")
        self.status_label.setObjectName("statusLabel")

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setFormat("%v / %m")

        layout.addWidget(self.status_label)
        layout.addWidget(self.progress)

        # ----------------------------------------------------
        # 输出栏
        # ----------------------------------------------------

        self.output = QPlainTextEdit()
        self.output.setReadOnly(True)
        self.output.setFont(
            QFont("Consolas", 9)
        )
        self.output.setPlaceholderText(
            "脚本输出会实时显示在这里……"
        )

        layout.addWidget(self.output, 1)

        # ----------------------------------------------------
        # 继续 / 中断 / 清空
        # ----------------------------------------------------

        ctl_row = QHBoxLayout()
        ctl_row.setSpacing(10)

        self.btn_continue = QPushButton(
            "继续（相当于按回车）"
        )
        self.btn_continue.setEnabled(False)
        self.btn_continue.clicked.connect(
            self.send_enter
        )

        self.btn_stop = QPushButton("中断当前任务")
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(
            self.stop_scan
        )

        self.btn_clear = QPushButton("清空输出")
        self.btn_clear.clicked.connect(
            self.output.clear
        )

        ctl_row.addWidget(self.btn_continue, 3)
        ctl_row.addWidget(self.btn_stop, 2)
        ctl_row.addWidget(self.btn_clear, 2)

        layout.addLayout(ctl_row)

        # ----------------------------------------------------
        # Git 区块
        # ----------------------------------------------------

        git_box = QGroupBox("Git 提交")
        git_layout = QVBoxLayout(git_box)
        git_layout.setSpacing(8)

        status_row = QHBoxLayout()

        self.git_status_label = QLabel(
            "状态读取中……"
        )
        self.git_status_label.setObjectName(
            "gitStatusLabel"
        )

        self.btn_git_refresh = QPushButton("刷新状态")
        self.btn_git_refresh.clicked.connect(
            self.refresh_git_status
        )

        status_row.addWidget(self.git_status_label, 1)
        status_row.addWidget(self.btn_git_refresh)

        git_layout.addLayout(status_row)

        self.commit_msg = QLineEdit()
        self.commit_msg.setPlaceholderText(
            "更新内容，例如：更新鲁班大师-傩神司"
        )
        self.commit_msg.returnPressed.connect(
            self.commit_and_push
        )

        git_layout.addWidget(self.commit_msg)

        push_row = QHBoxLayout()

        self.proxy_check = QCheckBox(
            "推送走 Clash 代理 (127.0.0.1:7897)"
        )
        self.proxy_check.setChecked(True)
        self.proxy_check.setToolTip(
            "这台机器直连 GitHub 不通，"
            "建议保持勾选；Clash 没开时可取消试试直连"
        )

        self.btn_commit = QPushButton("提交并推送")
        self.btn_commit.clicked.connect(
            self.commit_and_push
        )

        push_row.addWidget(self.proxy_check, 1)
        push_row.addWidget(self.btn_commit)

        git_layout.addLayout(push_row)

        layout.addWidget(git_box)

        # ----------------------------------------------------
        # 底部：解释器信息
        # ----------------------------------------------------

        self.env_label = QLabel(
            f"工作目录：{self.workdir}    "
            f"Python：{self.python_exe}"
        )
        self.env_label.setObjectName("envLabel")

        layout.addWidget(self.env_label)

    def _apply_style(self):

        self.setStyleSheet("""
            QMainWindow, QWidget {
                background: #f6f8fa;
                color: #0a2540;
                font-size: 13px;
            }
            QPushButton {
                background: #ffffff;
                border: 1px solid #d0d7de;
                border-radius: 6px;
                padding: 8px 12px;
            }
            QPushButton:hover { background: #f0f2f5; }
            QPushButton:disabled {
                color: #9aa4b2;
                background: #eef1f4;
            }
            #btn_hero {
                background: #635bff;
                color: #ffffff;
                border: none;
                font-weight: bold;
            }
            #btn_hero:hover { background: #5148e0; }
            #btn_skin {
                background: #0a2540;
                color: #ffffff;
                border: none;
                font-weight: bold;
            }
            #btn_skin:hover { background: #16355a; }
            QPlainTextEdit {
                background: #ffffff;
                border: 1px solid #d0d7de;
                border-radius: 6px;
            }
            QProgressBar {
                border: 1px solid #d0d7de;
                border-radius: 6px;
                background: #ffffff;
                text-align: center;
                height: 18px;
            }
            QProgressBar::chunk {
                background: #635bff;
                border-radius: 5px;
            }
            QGroupBox {
                border: 1px solid #d0d7de;
                border-radius: 6px;
                margin-top: 8px;
                padding-top: 8px;
                background: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 4px;
            }
            QLineEdit {
                border: 1px solid #d0d7de;
                border-radius: 6px;
                padding: 6px 8px;
                background: #ffffff;
            }
            #statusLabel { font-weight: bold; }
            #envLabel, #gitStatusLabel { color: #57606a; }
        """)

        self.btn_hero.setObjectName("btn_hero")
        self.btn_skin.setObjectName("btn_skin")

    # ========================================================
    # 运行脚本（子进程）
    # ========================================================

    def start_scan(self, kind):

        if self.process is not None:
            QMessageBox.information(
                self, "提示", "已有任务在运行。"
            )
            return

        script = os.path.join(
            self.workdir,
            HERO_SCRIPT_NAME
            if kind == "hero"
            else SKIN_SCRIPT_NAME,
        )

        if not os.path.exists(script):
            QMessageBox.warning(
                self, "缺少文件",
                f"找不到脚本：\n{script}"
            )
            return

        self.process = QProcess(self)
        self.process.setWorkingDirectory(self.workdir)

        # stdout / stderr 合并成一路
        self.process.setProcessChannelMode(
            QProcess.MergedChannels
        )

        # 子进程 print 中文时强制 UTF-8，
        # 否则 Windows 管道默认 GBK 会乱码
        env = QProcess.systemEnvironment()
        env = self._with_env(
            env, "PYTHONIOENCODING", "utf-8"
        )
        self.process.setEnvironment(env)

        # GUI 程序拉起的 python.exe 会闪一个黑窗口，
        # 用 CREATE_NO_WINDOW 压住
        def no_window(args):
            args.creationFlags = (
                args.creationFlags | 0x08000000
            )

        self._no_window_hook = no_window
        self.process.setCreateProcessArgumentsModifier(
            no_window
        )

        self.process.readyReadStandardOutput.connect(
            self.on_read
        )
        self.process.finished.connect(
            self.on_scan_finished
        )

        self._line_buffer = ""
        self._scan_kind = kind

        self.progress.setRange(0, 0)
        self.set_status(
            f"{kind}_scan 启动中……"
        )

        self.append_output(
            f"\n{'=' * 60}\n"
            f"启动 {script}\n"
            f"{'=' * 60}\n"
        )

        self._set_running_ui(True)

        self.process.start(
            self.python_exe, [script]
        )

    @staticmethod
    def _with_env(env, key, value):

        prefix = key + "="

        for i, item in enumerate(env):

            if item.startswith(prefix):
                env[i] = prefix + value
                return env

        env.append(prefix + value)
        return env

    # --------------------------------------------------------
    # 读取子进程输出
    # --------------------------------------------------------

    def on_read(self):

        data = bytes(
            self.process.readAllStandardOutput()
        )

        text = data.decode("utf-8", "replace")

        self._feed_text(text)

    def _feed_text(self, text):

        self._line_buffer, parts = feed_lines(
            self._line_buffer, text
        )

        lines = []
        for part in parts:

            if not part.strip():
                lines.append("")
                continue

            if self._try_progress(part):
                continue

            lines.append(part)

        if lines:
            self.append_output(
                "\n".join(lines) + "\n"
            )

        # 进程结束后缓冲区可能还有内容，
        # 由 on_scan_finished 里 flush

    def _try_progress(self, line):

        result = classify_line(line)

        if result is None:
            return False

        done, total = result

        if (
            self.progress.maximum() != total
            or self.progress.minimum() != 0
        ):
            self.progress.setRange(0, total)

        self.progress.setValue(done)

        self.set_status(
            f"{self._scan_kind}_scan 运行中… "
            f"{done}/{total}"
        )

        return True

    # --------------------------------------------------------
    # 等待 input() 时：点「继续」= 按回车
    # --------------------------------------------------------

    def send_enter(self):

        if self.process is not None:
            self.process.write(b"\n")

    def stop_scan(self):

        if self.process is None:
            return

        answer = QMessageBox.question(
            self, "中断任务",
            "确定要中断当前任务吗？\n\n"
            "hero_scan 在扫描结束时才保存 Excel，\n"
            "中途杀掉本次扫描结果会丢失。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )

        if answer == QMessageBox.Yes:
            self.process.kill()

    def on_scan_finished(self):

        # 缓冲区里可能还有半行
        if self._line_buffer.strip():

            if not self._try_progress(
                self._line_buffer
            ):
                self.append_output(
                    self._line_buffer + "\n"
                )

        self._line_buffer = ""

        code = self.process.exitCode()
        kind = self._scan_kind

        self.process.deleteLater()
        self.process = None

        self._set_running_ui(False)

        if code == 0:

            self.set_status(
                f"{kind}_scan 已完成"
            )

            if self.progress.maximum() > 0:
                self.progress.setValue(
                    self.progress.maximum()
                )

        else:

            self.set_status(
                f"{kind}_scan 已结束"
                f"（退出码 {code}）"
            )

        # 扫描可能改了 xlsx，顺手刷新 git 状态
        self.refresh_git_status()

    def _set_running_ui(self, running):

        self.btn_hero.setEnabled(not running)
        self.btn_skin.setEnabled(not running)
        self.btn_continue.setEnabled(running)
        self.btn_stop.setEnabled(running)

    # ========================================================
    # 打开 Excel
    # ========================================================

    def open_excel(self):

        if not os.path.exists(self.excel_path):

            QMessageBox.warning(
                self, "缺少文件",
                f"找不到：\n{self.excel_path}"
            )
            return

        os.startfile(self.excel_path)

    # ========================================================
    # Git
    # ========================================================

    def refresh_git_status(self):

        if self.git_worker is not None:
            return

        steps = [{
            "argv": [
                "git",
                "-c", "core.quotepath=false",
                "status", "-sb",
            ],
            "timeout": 30,
        }]

        self._run_git(steps, self._on_status_done)

    def _on_status_done(self, rc, text):

        lines = [
            line for line in text.splitlines()
            if line.strip()
        ]

        if rc != 0 or not lines:

            self.git_status_label.setText(
                "git 状态读取失败，详见输出栏"
            )
            return

        branch = lines[0]
        changes = len(lines) - 1

        self.git_status_label.setText(
            f"{branch}    "
            f"未提交改动：{changes} 个文件"
        )

    def commit_and_push(self):

        if self.git_worker is not None:

            QMessageBox.information(
                self, "提示",
                "Git 操作正在进行中。"
            )
            return

        if self.process is not None:

            QMessageBox.information(
                self, "提示",
                "扫描正在运行，等它结束再提交。"
            )
            return

        message = self.commit_msg.text().strip()

        if not message:

            QMessageBox.warning(
                self, "提示",
                "请先填写更新内容。"
            )
            return

        steps = [
            {
                "argv": [
                    "git", "add", "-A",
                ],
                "timeout": 60,
            },
            {
                "argv": [
                    "git",
                    "-c", "core.quotepath=false",
                    "commit", "-F", "-",
                ],
                "stdin": message,
                "timeout": 60,
            },
        ]

        push_argv = ["git"]

        if self.proxy_check.isChecked():

            push_argv += [
                "-c",
                "http.https://github.com.proxy="
                + PROXY_URL,
            ]

        push_argv.append("push")

        steps.append({
            "argv": push_argv,
            "timeout": 300,
        })

        self._run_git(
            steps,
            lambda rc, text: self._on_commit_done(rc),
        )

    def _on_commit_done(self, rc):

        if rc == 0:

            self.set_status("Git 提交并推送完成")
            self.commit_msg.clear()

        else:

            self.set_status(
                "Git 操作失败，详见输出栏"
            )

        self.refresh_git_status()

    def _run_git(self, steps, on_done):

        collected = []

        worker = GitWorker(steps, self.workdir, self)
        self.git_worker = worker

        worker.output.connect(self.append_output)
        worker.output.connect(collected.append)

        def finished(rc):

            self.git_worker = None

            self.btn_commit.setEnabled(True)
            self.btn_git_refresh.setEnabled(True)

            on_done(rc, "".join(collected))

            worker.deleteLater()

        worker.done.connect(finished)

        self.btn_commit.setEnabled(False)
        self.btn_git_refresh.setEnabled(False)

        worker.start()

    # ========================================================
    # 小工具
    # ========================================================

    def append_output(self, text):

        self.output.moveCursor(
            self.output.textCursor().End
        )
        self.output.insertPlainText(text)
        self.output.ensureCursorVisible()

    def set_status(self, text):

        self.status_label.setText(text)

    # ========================================================
    # 关闭窗口时：有任务在跑要先确认
    # ========================================================

    def closeEvent(self, event):

        if self.process is not None:

            answer = QMessageBox.question(
                self, "任务运行中",
                "扫描还在运行，关闭窗口会中断它。\n"
                "确定关闭吗？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )

            if answer != QMessageBox.Yes:
                event.ignore()
                return

            self.process.kill()
            self.process.waitForFinished(3000)

        event.accept()


# ============================================================
# 入口
# ============================================================

def main():

    app = QApplication(sys.argv)

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())


if __name__ == "__main__":

    main()
