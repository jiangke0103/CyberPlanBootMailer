from __future__ import annotations

import argparse
import json
import os
import re
import smtplib
import ssl
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CONFIG = ROOT / "config" / "email_config.json"
LOG_DIR = ROOT / "logs"

PROVIDERS = {
    "qq": ("smtp.qq.com", 587, "starttls"),
    "163": ("smtp.163.com", 465, "ssl"),
    "gmail": ("smtp.gmail.com", 587, "starttls"),
    "outlook": ("smtp-mail.outlook.com", 587, "starttls"),
}


def log(message: str):
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    path = LOG_DIR / "email.log"
    with path.open("a", encoding="utf-8") as f:
        f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {message}\n")


def read_config():
    if not CONFIG.exists():
        log("email skipped: config missing")
        return None
    try:
        return json.loads(CONFIG.read_text(encoding="utf-8"))
    except Exception as exc:
        log(f"email config error: {exc}")
        return None


def parse_tasks(markdown: str):
    lines = markdown.splitlines()
    phase = ""
    for line in lines:
        if line.startswith("> 阶段："):
            phase = line.replace("> 阶段：", "", 1).strip()
            break
    results = []
    in_results = False
    for line in lines:
        if line.startswith("## 今日三个结果"):
            in_results = True
            continue
        if in_results and line.startswith("## "):
            break
        if in_results and re.match(r"^\d+\.\s+", line):
            results.append(re.sub(r"^\d+\.\s+", "", line).strip())
    rows = []
    for line in lines:
        if re.match(r"^\|\s*\d{2}:\d{2}-\d{2}:\d{2}\s*\|", line):
            cols = [c.strip() for c in line.strip().strip("|").split("|")]
            if len(cols) >= 4:
                rows.append(cols[:4])
    return phase, results, rows


def html_escape(value: str) -> str:
    return (value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


def task_html(day: str, markdown: str) -> str:
    phase, results, rows = parse_tasks(markdown)
    result_html = "".join(f"<li>{html_escape(x)}</li>" for x in results)
    rows_html = "".join(
        "<tr>" + "".join(f"<td>{html_escape(cell)}</td>" for cell in row) + "</tr>"
        for row in rows
    )
    external_text = ""
    marker = "## 外部 AI 计划补充"
    if marker in markdown:
        external_text = markdown.split(marker, 1)[1].strip()
    external_html = ""
    if external_text:
        external_html = f'<h3>外部 AI 计划补充</h3><pre style="white-space:pre-wrap;background:#F2F3F5;padding:12px;border-radius:6px">{html_escape(external_text)}</pre>'
    return f"""
    <div style="font-family:'Microsoft YaHei',Arial,sans-serif;line-height:1.65;color:#222">
      <h2 style="color:#1F4E79">网络安全就业计划 {day}</h2>
      <p style="background:#FFF4CE;padding:10px">{html_escape(phase)}</p>
      <h3>今日三个结果</h3><ol>{result_html}</ol>
      <h3>时间表</h3>
      <table style="border-collapse:collapse;width:100%;font-size:13px">
        <thead><tr style="background:#1F4E79;color:white">
          <th style="border:1px solid #bbb;padding:6px">时间</th>
          <th style="border:1px solid #bbb;padding:6px">任务</th>
          <th style="border:1px solid #bbb;padding:6px">具体怎么做</th>
          <th style="border:1px solid #bbb;padding:6px">必须产出</th>
        </tr></thead><tbody>{rows_html}</tbody>
      </table>
      <p style="margin-top:16px">先做第一步，不要同时开多个方向。所有安全实验只能在自有虚拟机或明确授权靶场中进行。</p>
      {external_html}
    </div>
    """


def review_html(day: str) -> str:
    return f"""
    <div style="font-family:'Microsoft YaHei',Arial,sans-serif;line-height:1.8;color:#222">
      <h2 style="color:#1F4E79">网络安全计划晚间复盘 {day}</h2>
      <p>21:40 开始复盘，20 分钟内完成。</p>
      <ol>
        <li>今日完成：</li><li>今日产出：</li><li>最大卡点：</li>
        <li>明天三件事：</li><li>今日得分：__ / 10</li>
      </ol>
      <p>请把答案填写到复盘表或回复邮件，以便后续调整计划。</p>
    </div>
    """


def send_email(mode: str, day: str):
    config = read_config()
    if not config:
        return 0
    password = os.environ.get("CYBER_SMTP_PASSWORD")
    if not password:
        log("email skipped: password missing")
        return 0

    day_dir = ROOT / "data" / "每日任务" / day
    task_md = day_dir / f"{day}_今日任务.md"
    task_docx = day_dir / f"{day}_今日任务.docx"
    review_md = day_dir / f"{day}_晚间复盘.md"
    review_docx = day_dir / f"{day}_晚间复盘.docx"

    def latest_docx(base):
        pointer = base.parent / "current_docx.txt"
        if pointer.exists():
            candidate = base.parent / pointer.read_text(encoding="utf-8").strip()
            if candidate.exists():
                return candidate
        return base

    if mode == "Morning":
        if not task_md.exists():
            log(f"email skipped: task markdown missing for {day}")
            return 0
        subject = f"今日网络安全计划 {day}"
        html = task_html(day, task_md.read_text(encoding="utf-8"))
        attachments = [task_md, latest_docx(task_docx)]
    elif mode == "Review":
        subject = f"网络安全计划复盘提醒 {day}"
        html = review_html(day)
        attachments = [review_md, review_docx]
    else:
        subject = "网络安全计划邮件测试"
        html = "<p>邮件配置成功，之后的每日计划和复盘提醒会发送到这里。</p>"
        attachments = []

    sender = config["sender"]
    recipient = config.get("recipient") or sender
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = recipient
    msg.set_content("请使用支持 HTML 的邮件客户端查看本邮件。")
    msg.add_alternative(html, subtype="html")
    for path in attachments:
        if path.exists():
            data = path.read_bytes()
            if path.suffix.lower() == ".docx":
                msg.add_attachment(data, maintype="application", subtype="vnd.openxmlformats-officedocument.wordprocessingml.document", filename=path.name)
            else:
                msg.add_attachment(data, maintype="text", subtype="markdown", filename=path.name)

    host = config["smtp_host"]
    port = int(config["smtp_port"])
    security = config.get("security", "starttls").lower()
    try:
        context = ssl.create_default_context()
        if security == "ssl" or port == 465:
            smtp = smtplib.SMTP_SSL(host, port, timeout=30, context=context)
        else:
            smtp = smtplib.SMTP(host, port, timeout=30)
            smtp.ehlo()
            smtp.starttls(context=context)
            smtp.ehlo()
        smtp.login(config["username"], password)
        smtp.send_message(msg)
        smtp.quit()
        log(f"email sent: mode={mode}, day={day}, to={recipient}")
        return 0
    except Exception as exc:
        log(f"email failed: mode={mode}, day={day}, error={exc}")
        return 1


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--mode", choices=["Morning", "Review", "Test"], default="Morning")
    parser.add_argument("--date", default=datetime.now().strftime("%Y-%m-%d"))
    args = parser.parse_args()
    raise SystemExit(send_email(args.mode, args.date))


if __name__ == "__main__":
    main()



