import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formataddr
from pathlib import Path


def send_email(
        sender_email: str,
        sender_name: str,
        sender_auth: str,
        receiver_email: str,
        subject: str,
        content: str,
        content_type: str,
        attachment_path: str
) -> None:
    """
    简化版邮件发送函数（支持附件）

    :param sender_email: 发件人邮箱地址
    :param sender_name: 发件人显示名称
    :param sender_auth: 邮箱授权码
    :param receiver_email: 收件人邮箱地址
    :param subject: 邮件主题
    :param content: 邮件内容
    :param content_type: 内容类型（html/plain）
    :param attachment_path: 附件文件路径
    """
    msg = MIMEMultipart()
    msg["From"] = formataddr([sender_name, sender_email])
    msg["To"] = receiver_email
    msg['Subject'] = subject

    msg.attach(MIMEText(content, content_type, "utf-8"))

    resolved_path = Path(attachment_path).resolve()
    if resolved_path.exists() and resolved_path.is_file():
        with open(resolved_path, "rb") as attachment_file:
            part = MIMEApplication(
                attachment_file.read(),
                Name=resolved_path.name
            )
        part['Content-Disposition'] = f'attachment; filename="{resolved_path.name}"'
        msg.attach(part)

    server = smtplib.SMTP_SSL("smtp.qq.com", 465)
    server.login(sender_email, sender_auth)
    server.sendmail(sender_email, receiver_email, msg.as_string())
    try:
        server.quit()
    except:
        pass



