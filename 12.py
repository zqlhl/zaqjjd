import os
import resend
import time
from datetime import datetime
import base64
import random
import string

# 使用环境变量获取API密钥
resend.api_key = os.getenv("RESEND_API_KEY", "re_18USgDeb_Q7QYHb9sEv73TY9Px4LnDoqH")

def generate_random_username():
    """生成随机用户名"""
    adjectives = ['happy', 'smart', 'quick', 'bright', 'cool', 'fast', 'wise', 'bold', 'clever', 'swift']
    nouns = ['user', 'sender', 'report', 'daily', 'auto', 'mail', 'doc', 'file', 'robot', 'agent']
    return f"{random.choice(adjectives)}{random.choice(nouns)}{random.randint(10, 999)}"

def send_email_with_word_attachment():
    # Word文档路径 - GitHub环境中需要调整
    word_document_path = "举报信.docx"  # 假设文档在仓库根目录

    # 检查文件是否存在
    if not os.path.exists(word_document_path):
        print(f"错误：找不到文件 {word_document_path}")
        return False

    # 读取Word文档内容并进行Base64编码
    with open(word_document_path, "rb") as f:
        attachment_content = f.read()
        # 将字节内容转换为Base64字符串
        encoded_content = base64.b64encode(attachment_content).decode('utf-8')

    # 生成随机用户名
    random_username = generate_random_username()

    params = {
        "from": f"{random_username}@zzqqyy.sbs",
        "to": ["3062312510@qq.com"],
        "subject": f"报告 - {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "html": "<strong>贫困补助匿名举报信。请领导看一下，谢谢</strong>",
        "attachments": [
            {
                "filename": os.path.basename(word_document_path),
                "content": encoded_content,
                "content_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            }
        ]
    }

    try:
        response = resend.Emails.send(params)
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 邮件发送成功")
        return True
    except Exception as e:
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 邮件发送失败: {e}")
        return False

# GitHub Actions中只执行一次，不使用循环
if __name__ == "__main__":
    send_email_with_word_attachment()




