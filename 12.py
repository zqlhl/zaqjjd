import os
import resend
import time
from datetime import datetime
import base64

resend.api_key = "re_18USgDeb_Q7QYHb9sEv73TY9Px4LnDoqH"

def send_email_with_word_attachment():
    # Word文档路径
    word_document_path = "C:/Users/zq/Desktop/document.docx"
    
    # 检查文件是否存在
    if not os.path.exists(word_document_path):
        print(f"错误：找不到文件 {word_document_path}")
        return False
    
    # 读取Word文档内容并进行Base64编码
    with open(word_document_path, "rb") as f:
        attachment_content = f.read()
        # 将字节内容转换为Base64字符串
        encoded_content = base64.b64encode(attachment_content).decode('utf-8')
    
    params = {
        "from": "szjsjdc@zzqqyy.sbs",
        "to": ["lhl2006111@163.com"],
        "subject": f"报告 - {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "html": "<strong>请查收附件中的Word文档。</strong>",
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

# 每隔6小时发送一次（6小时 = 6 * 60 * 60 秒 = 21600秒）
while True:
    send_email_with_word_attachment()
    print("等待6小时后再次发送...")
    time.sleep(21600)  # 休眠6小时