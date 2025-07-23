import pandas as pd
import yagmail
import os
import json

# 读取邮箱配置
with open('email_config.json', 'r') as f:
    config = json.load(f)
your_email = config['email']
your_password = config['password']

# 附件目录
attachment_dir = './attachments'
os.makedirs(attachment_dir, exist_ok=True)

# 初始化发件客户端（SMTP）
yag = yagmail.SMTP(
    user=your_email,
    password=your_password,
    host='smtp.exmail.qq.com',
    port=465,
    smtp_ssl=True
)

# 读取 Excel
df = pd.read_excel('emails_comtent.xlsx')

# 批量处理邮件
for index, row in df.iterrows():
    if str(row.get('是否发送', '')).strip().upper() != 'Y':
        print(f"跳过第 {index + 2} 行（未标记发送）")
        continue

    # 收件人/抄送人
    to_raw = str(row['收件人邮箱']).replace('，', ',') if pd.notna(row.get('收件人邮箱')) else ''
    cc_raw = str(row['抄送人邮箱']).replace('，', ',') if pd.notna(row.get('抄送人邮箱')) else ''

    to_list = [addr.strip() for addr in to_raw.split(',') if addr.strip()]
    cc_list = [addr.strip() for addr in cc_raw.split(',') if addr.strip()]

    if not to_list:
        print(f"第 {index + 2} 行没有有效收件人，跳过")
        continue

    subject = str(row.get('邮件标题', ''))
    contents = str(row.get('邮件正文', ''))
    attachment_name = row.get('附件名称')
    attachment_path = None

    if pd.notna(attachment_name):
        attachment_path = os.path.join(attachment_dir, attachment_name)
        if not os.path.exists(attachment_path):
            with open(attachment_path, 'w') as f:
                pass
        print(f"附件已准备：{attachment_path}")

    try:
        yag.send(to=to_list, cc=cc_list, subject=subject, contents=contents, attachments=attachment_path)
        print(f"已发送：{', '.join(to_list)}")
        if cc_list: print(f"抄送：{', '.join(cc_list)}")
    except Exception as e:
        print(f"发送失败：{', '.join(to_list)}，原因：{e}")
