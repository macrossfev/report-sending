#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试163邮箱配置"""

import json
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# 读取配置
with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

smtp_config = config['smtp']

print("="*60)
print("163邮箱配置测试")
print("="*60)
print(f"发件邮箱: {smtp_config['sender_email']}")
print(f"SMTP服务器: {smtp_config['server']}")
print(f"SMTP端口: {smtp_config['port']}")
print(f"授权码长度: {len(smtp_config['sender_password'])} 位")
print("="*60)

# 测试邮箱地址（发送给自己）
test_recipient = input("\n请输入测试收件邮箱（直接回车使用发件邮箱自己）: ").strip()
if not test_recipient:
    test_recipient = smtp_config['sender_email']

print(f"\n准备发送测试邮件到: {test_recipient}")
print("正在连接163邮箱服务器...\n")

try:
    # 创建邮件
    msg = MIMEMultipart()
    msg['From'] = f"{smtp_config['sender_name']} <{smtp_config['sender_email']}>"
    msg['To'] = test_recipient
    msg['Subject'] = "测试邮件 - 163邮箱配置验证"

    body = """这是一封测试邮件。

如果您收到这封邮件，说明163邮箱配置正确！

发送时间: 2026-02-04
系统: PDF加密邮件发送系统
"""

    msg.attach(MIMEText(body, 'plain', 'utf-8'))

    # 连接SMTP服务器
    print("步骤1: 连接SMTP服务器...")
    server = smtplib.SMTP_SSL(smtp_config['server'], smtp_config['port'], timeout=30)
    print("✓ 连接成功\n")

    # 登录
    print("步骤2: 登录163邮箱...")
    server.login(smtp_config['sender_email'], smtp_config['sender_password'])
    print("✓ 登录成功\n")

    # 发送邮件
    print("步骤3: 发送邮件...")
    server.send_message(msg)
    print("✓ 邮件发送成功\n")

    server.quit()

    print("="*60)
    print("✓ 测试成功！")
    print("="*60)
    print(f"\n请检查邮箱 {test_recipient}")
    print("如果没收到，请检查：")
    print("1. 垃圾邮件/广告邮件文件夹")
    print("2. 邮箱是否设置了拦截规则")
    print("3. 稍等几分钟再查看")

except smtplib.SMTPAuthenticationError as e:
    print("✗ 认证失败！")
    print("\n可能的原因：")
    print("1. 授权码错误")
    print("2. 163邮箱未开启SMTP服务")
    print("3. 授权码已过期")
    print(f"\n详细错误: {e}")

except smtplib.SMTPException as e:
    print(f"✗ SMTP错误: {e}")

except Exception as e:
    print(f"✗ 发送失败: {e}")
    import traceback
    traceback.print_exc()
