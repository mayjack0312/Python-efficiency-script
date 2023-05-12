import smtplib                        # smtplib发送邮件
from email.mime.text import MIMEText  # 构造文本内容
from email.header import Header       # 构造标题内容

# 设置服务器地址
mail_host = "smtp.qq.com"
# 设置服务器端口
mail_port = 465

# 初始化发送方账号
sender = "28480*****@qq.com"

# QQ邮件登录账号
mail_user = "28480*****"
# QQ邮箱第三方授权码
mail_pass = "gwgcqikk********"

def post_email(per_info):
    # 构造文本对象，三个参数：文本内容，设置文本格式,设置编码
    message = MIMEText(per_info["正文"],"plain","utf-8")
    # 文本对象 添加 发送者
    message["From"] = sender
    # 文本对象 添加 接收者
    message["To"] = per_info["收件人邮箱"]
    # 文本对象 添加 标题
    message["Subject"] = Header(per_info["邮件主题"])

    # 创建 SMTP 对象，连接目标服务器
    smtpObj = smtplib.SMTP_SSL(mail_host,mail_port)
    # 自己账号登录
    smtpObj.login(mail_user,mail_pass)
    # 发送邮件到目标地址  注意：信息由MTMEText对象 转为 字符串对象
    smtpObj.sendmail(sender,per_info["收件人邮箱"],message.as_string())
    # 结束 SMTP 对象
    smtpObj.quit()

if __name__ == '__main__':
    # 读取 邮件.xlsx
    email_info_df = pd.read_excel("邮件.xlsx")

    # 使用apply方法 将email_info_df中每一行 映射到post_email函数中
    email_info_df.apply(post_email,axis=1)
