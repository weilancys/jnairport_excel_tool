import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import datetime
import os


def prepare_email(rows_lt_30_days, rows_expired, excel_filename, sender_email, recipients):
    now = datetime.datetime.now()
    email_message = MIMEMultipart()

    table_body_for_lt_30_days = ""
    for row_index, row in enumerate(rows_lt_30_days):
        table_row_for_lt_30_days = """
            <tr>
                <td>{row_index}</td>
                <td>{row[0]}</td>
                <td>{row[1]}</td>
                <td>{row[2]}</td>
                <td>{row[3]}</td>
                <td>{row[4]}</td>
                <td>{row[5]}</td>
                <td>{row[6]}</td>
                <td>{row[7]}</td>
                <td>{row[8]}</td>
                <td>{row[9]}</td>
                <td>{row[10]}</td>
                <td>{row[11]}</td>
                <td>{row[12]}</td>
                <td>{row[13]}</td>
                <td>{row[14]}</td>
                <td>{row[15]}</td>
                <td>{row[16]}</td>
                <td>{row[17]}</td>
                <td>{row[18]}</td>
                <td>{row[19]}</td>
                <td>{row[20]}</td>
                <td>{row[21]}</td>
            </tr>
        """.format(row_index=row_index+1, row=row)
        table_body_for_lt_30_days += table_row_for_lt_30_days

    table_body_for_expired = ""
    for row_index, row in enumerate(rows_expired):
        table_row_for_expired = """
            <tr>
                <td>{row_index}</td>
                <td>{row[0]}</td>
                <td>{row[1]}</td>
                <td>{row[2]}</td>
                <td>{row[3]}</td>
                <td>{row[4]}</td>
                <td>{row[5]}</td>
                <td>{row[6]}</td>
                <td>{row[7]}</td>
                <td>{row[8]}</td>
                <td>{row[9]}</td>
                <td>{row[10]}</td>
                <td>{row[11]}</td>
                <td>{row[12]}</td>
                <td>{row[13]}</td>
                <td>{row[14]}</td>
                <td>{row[15]}</td>
                <td>{row[16]}</td>
                <td>{row[17]}</td>
                <td>{row[18]}</td>
                <td>{row[19]}</td>
                <td>{row[20]}</td>
                <td>{row[21]}</td>
            </tr>
        """.format(row_index=row_index+1, row=row)
        table_body_for_expired += table_row_for_expired
    
    table_head = """
    <thead>
        <tr>
            <td>序号</td>
            <td>原Excel文件内序号</td>
            <td>租赁单位名称</td>
            <td>别名</td>
            <td>详细地址</td>
            <td>联系人</td>
            <td>联系电话</td>
            <td>申请时间</td>
            <td>上级主管责任单位</td>
            <td>信息点号</td>
            <td>VLAN</td>
            <td>上联交换机</td>
            <td>端口</td>
            <td>IP地址</td>
            <td>实施人员</td>
            <td>是否签订协议</td>
            <td>租赁带宽</td>
            <td>上期缴纳费用金额</td>
            <td>上期缴纳费用详细说明</td>
            <td>缴费开始时间</td>
            <td>缴费截止时间</td>
            <td>是否电话催缴</td>
            <td>备注</td>
        </tr>
    </thead>
    """

    html_msg = """
    <html>
        <body>
            <div class="header">
                <h3>筛选记录</h3>
                <p>筛选时间：{time}</p>
                <p>Excel文件名：{excel_filename}</p>
            </div>
            <div class="data">
                <div class="table-lt-30-days">
                    <h4>缴费截止日期至{date}不足30天的记录：</h4>
                    <p>共{row_count_for_lt_30_days}条记录</p>
                    <table border="1">
                        {table_head}
                        <tbody>
                            {table_body_for_lt_30_days}
                        </tbody>
                    </table>
                </div>
                <div class="table-expired">
                    <h4>已超期记录：</h4>
                    <p>共{row_count_for_expired}条记录</p>
                    <table border="1">
                        {table_head}
                        <tbody>
                            {table_body_for_expired}
                        </tbody>
                    </table>
                </div>
            </div>
        </body>
    </html>
    """.format(
        table_head = table_head,
        date = now.strftime("%Y-%m-%d"),
        time = now.strftime("%Y-%m-%d %H:%M:%S"),
        row_count_for_lt_30_days = len(rows_lt_30_days),
        row_count_for_expired = len(rows_expired),
        excel_filename = excel_filename,
        table_body_for_lt_30_days = table_body_for_lt_30_days,
        table_body_for_expired = table_body_for_expired,
    )


    html_part = MIMEText(html_msg, "html")
    email_message.attach(html_part)

    email_message["From"] = sender_email
    email_message["To"] = ", ".join(recipients)
    email_message["Subject"] = "{time} 筛选记录".format(time=now.strftime("%Y-%m-%d %H:%M:%S"))

    return email_message


def send_notice_mail(email_message, recipients, smtp_auth):
    smtp_client = smtplib.SMTP(smtp_auth["SMTP_SERVER_ADDR"])
    # smtp_client.connect()
    smtp_client.starttls()
    smtp_client.login(smtp_auth["SENDER_EMAIL"], smtp_auth["PASSWORD"])
    smtp_client.sendmail(smtp_auth["SENDER_EMAIL"], recipients, email_message.as_string())

