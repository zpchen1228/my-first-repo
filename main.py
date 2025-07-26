from get_exchange_rate import download_exchange_rate_data
from send_email import send_email
from get_data_from_excel import get_latest_rates
import schedule
import time
import datetime
import pytz


def main_task():
    """主任务函数：下载汇率数据并发送邮件"""
    print(f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - 开始执行汇率邮件任务")

    # 下载汇率数据
    exchange_file = r"D:\czp\data\auto_download\exchange_rate\exchange_rate.xlsx"
    download_exchange_rate_data()

    # 获取最新的美元和欧元汇率
    currencies = ["USD", "EUR"]
    latest_rates = get_latest_rates(exchange_file, currencies)

    # 动态生成汇率显示内容
    rates_html = ""
    for currency_code in currencies:
        if currency_code in latest_rates:
            rate = latest_rates[currency_code]
            currency_name = "USD" if currency_code == "USD" else "EUR"
            rates_html += f"<li>{currency_name}/CNY: <b>{rate}</b></li>"
        else:
            rates_html += f"<li>{currency_code} rate not found</li>"

    # 更新的邮件内容
    email_content = f"""
    <html>
        <body>
            <p><b>Dear Zhi Ping,</b></p>
            <p>Please find today's foreign exchange rate report attached.</p>

            <p><b>Latest Key Currency Rates ({datetime.datetime.now().strftime('%Y-%m-%d')}):</b></p>
            <ul>
                {rates_html}
            </ul>

            <p><b>Report Highlights:</b></p>
            <ul>
                <li>Rates are the most recent entries from the People's Bank of China</li>
                  <li>Data reflects the latest market exchange rates</li>
                <li>Full data set available in the attached Excel file</li>
            </ul>

            <p>Please let me know if you need additional analysis or specific currency pairs.</p>
            <br>
            <p><b>Best Regards,</b></p>
            <p>Chen Zhi Ping<br>
            Financial Data Analyst</p>
        </body>
    </html>
    """

    # 发送邮件
    receive_address=["qidao1200@outlook.com"]
    # 在调用send_email前添加转换（新增代码）
    receive_address_str = ", ".join(receive_address)  # 将列表转为字符串

    try:
        send_email(
            sender_email="3275981857@qq.com",
            sender_name="Chen Zhi Ping",
            sender_auth="gedcfexluvakciaf",
            receiver_email=receive_address_str,
            subject=f"Daily Exchange Rate Report - {datetime.datetime.now().strftime('%Y-%m-%d')}",
            content=email_content,
            content_type="html",
            attachment_path=exchange_file
        )
        print(f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - 邮件发送成功")
    except Exception as e:
        print(f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - 邮件发送失败: {str(e)}")


if __name__ == "__main__":
    # 设置中国时区
    china_tz = pytz.timezone('Asia/Shanghai')

    # 立即执行一次任务
    print(f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - 程序启动，立即执行首次任务")
    main_task()

    # 添加每日9:30的定时任务
    schedule.every().day.at("09:30", china_tz).do(main_task)

    print(f"定时任务已启动，下次执行时间: {schedule.next_run().astimezone(china_tz).strftime('%Y-%m-%d %H:%M:%S')}")

    # 主循环保持程序运行
    while True:
        schedule.run_pending()
        # 每60秒检查一次是否有任务需要执行
        time.sleep(60)