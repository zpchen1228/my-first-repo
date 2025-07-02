import pandas as pd
import os
import time
from datetime import datetime, timedelta
import sched
import threading


def process_data():
    try:
        # 设置文件路径
        input_dir = r'D:\czp\data\python_data'
        input_path = os.path.join(input_dir, 'product_pk.xlsx')

        # 检查输入文件是否存在
        if not os.path.exists(input_path):
            print(f"警告: 输入文件不存在 ({input_path})")
            return

        # 读取Excel数据
        df = pd.read_excel(input_path, engine='openpyxl')

        # ===== 数据处理逻辑 =====
        # 1. 新增MFLB列（Material Description的第2-19位）
        if 'Material Description' in df.columns:
            df['MFLB'] = df['Material Description'].str[1:19]
        else:
            print("警告: 数据中缺少 'Material Description' 列")
            return

        # 2. 新增Series列（MFLB的1-4位）
        df['Series'] = df['MFLB'].str[0:4]

        # 3. 新增Brand列（根据MFLB第5位判断）
        def get_brand(mflb):
            if not isinstance(mflb, str) or len(mflb) < 5:
                return 'Unknown'
            return 'LI' if mflb[4] == '1' else 'HI' if mflb[4] == '3' else 'Unknown'

        df['Brand'] = df['MFLB'].apply(get_brand)

        # 4. 新增Frame列（根据MFLB第5-7位判断）
        def get_frame(mflb):
            if not isinstance(mflb, str) or len(mflb) < 7:
                return 'Unknown'
            segment = mflb[4:7]
            return '30' if segment == '103' else '65' if segment == '306' else 'Unknown'

        df['Frame'] = df['MFLB'].apply(get_frame)

        # ===== 保存处理结果 =====
        # 生成带日期时间戳的文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"{timestamp}_production_pk.xlsx"
        output_path = os.path.join(input_dir, output_filename)

        # 保存Excel文件
        df.to_excel(output_path, index=False)

        print(f"[{datetime.now().strftime('%H:%M:%S')}] 数据处理完成: {output_filename}")

    except Exception as e:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] 处理失败: {str(e)}")
        # 记录详细错误日志
        error_log = os.path.join(input_dir, f"error_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
        with open(error_log, 'w') as f:
            f.write(f"错误时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"错误详情: {str(e)}\n")


def schedule_task():
    """每5分钟执行一次任务"""
    # 第一次执行在23:13
    next_time = datetime.now().replace(hour=23, minute=13, second=0, microsecond=0)

    # 如果当前时间已经过了23:13，设置下一个执行点为5分钟后
    if datetime.now() >= next_time:
        next_time = datetime.now() + timedelta(minutes=5)

    print(f"数据处理定时任务已启动")
    print(f"下次执行时间: {next_time.strftime('%H:%M:%S')}")

    while True:
        # 计算到下次执行需要等待的秒数
        wait_seconds = (next_time - datetime.now()).total_seconds()

        if wait_seconds > 0:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] 等待执行...")
            time.sleep(wait_seconds)

        # 执行数据处理
        process_data()

        # 设置下一个执行时间为5分钟后
        next_time += timedelta(minutes=5)


def run_scheduler():
    """运行定时任务"""
    # 添加启动延迟（解决可能的初始化问题）
    time.sleep(5)

    print(f"===== 生产数据定时处理程序 =====")
    print(f"启动时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    try:
        # 检查工作目录是否存在
        input_dir = r'D:\czp\data\python_data'
        if not os.path.exists(input_dir):
            os.makedirs(input_dir)
            print(f"创建目录: {input_dir}")

        # 启动定时任务
        schedule_task()
    except Exception as e:
        print(f"严重错误: {str(e)}")
        print("程序将在10秒后退出...")
        time.sleep(10)


if __name__ == "__main__":
    # 在独立线程中运行定时器
    thread = threading.Thread(target=run_scheduler)
    thread.daemon = True  # 设置为守护线程
    thread.start()

    # 主线程保持活动状态
    while True:
        time.sleep(3600)  # 每小时检查一次