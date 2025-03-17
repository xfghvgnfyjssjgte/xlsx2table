import mysql.connector
import pandas as pd
import re
from datetime import datetime
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import json

class ProgressLogger:
    def __init__(self, total_steps=100):
        self.start_time = datetime.now()
        self.total_steps = total_steps
        self.current_step = 0

    def update(self, message, step_increment=0):
        self.current_step += step_increment
        elapsed = (datetime.now() - self.start_time).total_seconds()
        progress = self.current_step / self.total_steps
        remaining = elapsed / progress * (1 - progress) if progress > 0 else 0

        sys.stdout.write("\r\033[K")
        sys.stdout.write(
            f"[{datetime.now().strftime('%H:%M:%S')}] "
            f"进度: {self.current_step}/{self.total_steps} "
            f"({progress:.1%}) | 已用: {elapsed:.0f}s | 剩余: {remaining:.0f}s | {message} "
        )
        sys.stdout.flush()

    def complete(self):
        total_time = (datetime.now() - self.start_time).total_seconds()
        print(f"\n处理完成！总耗时: {total_time:.2f}秒")

def determine_column_type(column_values):
    """根据列数据推断数据类型"""
    column_values = column_values.astype(str).str.strip()
    column_values = column_values[column_values != ""]  # 排除空值

    if column_values.empty:
        return 'VARCHAR(255)'

    # 判断是否为日期类型
    for fmt in ('%Y-%m-%d', '%Y-%m-%d %H:%M:%S'):
        try:
            pd.to_datetime(column_values, format=fmt, errors='raise')
            return 'DATETIME' if ' ' in fmt else 'DATE'
        except Exception:
            continue

    # 判断是否为数值类型
    numeric_values = column_values.str.replace(',', '', regex=True)
    if numeric_values.str.match(r'^-?\d+$').all():
        return 'VARCHAR(255)'

    if numeric_values.str.match(r'^-?\d+(\.\d+)?$').all():
        return 'DECIMAL(18,6)'

    return 'VARCHAR(255)'

def excel2mariadb_with_progress(excel_path, username, password, host, database, port):
    try:
        plog = ProgressLogger(total_steps=8)
        plog.update("正在读取Excel文件...")

        # 根据文件扩展名选择引擎
        if excel_path.endswith('.xls'):
            engine_type = 'xlrd'
        elif excel_path.endswith('.xlsx'):
            engine_type = 'openpyxl'
        else:
            raise ValueError("不支持的文件类型，请提供 .xls 或 .xlsx 文件。")

        df = pd.read_excel(excel_path, dtype=str, keep_default_na=False, engine=engine_type)
        original_columns = [str(col).strip() for col in df.columns]
        plog.update(f"读取到 {len(df)} 行数据", step_increment=1)

        table_name = Path(excel_path).stem
        table_name = re.sub(r'[^a-zA-Z0-9_\u4e00-\u9fff]', '_', table_name)[:30]
        plog.update(f"表名生成完成：{table_name}", step_increment=1)

        plog.update("正在连接数据库...")
        conn = mysql.connector.connect(user=username, password=password, host=host, database=database, port=int(port))
        plog.update("数据库连接成功", step_increment=1)

        plog.update("正在创建数据表...")
        cursor = conn.cursor()
        cursor.execute(f'DROP TABLE IF EXISTS `{table_name}`')

        # 生成字段类型
        columns_definition = []
        for col in original_columns:
            column_type = determine_column_type(df[col])
            columns_definition.append(f'`{col}` {column_type}')

        ddl = f'CREATE TABLE `{table_name}` ({", ".join(columns_definition)})'
        cursor.execute(ddl)
        plog.update("数据表创建完成", step_increment=1)

        # 禁用索引和约束检查来提高性能
        cursor.execute("SET autocommit=0")
        cursor.execute("SET unique_checks=0")
        cursor.execute("SET foreign_key_checks=0")

        # 插入数据
        plog.update("开始插入数据...")
        total_rows = len(df)
        batch_size = 5000  # 增大批量大小以提高性能
        total_batches = (total_rows + batch_size - 1) // batch_size

        column_names = [f"`{col}`" for col in original_columns]
        placeholder = ", ".join(["%s"] * len(original_columns))
        insert_sql = f'INSERT INTO `{table_name}` ({", ".join(column_names)}) VALUES ({placeholder})'

        for batch_num in range(total_batches):
            start = batch_num * batch_size
            end = min(start + batch_size, total_rows)

            # 替换空字符串为空值（None）
            batch = [tuple(None if v == '' else v for v in row) for row in df.iloc[start:end].values]

            cursor.executemany(insert_sql, batch)
            conn.commit()

            progress_msg = f"已插入：{end}/{total_rows} ({end/total_rows:.1%})"
            plog.update(progress_msg)

        # 重新启用索引和约束检查
        cursor.execute("SET unique_checks=1")
        cursor.execute("SET foreign_key_checks=1")
        cursor.execute("SET autocommit=1")

        plog.complete()
    except mysql.connector.Error as err:
        messagebox.showerror("数据库连接错误", f"数据库连接失败: {err}")
        raise  # 重新抛出异常，以便在 submit 函数中处理
    except Exception as e:
        messagebox.showerror("错误", f"发生错误：{str(e)}")
        if 'conn' in locals():
            conn.rollback()
        raise
    finally:
        if 'conn' in locals():
            conn.close()

def open_file_dialog():
    """打开文件选择对话框"""
    filename = filedialog.askopenfilename(
        title="选择Excel文件",
        filetypes=[("Excel files", "*.xlsx")]
    )
    return filename

def load_config():
    """加载配置文件"""
    try:
        with open("config.json", "r") as f:
            config = json.load(f)
        return config
    except FileNotFoundError:
        return {}
    except json.JSONDecodeError:
        return {}

def save_config(config):
    """保存配置文件"""
    try:
        with open("config.json", "w") as f:
            json.dump(config, f, indent=4)
    except Exception as e:
        messagebox.showerror("错误", f"保存配置文件失败: {str(e)}")

def submit():
    """提交按钮的回调函数"""
    excel_file = file_path_var.get()
    host = host_var.get()
    username = username_var.get()
    password = password_var.get()
    database = database_var.get()
    port = port_var.get()

    if not excel_file:
        messagebox.showerror("错误", "请选择Excel文件")
        return

    config = {
        "host": host,
        "username": username,
        "password": password,
        "database": database,
        "port": port
    }
    save_config(config)

    try:
        excel2mariadb_with_progress(
            excel_path=excel_file,
            username=username,
            password=password,
            host=host,
            database=database,
            port=port
        )
        messagebox.showinfo("成功", "数据导入完成！")

    except Exception as e:
        messagebox.showerror("错误", f"发生错误: {str(e)}")



def browse_file():
    """浏览文件选择"""
    file_path = open_file_dialog()
    if file_path:
        file_path_var.set(file_path)

# 创建GUI
root = tk.Tk()
root.title("Excel 导入 MariaDB")

# 配置输入框和标签
file_path_var = tk.StringVar()
config = load_config()
host_var = tk.StringVar(value=config.get("host", "192.168.1.1"))
username_var = tk.StringVar(value=config.get("username", "root"))
password_var = tk.StringVar(value=config.get("password", "test"))
database_var = tk.StringVar(value=config.get("database", "test"))
port_var = tk.StringVar(value=config.get("port", "3306"))

tk.Label(root, text="Excel文件路径").grid(row=0, column=0, padx=10, pady=5)
tk.Entry(root, textvariable=file_path_var, width=40).grid(row=0, column=1, padx=10, pady=5)
tk.Button(root, text="浏览", command=browse_file).grid(row=0, column=2, padx=10, pady=5)

tk.Label(root, text="数据库地址").grid(row=1, column=0, padx=10, pady=5)
tk.Entry(root, textvariable=host_var, width=40).grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="用户名").grid(row=2, column=0, padx=10, pady=5)
tk.Entry(root, textvariable=username_var, width=40).grid(row=2, column=1, padx=10, pady=5)

tk.Label(root, text="密码").grid(row=3, column=0, padx=10, pady=5)
tk.Entry(root, textvariable=password_var, show="*", width=40).grid(row=3, column=1, padx=10, pady=5)

tk.Label(root, text="数据库").grid(row=4, column=0, padx=10, pady=5)
tk.Entry(root, textvariable=database_var, width=40).grid(row=4, column=1, padx=10, pady=5)

tk.Label(root, text="端口").grid(row=5, column=0, padx=10, pady=5)
tk.Entry(root, textvariable=port_var, width=40).grid(row=5, column=1, padx=10, pady=5)

tk.Button(root, text="开始导入", command=submit).grid(row=6, column=0, columnspan=3, padx=10, pady=20)

root.mainloop()