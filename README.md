# Excel2MariaDB-Importer

一个用于将 Excel 文件数据导入 MariaDB 数据库的工具。

## 功能特性

* 支持 `.xls` 和 `.xlsx` 文件格式。
* 图形化用户界面，操作简单。
* 自动推断数据类型并创建 MariaDB 表。
* 显示导入进度。
* 支持配置文件。
* 错误提示。
* 支持配置MariaDB端口。

## 依赖

* Python 3.x
* pandas
* mysql-connector-python
* openpyxl
* xlrd
* tkinter

## 安装

1.  确保已安装 Python 3.x。
2.  使用 pip 安装依赖：

    ```bash
    pip install pandas mysql-connector-python openpyxl xlrd
    ```

## 使用

1.  下载或克隆本项目。
2.  运行 `xlsx2table.py`。
3.  在图形界面中：
    * 点击“浏览”选择 Excel 文件。
    * 输入 MariaDB 数据库连接信息（地址、用户名、密码、数据库名、端口）。
    * 点击“开始导入”。
4.  程序会自动创建数据表并导入数据，并在完成后显示提示信息。

## 配置文件

* 数据库连接信息会自动保存到 `config.json` 文件中，下次启动程序时会自动加载。

## 错误处理

* 如果数据库连接失败，程序会弹出错误信息框。
* 如果数据导入过程中出现错误，程序会显示错误信息。

## 贡献

欢迎提交 issue 和 pull request。

## 许可证

本项目使用 MIT 许可证。

## 注意事项

* 请确保 MariaDB 数据库服务已启动。
* 请确保输入的数据库连接信息正确。
* 请确保 Excel 文件格式正确。
