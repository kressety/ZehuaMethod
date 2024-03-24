# Zehua Method

环境：

- Windows
- Microsoft Office

# 使用

前往 `Release` 下载打包好的文件，将 `0323.xlsx` 和 `REPAIR_RPT_P1_MS-158X_0313.xlsx` 文件放入 `data` 文件夹中。

1. 运行 `main.exe`，程序将在 `output` 文件夹下生成一个 `processed.csv` 文件；
2. 根据 `processed.csv` 文件的内容进行一些详细处理，根据提示继续操作 `main.exe`；

> 注意：程序会自动保存进度，中途不用了可以直接退出；但不同机器的处理速度不同，务必等机器处理好当前任务后再退出程序，否则可能不同程度上遗失进度。

# 开发

1. 克隆代码
   `git clone https://github.com/kressety/ZehuaMethod.git; cd ZahuaMethod`
2. 安装依赖
   `pip install -r requirements.txt`
3. 运行
   `python main.py`
