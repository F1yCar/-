import sys
import traceback
import os
import datetime
from gui import MainWindow
from PyQt5.QtWidgets import QApplication

def create_log(message):
    log_dir = r"E:\test"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(log_dir, f"app_log_{timestamp}.txt")

    with open(log_file, "w", encoding="utf-8") as f:
        f.write(message)

    print(f"日志已保存到: {log_file}")

def run_app():
    create_log("程序启动\n")
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    exit_code = app.exec_()
    create_log(f"程序正常退出，退出代码: {exit_code}\n")
    return exit_code

if __name__ == "__main__":
    try:
        sys.exit(run_app())
    except Exception as e:
        error_message = f"发生错误:\n{traceback.format_exc()}"
        print(error_message)
        create_log(error_message)
        input("按回车键退出...")
