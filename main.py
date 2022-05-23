import time,os,sys,traceback
import win32api
import win32con
import win32clipboard
from ctypes import windll
from selenium import webdriver
import mainwindow
from PyQt5.QtWidgets import QWidget,QApplication,QMainWindow,QFileDialog
from qt_material import apply_stylesheet

file_path = None

def get_file_path():
    global file_path
    try:
        file_path = QFileDialog.getExistingDirectory()
        ui.lineEdit_filepath.setText(file_path)
    except:
        print(traceback.format_exc())

def process_image():
    options = webdriver.ChromeOptions()
    options.add_argument('--save-page-as-mhtml')
    driver = webdriver.Chrome(r'C:\Users\nj36dh\Downloads\chromedriver_win32\chromedriver.exe', chrome_options=options)
    # driver = webdriver.Chrome(r'C:\Users\nj36dh\Downloads\chromedriver_win32\chromedriver.exe')    # Chrome浏览器
    url = 'http://mkweb.bcgsc.ca/color-summarizer/?analyze'
    # 打开网页
    driver.get(url)  # 打开url网页 比如 driver.get("http://www.baidu.com")

    for item in os.listdir(file_path):
        driver.find_element_by_xpath('//input[@value="2"]').click()
        driver.find_element_by_xpath('//input[@value="vhigh"]').click()
        btn = driver.find_element_by_name('image')
        btn.send_keys(os.path.join(file_path,item))
        time.sleep(1)
        sub = driver.find_element_by_name('.submit')
        sub.click()
        time.sleep(1)
        title = item
        # 设置路径为：当前项目的绝对路径+文件名
        path = (os.path.join(r'D:\mhtml',title + ".mhtml"))

        # 将路径复制到剪切板
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(path)
        win32clipboard.CloseClipboard()

        # 按下ctrl+s
        win32api.keybd_event(0x11, 0, 0, 0)
        win32api.keybd_event(0x53, 0, 0, 0)
        win32api.keybd_event(0x53, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(1)

        # 鼠标定位输入框并点击
        windll.user32.SetCursorPos(700, 510)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(1)

        # 按下ctrl+a
        win32api.keybd_event(0x11, 0, 0, 0)
        win32api.keybd_event(0x41, 0, 0, 0)
        win32api.keybd_event(0x41, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(1)

        # 按下ctrl+v
        win32api.keybd_event(0x11, 0, 0, 0)
        win32api.keybd_event(0x56, 0, 0, 0)
        win32api.keybd_event(0x56, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(1)

        # 按下回车
        win32api.keybd_event(0x0D, 0, 0, 0)
        win32api.keybd_event(0x0D, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(1)

        driver.get(url)

    driver.close()



def init_ui(ui):
    ui.pushButton_path.clicked.connect(lambda :get_file_path())
    ui.pushButton_processimage.clicked.connect(lambda :process_image())

if __name__ == '__main__':
    app = QApplication(sys.argv)

    ui = mainwindow.Ui_MainWindow()
    main_window = QMainWindow()
    ui.setupUi(main_window)
    apply_stylesheet(app, theme='dark_teal.xml')
    init_ui(ui)
    main_window.show()

    sys.exit(app.exec_())