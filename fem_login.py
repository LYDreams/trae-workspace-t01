import os
import time
from pywinauto import Application
from pywinauto.keyboard import send_keys

print("FEM软件自动登录脚本启动")

# 获取桌面路径
desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')

# 查找FEM软件
fem_shortcut = None
for file in os.listdir(desktop_path):
    if 'FEM' in file:
        fem_shortcut = os.path.join(desktop_path, file)
        break

if not fem_shortcut:
    print("未找到FEM软件快捷方式")
    exit()

print(f"找到FEM软件: {fem_shortcut}")

# 启动FEM软件
print("正在启动FEM软件...")
os.startfile(fem_shortcut)

# 等待应用程序启动和登录窗口出现
time.sleep(5)

# 连接到已启动的应用程序 - 使用win32后端更可靠
print("正在连接到登录窗口...")
app = Application(backend='win32').connect(title_re='.*Login Window.*', timeout=60)

# 获取登录窗口
login_window = app.window(title_re='.*Login Window.*')
login_window.wait('visible', timeout=5)
print("登录窗口已出现")

# 确保窗口置顶
login_window.set_focus()
time.sleep(1)

# 方式1：使用纯键盘模拟
print("\n方法1：使用纯键盘模拟...")
try:
    login_window.set_focus()
    time.sleep(1)
    
    # 清空用户名框（假设当前有焦点）
    send_keys('{DELETE}')
    send_keys('{BACKSPACE}')
    time.sleep(0.5)
    
    # 输入用户名
    send_keys('231599')
    print("用户名输入成功")
    time.sleep(1)
    
    # 切换到密码框
    send_keys('{TAB}')
    time.sleep(0.5)
    
    # 清空密码框
    send_keys('{DELETE}')
    send_keys('{BACKSPACE}')
    time.sleep(0.5)
    
    # 输入密码
    send_keys('1')
    print("密码输入成功")
    time.sleep(1)
    
    # 切换到OK按钮并点击
    send_keys('{TAB}')
    time.sleep(0.5)
    send_keys('{ENTER}')
    print("OK按钮点击成功")
except Exception as e:
    print(f"登录过程出错：{e}")
    print("尝试方法2：纯键盘模拟...")
    # 方式2：最可靠的纯键盘模拟，假设焦点在用户名框
    login_window.set_focus()
    time.sleep(1)
    
    # 输入用户名
    send_keys('231599')
    time.sleep(0.5)
    
    # 切换到密码框
    send_keys('{TAB}')
    time.sleep(0.5)
    
    # 输入密码
    send_keys('1')
    time.sleep(0.5)
    
    # 点击OK
    send_keys('{TAB}')
    time.sleep(0.5)
    send_keys('{ENTER}')
    print("纯键盘模拟登录成功")

print("\n登录操作已完成")
print("脚本执行结束")