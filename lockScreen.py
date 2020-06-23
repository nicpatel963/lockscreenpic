
import win32com.client

shell = win32com.client.Dispatch("WScript.Shell")
shell.SendKeys('c:{ENTER}')
shell.SendKeys('cd C:\\Users\\nic\\AppData\\Local\\Packages\\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\\LocalState\\Assets {ENTER}')
# change the username from 'nic' to yours.
shell.SendKeys('dir{ENTER}')
shell.SendKeys('copy *.* D:\\wallpaper{ENTER}')
shell.SendKeys('d:{ENTER}')
shell.SendKeys('cd D:\\wallpaper{ENTER}')
shell.SendKeys('ren *.* *.png{ENTER}')
