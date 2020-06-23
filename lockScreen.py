# import sys
# from subprocess import Popen,PIPE,call
# process=Popen("pip list --outdated",stdout=PIPE,stdin=PIPE,shell=True,bufsize=1)
# data = process.stdout.readlines()
# for i in range(2,len(data)):
# 	temp=data[i].decode("utf-8").split(' ')
# 	print(temp[0])
# 	command="runas /user:administrator 'pip install {0} --upgrade'".format(temp[0])
# 	print(command)


# import win32com.client
# import sys
# from subprocess import Popen,PIPE
# import time
# from  win32gui import GetWindowText, GetForegroundWindow, SetForegroundWindow, EnumWindows
# from win32process import GetWindowThreadProcessId

# process=Popen("pip list --outdated",stdout=PIPE,stdin=PIPE,shell=True,bufsize=1)
# data= process.stdout.readlines()

# for i in range(2,len(data)):
# 	temp=data[i].decode("utf-8").split(' ')
# 	print(temp[0])
# 	command='runas /user:administrator "pip install {0} --upgrade"'.format(temp[0])
# 	print(command)
# 	shell = win32com.client.Dispatch("WScript.Shell")
# 	print(shell)
# 	shell.SendKeys(command+"{ENTER}")
# 	shell.SendKeys("161040107035{ENTER}")


import win32com.client
# import win32gui
# import win32process

# hwnd = win32gui.GetForegroundWindow()

# _, pid = win32process.GetWindowThreadProcessId(hwnd)
# print(pid)
shell = win32com.client.Dispatch("WScript.Shell")
shell.SendKeys('c:{ENTER}')
shell.SendKeys('cd C:\\Users\\nic\\AppData\\Local\\Packages\\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\\LocalState\\Assets {ENTER}')
shell.SendKeys('dir{ENTER}')
shell.SendKeys('copy *.* D:\\wallpaper{ENTER}')
shell.SendKeys('d:{ENTER}')
shell.SendKeys('cd D:\\wallpaper{ENTER}')
shell.SendKeys('ren *.* *.png{ENTER}')