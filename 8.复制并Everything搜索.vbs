Set Shell = CreateObject("Wscript.Shell")


'---------------自定义区域---------------
'模拟按键，Shift是+，Ctrl是^，Alt是%
Shell.SendKeys "^{c}"
Path = Shell.CurrentDirectory & "\Grid\Everything搜索.vbs"
Shell.Run("""" & Path & """")
'---------------自定义区域---------------

WScript.Quit