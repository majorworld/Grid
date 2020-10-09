Set Shell = CreateObject("Wscript.Shell")


data = inputbox("输入时间，单位为秒，0为取消关机","倒计时关机",0)
if data > 0 then
'---------------自定义区域---------------
'启动软件
Order = "cmd.exe /c shutdown -s -f -t " & data
Shell.Run Order , 0
'---------------自定义区域---------------
else 
Shell.Run("cmd.exe /c shutdown -a")
end if

WScript.Quit