Set Shell = CreateObject("Wscript.Shell")

Set objHtmlDoc = CreateObject("htmlfile")
'Text为剪切板内容
Text = objHtmlDoc.parentWindow.clipboardData.GetData("text")

'---------------自定义区域---------------
'打开软件
Path = "C:\Program Files\Everything\Everything.exe"

if IsNull(Text) then
'为空的时候替换会报错
else
Text=Replace(Text,"""", "")
end if

Key = " -s " & """" & Text & """"
'---------------自定义区域---------------

Order= """" & Path & """" & Key
Shell.Run(Order)
WScript.Quit