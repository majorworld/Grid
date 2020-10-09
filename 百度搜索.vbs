Set Shell = CreateObject("Wscript.Shell")

Set objHtmlDoc = CreateObject("htmlfile")
'Text为剪切板内容
Text = objHtmlDoc.parentWindow.clipboardData.GetData("text")

'---------------自定义区域---------------
'打开网站
Path = "https://www.baidu.com/s?wd="
Key = Text
'---------------自定义区域---------------

Order= """" & Path & Key & """"
Shell.Run(Order)
WScript.Quit