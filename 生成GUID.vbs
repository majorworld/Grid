Set Shell = CreateObject("Wscript.Shell")

'生成GUID
Set TypeLib = CreateObject("Scriptlet.TypeLib")
GUID = Left(TypeLib.Guid,38)
'保存并打开
Path = ".\生成GUID.txt"
Set FSO = CreateObject("Scripting.FileSystemObject")

Set File = FSO.CreateTextFile(Path, true)

File.WriteLine GUID
Shell.Run(Path)
'退出
WScript.Quit