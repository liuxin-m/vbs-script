'开启和关闭本地服务'

Set ServiceSet = GetObject("winmgmts:").ExecQuery("select * from Win32_Service where Name='PolicyAgent'")

For each Service in ServiceSet
	RetVal = Service.StartService()
If RetVal = 0 Then 
	WScript.Echo "Service started"
If RetVal = 10 Then 
	WScript.Echo "Service already running"
Next

'以下为停止的代码:

Set ServiceSet = GetObject("winmgmts:").ExecQuery("select * from Win32_Service where Name='PolicyAgent'")

For each Service in ServiceSet
RetVal = Service.StopService()
If RetVal = 0 then
	WScript.Echo "Service stopped"
ElseIf RetVal = 5 then
	WScript.Echo "Service already stopped"
End If
Next