'禁用本地IPSec服务'

Dim OperationRegistry
Set OperationRegistry=WScript.CreateObject("WScript.Shell") 
OperationRegistry.RegWrite "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\PolicyAgent\Start",4,"REG_DWORD"
Set ServiceSet = GetObject("winmgmts:").ExecQuery("select * from Win32_Service where Name='PolicyAgent'")
for each Service in ServiceSet
	RetVal = Service.StopService()
next