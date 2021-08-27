''获取指定进程的PID''
Set w = GetObject("winmgmts:")
procname = "wechat.exe"
Set p = w.ExecQuery("select * from win32_process where name='" + procname + "' ")
if p.Count = 0 then
	msgbox "指定进程未运行或用户权限不足以获得其信息。"
else
For Each i In p
	msgbox "进程 " & i.name & " 的 PID 是 " & i.ProcessId
Next
end if