''UNIX时间戳获取''

Set objWMIService = _
GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery _
("Select * from Win32_OperatingSystem",,48)
For Each objItem in colItems
	TimeZone = objItem.CurrentTimeZone
Next
UnixTime = DateDiff("s", "01/01/1970 00:00:00", Now())
UnixTime = UnixTime - TimeZone * 60
