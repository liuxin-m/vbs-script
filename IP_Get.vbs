Function GetIP
    GetIP = ""
    
    Dim objWMIService, colAdapters, objAdapter
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colAdapters = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    If colAdapters.Count = 0 Then
        Exit Function
    End If
    If Ubound(colAdapters.ItemIndex(0).IPAddress) = 0 Then
        Exit Function
    End If

    GetIP = colAdapters.ItemIndex(0).IPAddress(0)

End Function

ip = GetIP()

If Not ip = "" Then
    WScript.Echo ip
End If