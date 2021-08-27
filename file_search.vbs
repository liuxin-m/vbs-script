'用于遍历本地一个目录下的所有文件'

Function GetIP
'获取本地IP'
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

local_IP = GetIP
WScript.Echo local_IP
File_name = "c:\" + local_IP + ".txt"
'打开一个文件，用于存储信息'
Set oFs =createobject("scripting.filesystemobject")
Set ts=oFs.opentextfile(File_name,8,true)

Function FilesTree(sPath)    
'遍历一个文件夹下的所有文件夹文件夹    
    'Set oFso = CreateObject("Scripting.FileSystemObject")    
    Set oFolder = oFs.GetFolder(sPath)    
    Set oSubFolders = oFolder.SubFolders    
        
		
		
    Set oFiles = oFolder.Files    
    For Each oFile In oFiles    
        'WScript.Echo oFile.Path    
		ts.writeline(oFile.Path)
        'oFile.Delete'    
    Next    
        
    For Each oSubFolder In oSubFolders    
        'WScript.Echo oSubFolder.Path'    
        'oSubFolder.Delete'
		'ts.writeline oSubFolder.Path
        FilesTree(oSubFolder.Path)'递归    
    Next    
        
    Set oFolder = Nothing    
    Set oSubFolders = Nothing    
    Set oFso = Nothing    
End Function    


FilesTree("E:\test") '遍历
ts.close


