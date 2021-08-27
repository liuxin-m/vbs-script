'==========================================================================
'
' COMMENT:判断是否存在一个文件，如果存在，则删除，如果不存在，则建立 
'
'==========================================================================

Function IsExitAFile(filespec)
        Dim fso
        Set fso=CreateObject("Scripting.FileSystemObject")        
        If fso.fileExists(filespec) Then         
        IsExitAFile=True        
        Else IsExitAFile=False        
        End If
End Function 

Sub CreateAFile(filespec)
        Dim fso
        Set fso=CreateObject("Scripting.FileSystemObject")
        fso.CreateTextFile(filespec)
End Sub

Sub DeleteAFile(filespec)
        Dim fso
        Set fso= CreateObject("Scripting.FileSystemObject")
        fso.DeleteFile(filespec)
End Sub

If IsExitAFile("D:\\test.tst") Then
	DeleteAFile("D:\\test.tst")
Else 
	CreateAFile("D:\\test.tst")
End If