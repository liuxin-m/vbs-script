Function GetFileHash(file_name)
	Dim file_hash
	Dim hash_value
	Dim i
	Set wi = CreateObject("WindowsInstaller.Installer")
	Set file_hash = wi.FileHash(file_name, 0)
	hash_value = ""
	For i = 1 To file_hash.FieldCount
		hash_value = hash_value & BigEndianHex(file_hash.IntegerData(i))
	Next
	GetFileHash = LCase(hash_value)
	Set file_hash = Nothing
End Function

Function BigEndianHex(Int)
	Dim result
	Dim b1, b2, b3, b4
	result = Hex(Int) 
	Sub_result = 8 - Len(result) '不足8位的需要前面补0，否则计算md5错误
	If Sub_result > 0 Then
		For i = 1 To Sub_result
			result = "0" + result
		Next
	End If
	b1 = Mid(result, 7, 2)
	b2 = Mid(result, 5, 2)
	b3 = Mid(result, 3, 2)
	b4 = Mid(result, 1, 2)
	BigEndianHex = b1 & b2 & b3 & b4
End Function

Set Shell=CreateObject("wscript.shell")
file=Shell.expandenvironmentstrings("%tmp%") + "\file.txt"

Md5 = GetFileHash(file)
MsgBox Md5