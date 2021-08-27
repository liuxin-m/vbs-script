Sub download(url,target)
	Const adTypeBinary = 1
	Const adSaveCreateOverWrite = 2
	Dim http,ado
	Set http = CreateObject("Msxml2.XMLHTTP")
	http.open "GET",url,False
	http.send
	Set ado = createobject("Adodb.Stream")
	ado.Type = adTypeBinary
	ado.Open
	ado.Write http.responseBody
	ado.SaveToFile target
	ado.Close
End Sub

''下载后存储文件位置''
Set ws = createobject("wscript.shell")
file = ws.expandenvironmentstrings("%TMP%") + "\file.txt" + 

download "http://abc.com/test.txt",file