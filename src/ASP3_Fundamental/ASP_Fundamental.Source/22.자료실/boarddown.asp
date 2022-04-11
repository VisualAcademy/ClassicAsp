<%

Call Main()

Sub Main()

	strFileName = Request("strFileName")

	If strFileName = "" Then
		Call RedirFail()
		Exit Sub	
	Else
		Call RedirectFile(strFileName)
	End If

End Sub


Sub RedirectFile(strFileName)

	'서버에 저장될 파일의 위치와 이름
	strFilePath = Server.MapPath(".") + "\files\" & strFileName

	Set objFSO = Server.CreateObject("SiteGalaxyUpload.FileSystemObject")
	Set objFile = objFSO.OpenBinaryFile(strFilePath, 1, False)
	
	Response.Clear
	Response.ContentType = "application/octet-stream"
	Response.AddHeader "Content-disposition", "attachment; filename=" & strFileName
	Response.BinaryWrite objFile.ReadAll
	Response.End
	
	Set objFile = Nothing
	Set objFSO = Nothing

End Sub


Sub RedirFail()

	Response.Redirect("./boardlist.asp")

End Sub

%>