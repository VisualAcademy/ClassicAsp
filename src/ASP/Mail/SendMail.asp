<%@ Language=VBScript %>
<% Option Explicit %>
<%
	Dim receiver, sender, subject, content
	Dim remember, bodyFormat, file, path
	Dim i, cdo, galaxy

	Set galaxy = Server.CreateObject("SiteGalaxyUpload.Form")

	receiver = Split(galaxy("tb_receiver"), ";")
	sender = galaxy("tb_sender")
	remember = galaxy("cb_remember")
	bodyFormat = galaxy("radio_bodyformat")
	subject = galaxy("tb_subject")
	content = galaxy("tb_content")
	file = galaxy("tb_attach")

	If (Len(remember) > 0) Then
		Response.Cookies("mail_remember_me") = sender
		Response.Cookies("mail_remember_me").Expires = DateAdd("m", 1, Now())
	Else
		Response.Cookies("mail_remember_me").Expires = DateAdd("m", -1, Now())
	End If

	Set cdo = Server.CreateObject("CDO.Message")
	cdo.From = sender
	cdo.Subject = subject

	If Len(file) > 0 Then
		file = Mid(file, InStrRev(file, "\") + 1)
		path = Server.MapPath("Upload/" & file)
		galaxy("tb_attach").SaveAs path
		cdo.AddAttachment(path)
	End If

	If CInt(bodyFormat) = 0 Then
		cdo.HtmlBody = content
	Else
		cdo.TextBody = content
	End If

	For i = 0 To UBound(receiver)
		cdo.To = receiver(i)
		cdo.Send
	Next

	Response.Write "<script language='javascript'>"
	Response.Write "	alert('메일을 모두 전송했습니다.');"
	Response.Write "	location.href='Default.asp';"
	Response.Write "</script>"

	Set cdo = Nothing
	Set galaxy = Nothing
%>