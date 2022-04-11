<%
Set objUploadForm = Server.CreateObject("SiteGalaxyUpload.Form")
%>
넘겨져 온 전체 경로 <%= objUploadForm("FileName").FilePath  %><br />
넘겨져 온 파일 크기 <%= objUploadForm("FileName").Size %> Byte(s)<br />
파일명
<%
FileName = objUploadForm("FileName").FilePath
FileName = Mid(FileName, InStrRev(FileName, "\") + 1)
Response.Write(FileName)
%><br />
실제 업로드
<%
objUploadForm("FileName").SaveAs Server.MapPath(".") & "\files\" & FileName
%>
