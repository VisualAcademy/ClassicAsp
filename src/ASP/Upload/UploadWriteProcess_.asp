<%
Set objUploadForm = Server.CreateObject("SiteGalaxyUpload.Form")
%>
�Ѱ��� �� ��ü ��� <%= objUploadForm("FileName").FilePath  %><br />
�Ѱ��� �� ���� ũ�� <%= objUploadForm("FileName").Size %> Byte(s)<br />
���ϸ�
<%
FileName = objUploadForm("FileName").FilePath
FileName = Mid(FileName, InStrRev(FileName, "\") + 1)
Response.Write(FileName)
%><br />
���� ���ε�
<%
objUploadForm("FileName").SaveAs Server.MapPath(".") & "\files\" & FileName
%>
