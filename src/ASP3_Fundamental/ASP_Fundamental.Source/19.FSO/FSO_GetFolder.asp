<%
	'Option Explicit

	Dim objFso   '����̺��� ������ ���� ��ü�� �ν��Ͻ��� ����.
	Dim objFolder   'Ư�� ������ ��� ������ ���� �ڵ鸵 ��ü����.

	'FileSystemObject��ü�� �ν��Ͻ� ����
	Set objFso = Server.CreateObject("Scripting.FileSystemObject")

	'C:\InetPub�� �ڵ鸵 ������
	'Set objFolder = objFso.GetFolder("C:\InetPub")
	Set objFolder = objFso.GetFolder("C:\Temp")
%>
[1] C:\Temp�� �뷮 : <%=objFolder.Size%><br>
<%
Set objFolder = Nothing
Set objFso = Nothing
%>