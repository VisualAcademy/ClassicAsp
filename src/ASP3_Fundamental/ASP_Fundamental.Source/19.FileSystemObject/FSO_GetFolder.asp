<%
	'Option Explicit

	Dim objFSO   '����̺��� ������ ���� ��ü�� �ν��Ͻ��� ����.
	Dim objFolder   'Ư�� ������ ��� ������ ���� �ڵ鸵 ��ü����.

	'FileSystemObject��ü�� �ν��Ͻ� ����
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	'C:\InetPub�� �ڵ鸵 ������
	Set objFolder = objFSO.GetFolder("C:\InetPub")
%>

[1] C:\InetPub�� �뷮 : <%=objFolder.Size%><br>


