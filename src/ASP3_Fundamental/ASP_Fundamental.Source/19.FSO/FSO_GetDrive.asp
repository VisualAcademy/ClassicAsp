<%
'[0] ��������
Option Explicit
Dim objFSO   '����̺��� ������ ���� ��ü�� �ν��Ͻ��� ����.
Dim objCdrive   'C����̺��� ��� ������ ���� �ڵ鸵 ��ü����.

'[1] ��ü ����
'FileSystemObject��ü�� �ν��Ͻ� ���� : FSO��ü�� objFSO�� �̸����� ����ϰڴٸ� ����.
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

'[2] �ڵ鸵 ������
'C����̺��� �ڵ鸵 ������ : C����̺��� ������ �޾Ƽ� 
'objCDrive�� �̸����� ����ϰڴٸ� ����.
Set objCDrive = objFSO.GetDrive("C:\")

'[3] ���α׷� ó��
%>
[1] C����̺��� ��ü �뷮 : <%=objCdrive.TotalSize%><br>
[2] C����̺��� ���� �뷮 : <%=objCdrive.FreeSpace%><br>
[3] C����̺��� ���Ͻý��� : <%=objCdrive.FileSystem%><br>
<%
'[4] ��ü ����
Set objCDrive = Nothing
Set objFSO = Nothing
%>