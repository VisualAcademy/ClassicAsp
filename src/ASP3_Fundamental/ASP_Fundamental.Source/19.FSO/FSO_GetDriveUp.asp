<%
'���� ũ�⸦ ����ؼ� �˸��� ������ ��ȯ����. (����Ʈ ��)
Function ConvertToFileSize(intByte)

	intFileSize = Int(intByte)
	If intFileSize >= 1073741824 Then
		ConvertToFileSize = FormatNumber((intByte / 1073741824), 2) & " GB"
	Else
		If intFileSize >= 1048576 Then
			ConvertToFileSize = FormatNumber((intByte / 1048576), 2) & " MB"
		Else
			If intFileSize >= 1024 Then
				ConvertToFileSize = FormatNumber((intByte / 1024), 0) & " KB"
			Else
				ConvertToFileSize = intByte & " Byte(s)"
			End If
		End If
	End If

End Function
%>
<%
	'Option Explicit

	Dim objFSO   '����̺��� ������ ���� ��ü�� �ν��Ͻ��� ����.
	Dim objCdrive   'C����̺��� ��� ������ ���� �ڵ鸵 ��ü����.

	'FileSystemObject��ü�� �ν��Ͻ� ����
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	'C����̺��� �ڵ鸵 ������
	Set objCdrive = objFSO.GetDrive("C:")
%>

[1] C����̺��� ��ü �뷮 : <%=ConvertToFileSize(objCdrive.TotalSize)%><br>
[2] C����̺��� ���� �뷮 : <%=ConvertToFileSize(objCdrive.FreeSpace)%><br>
[3] C����̺��� ���Ͻý��� : <%=objCdrive.FileSystem%><br>


