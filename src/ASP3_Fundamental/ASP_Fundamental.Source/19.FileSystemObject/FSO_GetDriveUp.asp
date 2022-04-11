<%
'파일 크기를 계산해서 알맞은 단위로 변환해줌. (바이트 수)
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

	Dim objFSO   '드라이브의 정보를 얻어올 객체의 인스턴스용 변수.
	Dim objCdrive   'C드라이브의 모든 정보를 얻어올 핸들링 객체변수.

	'FileSystemObject객체의 인스턴스 생성
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	'C드라이브의 핸들링 얻어오기
	Set objCdrive = objFSO.GetDrive("C:")
%>

[1] C드라이브의 전체 용량 : <%=ConvertToFileSize(objCdrive.TotalSize)%><br>
[2] C드라이브의 남은 용량 : <%=ConvertToFileSize(objCdrive.FreeSpace)%><br>
[3] C드라이브의 파일시스템 : <%=objCdrive.FileSystem%><br>


