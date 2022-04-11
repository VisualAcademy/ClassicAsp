<%
'파일 크기를 계산해서 알맞은 단위로 변환해줌. (바이트 수)
Function ConvertToFileSize(intByte)

	intFileSize = Int(intByte)
	
	If intFileSize >= 1048576 Then
		ConvertToFileSize = FormatNumber((intByte / 1048576), 2) & " MB"
	Else
		If intFileSize >= 1024 Then
			ConvertToFileSize = FormatNumber((intByte / 1024), 0) & " KB"
		Else
			ConvertToFileSize = intByte & " Byte(s)"
		End If
	End If

End Function
%>

<%=ConvertToFileSize(2345678)%><br>
<%=ConvertToFileSize(23456)%><br>
<%=ConvertToFileSize(234)%><br>
<%=ConvertToFileSize(23494344032)%>