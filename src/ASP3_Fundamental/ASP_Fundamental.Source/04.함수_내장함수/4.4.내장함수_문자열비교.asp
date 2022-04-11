<%
'--------------------------------------------------------------------
' Title : StrComp() 사용 예제
' Program Name : strcomp함수_문자열비교.asp
' Program Description : 서로다른 문자열의 크기를 비교하는 함수
' Include Files : None
' Last Update : 2001.10.25. 04:30
' Copyright (C) 2001 Park Yong Jun
' E-mail: redplus@redplus.net
' Support: http://www.redplus.net/
'--------------------------------------------------------------------
%>
<%
Dim varA
Dim varB

varA = "123456"
varB = "123456"

If StrComp(varA, varB) = 0 Then
	Response.Write(varA & "와 " & varB & " 두 값이 같습니다.<br>")
End If

varA = "123456"
varB = "654321"

If StrComp(varA, varB) = -1 Then
	Response.Write(varA & "가 " & varB & "보다 작습니다.<br>")
End If

varA = "654321"
varB = "123456"

If StrComp(varA, varB) = 1 Then
	Response.Write(varA & "가 " & varB & "보다 큽니다.<br>")
End If
%>
