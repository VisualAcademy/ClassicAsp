<%
'--------------------------------------------------------------------
' Title : StrComp() ��� ����
' Program Name : strcomp�Լ�_���ڿ���.asp
' Program Description : ���δٸ� ���ڿ��� ũ�⸦ ���ϴ� �Լ�
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
	Response.Write(varA & "�� " & varB & " �� ���� �����ϴ�.<br>")
End If

varA = "123456"
varB = "654321"

If StrComp(varA, varB) = -1 Then
	Response.Write(varA & "�� " & varB & "���� �۽��ϴ�.<br>")
End If

varA = "654321"
varB = "123456"

If StrComp(varA, varB) = 1 Then
	Response.Write(varA & "�� " & varB & "���� Ů�ϴ�.<br>")
End If
%>
