<%
'--------------------------------------------------
' Title : MyBoard Education Version
' Program Name : board_function.asp
' Program Description : �Խ��� ���� ����� �����Լ�
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>
<%

'HTML �ڵ� ��� ����
Function ConvertToHTML(ByVal strContent)

	Dim strTemp
	
	If Len(strContent) = 0 Or IsNull(Len(strContent)) = True Then
		strTemp = ""
	Else
		strTemp = strContent
		strTemp = Replace(strTemp, "&", "&amp;")
		strTemp = Replace(strTemp, ">", "&gt;")
		strTemp = Replace(strTemp, "<", "&lt;")
	End If

	ConvertToHTML = strTemp

End Function


'����(Ȭ)����ǥ ó��
Function ConvertToSQL(strContent)
	Dim strTemp
	strTemp = strContent
	strTemp = Replace(strTemp, "'", "''")

	ConvertToSQL = strTemp

End Function

%>