<%
'--------------------------------------------------
' Title : MyBoard Education Version
' Program Name : board_function.asp
' Program Description : 게시판 관련 사용자 정의함수
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>
<%

'HTML 코드 사용 제한
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


'작은(홑)따옴표 처리
Function ConvertToSQL(strContent)
	Dim strTemp
	strTemp = strContent
	strTemp = Replace(strTemp, "'", "''")

	ConvertToSQL = strTemp

End Function

%>