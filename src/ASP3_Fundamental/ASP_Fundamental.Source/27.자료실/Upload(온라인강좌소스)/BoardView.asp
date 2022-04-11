<%
'변수
Dim objCon: Dim objCmd: Dim objRs
'커넥션
Set objCon = Server.CreateObject("ADODB.Connection")
'오픈
objCon.Open(Application("CONNECTION_STRING"))
	'커멘드
	Set objCmd = Server.CreateObject("ADODB.Command")
	'액티브커넥션
	objCmd.ActiveConnection = objCon
	'커멘드 텍스트 1
	objCmd.CommandText = "Update Upload Set ReadCount = ReadCount + 1 Where Num = " & Request("Num")
	'명령 실행
	objCmd.Execute()
		'커멘드 텍스트 2
		objCmd.CommandText = "Select * From Upload Where Num = " & Request("Num")
		'명령 실행 및 레코드셋 반환
		Set objRs = objCmd.Execute()
			'모양 만들어 출력
			Call ShowRecordSet(objRs)
		'레코드셋 닫기
		objRs.Close()
		'레코드셋 해제
		Set objRs = Nothing
	'해제
	Set objCmd = Nothing
'닫기
objCon.Close()
'해제
Set objCon = Nothing
%>
<%
Sub ShowRecordSet(objRs)
%>
	<table border="1" width="100%">
<%
	If objRs.BOF Or objRs.EOF Then
%>
	<tr><td>입력된 데이터가 없습니다.</td></tr>
<%		
	Else
%>
	<tr><td>제목 : </td><td><%=objRs("Title")%></td></tr>
	<tr><td>작성자 : </td><td><%=objRs("Name")%></td></tr>
	<tr><td>작성일 : </td><td><%=objRs("PostDate")%></td></tr>
	<tr><td colspan="2">
	<%
	Dim strContent: strContent = objRs("Content")
	'Text로 저장
	If objRs("Encoding") = "Text" Then
		'HTML 태그 제한
		strContent = Replace(strContent, "&", "&amp;")
		strContent = Replace(strContent, "<", "&lt;")
		strContent = Replace(strContent, ">", "&gt;")
		'엔터 처리
		strContent = Replace(strContent, Chr(13) & Chr(10), "<br>")
		'strContent = Replace(strContent, vbCrLf, "<br>")
	ElseIf objRs("Encoding") = "Mixed" Then
		'엔터 처리
		strContent = Replace(strContent, Chr(13) & Chr(10), "<br>")
	Else
		'//strContent = objRs("Content")
	End If	
	'출력
	Response.Write(strContent)
	%>	
	</td></tr>
<%
	End If
%>
	<tr><td colspan="2" align="center"><a href="BoardList.asp">리스트</a></td></tr>
	</table>
<%
End Sub
%>














