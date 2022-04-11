<%
Dim objCon: Dim objCmd: Dim objRs
Dim intNum: intNum = Request("Num")
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=Basic;Persist Security Info=True;User ID=Basic;Initial Catalog=Basic;Data Source=(local)")
	Set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = objCon
	'//조회수 증가
	objCmd.CommandText = "Update Basic Set ReadCount = ReadCount + 1 Where Num = " & intNum
	objCmd.Execute()
	'//조회수 증가 끝
	' 순수데이터 -> " & 변수 & "
	objCmd.CommandText = "Select * From Basic Where Num = " & intNum 
		Set objRs = objCmd.Execute()
			Call ShowRecordSet(objRs)
		objRs.Close()
		Set objRs = Nothing
	Set objCmd = Nothing
objCon.Close()
Set objCon = Nothing
%>
<%
Sub ShowRecordSet(objRs)'// 모양 출력 전담 함수
			If objRs.BOF Or objRs.EOF Then
			Else
				Do Until objRs.EOF
%>	
	<TABLE border="1" width="600" align="center">
	<TR>
		<TD>번호 : </TD><TD><%=objRs.Fields("Num")%></TD>
	</TR>
	<TR>
		<TD>제목 : </TD><TD><%=objRs.Fields("Title")%></TD>
	</TR>
	<TR>
		<TD>작성자 : </TD><TD><%=objRs.Fields("Name")%></TD>
	</TR>
	<TR>
		<TD>작성일 : </TD><TD><%=objRs.Fields("PostDate")%></TD>
	</TR>
	<TR>
		<TD>조회수 : </TD><TD><%=objRs.Fields("ReadCount")%></TD>
	</TR>
	<TR>
		<TD colspan="2" height="200" valign="top">
<%	
	Dim strContent
	'[!] HTML 태그 처리
	strContent = objRs.Fields("Content")
	If objRs.Fields("Encoding") = "Text" Then ' Text
		strContent = Replace(strContent, "&", "&amp;")
		strContent = Replace(strContent, "<", "&lt;")
		strContent = Replace(strContent, ">", "&gt;")
		'[!] 엔터 처리
		strContent = Replace(strContent,Chr(13) & Chr(10),"<br>")
	ElseIf objRs("Encoding") = "Mixed" Then ' Mixed
		'[!] 엔터 처리
		strContent = Replace(strContent,Chr(13) & Chr(10),"<br>")
	End If
	Response.Write(strContent)
%>
		</TD>
	</TR>
	<TR>
		<TD colspan="2" align="right">
			<a href="./BoardModify.asp?Num=<%=objRs("Num")%>">[수정]</a>
			<a href="./BoardDelete.asp?Num=<%=objRs("Num")%>">[삭제]</a>
			<a href="./BoardList.asp">[리스트]</a>
		</TD>
	</TR>
	</TABLE>				
<%
					objRs.MoveNext()
				Loop
			End If
End Sub
%>
