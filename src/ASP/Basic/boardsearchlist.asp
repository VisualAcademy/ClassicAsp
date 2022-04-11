<%
'--------------------------------------------------
' Title : Basic 보드
' Program Name : boardsearchlist.asp
' Program Description : 검색 결과 리스트 페이지
' QueryString : 검색폼에서 SearchField, SearchQuery
' Copyright (C) 2004 Park Yong Jun
' E-mail: redplus@redplus.net
' Support: http://www.dotnetkorea.com/
'--------------------------------------------------
%>
<%
'[1] 변수 선언
Option Explicit
Dim SearchField: SearchField = Request("SearchField")
Dim SearchQuery: SearchQuery = Request("SearchQuery")
Dim objCon: Dim objRs: Dim strSql
'[2] 커넥션 객체의 인스턴스
Set objCon = Server.CreateObject("ADODB.Connection")
'[3] Open() : 커넥션 스트링 지정
objCon.Open(Application("ConnectionString"))
'[4] 레코드셋 객체의 인스턴스
Set objRs = Server.CreateObject("ADODB.RecordSet")
'[5] Open() : 명령 실행(Select)
strSql = "Select * From Basic Where " & _
		 SearchField & " Like '%" & _
		 SearchQuery & "%' Order By Num Desc"
objRs.Open strSql, objCon
'[6] 출력
If objRs.BOF Or objRs.EOF Then
	Response.Write("<b style='color:red;'>" & SearchQuery & "</b>로 검색한 결과가 없습니다.")
Else
	Call ShowRecordSet(objRs)
End If
'[7] Close()
objRs.Close(): objCon.Close()
'[8] 객체 해제
Set objRs = Nothing: Set objCon = Nothing
%>
<%
Sub ShowRecordSet(objRs)
%>
	<table border=1 width="100%">
<%
	While Not objRs.EOF
%>
	<TR onmouseover="this.style.backgroundColor='red';" onmouseout="this.style.backgroundColor='';">
		<TD><%=objRs("Num")%></TD>
		<TD><a href="./boardview.asp?Num=<%=objRs("Num")%>"><%=objRs("Title")%></a></TD>
		<TD><%=objRs("Name")%></TD>
		<TD><%=objRs("Email")%></TD>
		<TD><%=objRs("ReadCount")%></TD>
	</TR>
<%
		objRs.MoveNext()
	WEnd
%>
	<table>
<%
End Sub
%>
<a href="./boardlist.asp">검색종료</a>