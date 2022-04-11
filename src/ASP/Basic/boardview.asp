<%
'--------------------------------------------------
' Title : Basic 보드
' Program Name : boardview.asp
' Program Description : 세부 글보기 페이지
' QueryString : 반드시 list.asp에서 Num=1 식으로 전송
' Include Files : None
' Copyright (C) 2004 Park Yong Jun
' E-mail: redplus@redplus.net
' Support: http://www.dotnetkorea.com/
'--------------------------------------------------
%>
<%
'[1] 변수 선언
Dim Num: Num = Request("Num")
Dim objCon: Dim objRs: Dim strSql
'[2] 커넥션 인스턴스
Set objCon = Server.CreateObject("ADODB.Connection")
'[3] 열기
objCon.Open(Application("ConnectionString"))
'[!] 조회수 증가 : 한번만 체크 -> 쿠키 적용
If Request.Cookies("ReadCount")(Num) = "" Then 
  strSql = "Update Basic Set ReadCount = ReadCount + 1 Where Num = " & Num
  objCon.Execute(strSql)
  Response.Cookies("ReadCount")(Num) = "OK"
End If
'[4] 레코드셋 인스턴스
Set objRs = Server.CreateObject("ADODB.RecordSet")
'[5] 레코드셋 열기 : 명령 실행(select)
strSql = "Select * From Basic Where Num = " & Num
objRs.Open strSql, objCon
'[6] 출력
If objRs.BOF Or objRs.EOF Then
	Response.Write("해당 레코드가 존재하지 않습니다.")
Else 'HTML+ASP코드(스파게티코드)로 모양 만들어 출력
	Call ShowRecordSet(objRs)
End If
'[7] 닫기
objRs.Close(): objCon.Close()
'[8] 해제
Set objRs = Nothing: Set objCon = Nothing
%>
<%
Sub ShowRecordSet(objRs)
%>

	<center><h3>기본 게시판 세부 항목 보기</h3></center>
	
	<center>
	<TABLE border=1 width=100%>
	<TR>
		<TD>글쓴이</TD>	<TD><%=objRs("Name")%></TD>
		<TD>날짜</TD>	<TD><%=objRs("PostDate")%></TD>
	</TR>
	<TR>
		<TD>Email</TD><TD><%=objRs("Email")%></TD>
		<TD>조회수</TD>	<TD><%=objRs("ReadCount")%></TD>
	</TR>
	<TR>
		<TD>제목</TD>	<TD colspan=3><%=objRs("Title")%></TD>
	</TR>
	<TR>
		<TD colspan=4><%=objRs("Content")%></TD>
	</TR>
	</TABLE>
	<a href="./boardlist.asp">[게시판리스트]</a>
	<a href="./boardmodify.asp?Num=<%=objRs("Num")%>">[수정]</a>
	<a href="./boarddelete.asp?Num=<%=objRs("Num")%>">[삭제]</a>
	</center>
<%
End Sub
%>