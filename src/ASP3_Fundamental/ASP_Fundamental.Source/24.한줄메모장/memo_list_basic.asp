<%
'--------------------------------------
' MyMemo Education Version
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------
%>
<%
Option Explicit

'변수 선언
Dim objDB, objRS, strSQL

'데이터베이스(connection객체)의 인스턴스 생성
Set objDB = Server.CreateObject("ADODB.Connection")

'데이터베이스 열기
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd;")

'레코드셋객체의 인스턴스 생성
Set objRS = Server.CreateObject("ADODB.RecordSet") 

'데이터베이스 처리(검색) 구문 작성
strSQL = "Select * From my_memo Order By Num Desc"

'레코드셋 열기
objRS.Open strSQL, objDB, 1

%>
<HTML>
	<HEAD>
		<META name="GENERATOR" content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
		<%
If objRS.BOF OR objRS.EOF Then
%>
		<P align="center">
			입력된 데이터가 없습니다.
		</P>
		<%
Else
%>
		<DIV align="center">
			<TABLE border="1">
				<TR>
					<TD align="center">
						번 호
					</TD>
					<TD align="center">
						이 름
					</TD>
					<TD align="center">
						이 메 일
					</TD>
					<TD align="center">
						메&nbsp;&nbsp;&nbsp;모
					</TD>
					<TD align="center">
						날 짜
					</TD>
				</TR>
				<%
	Do Until objRS.EOF   '데이터베이스 끝까지 값을 출력.
	%>
				<TR>
					<TD align="center">
						<%'=objRS.Fields("Num")%>
						<%'=objRS.Fields(0)%>
						<%=objRS("Num")%>
						<%'=objRS(0)%>
					</TD>
					<TD align="center">
						<%=objRS.Fields("Name")%>
					</TD>
					<TD align="center">
						<%=objRS.Fields("Email")%>
					</TD>
					<TD align="center">
						<%=objRS.Fields("Title")%>
					</TD>
					<TD align="center">
						<%=objRS.Fields("PostDate")%>
					</TD>
				</TR>
				<%
	objRS.MoveNext   '레코드를 다음 레코드로 이동시키는 레코드셋 메소드
	Loop
	%>
			</TABLE>
		</DIV>
		<%
End If
%>
	</BODY>
</HTML>
<%
'레코드셋 닫기
objRS.Close

'데이터베이스 닫기
objDB.Close

'레코드셋 개체의 인스턴스 해제
Set objRS = Nothing

'데이터베이스 객체의 인스턴스 해제
Set objDB = Nothing
%>
