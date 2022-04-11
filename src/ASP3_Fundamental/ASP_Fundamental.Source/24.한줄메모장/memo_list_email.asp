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
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd")

'레코드셋객체의 인스턴스 생성
Set objRS = Server.CreateObject("ADODB.RecordSet") 

'데이터베이스 처리(검색) 구문 작성
strSQL = "Select * From my_memo Order By Num Desc"

'레코드셋 열기
objRS.Open strSQL, objDB, 1

%>
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
		<%
If objRS.BOF or objRS.EOF then
%>
		<p align="center">
			입력된 데이터가 없습니다.
		</p>
		<%
Else
%>
		<div align="center">
			<table border="1">
				<tr>
					<td align="center">
						번 호
					</td>
					<td align="center">
						이 름
					</td>
					<td align="center">
						메&nbsp;&nbsp;&nbsp;모
					</td>
					<td align="center">
						날 짜
					</td>
				</tr>
				<%
	Do Until objRS.EOF   '데이터베이스 끝까지 값을 출력.
	%>
				<tr>
					<td align="center">
						<%=objRS.Fields("Num")%>
					</td>
					<td align="center">
						<a href="mailto:<%=objRS.Fields("Email")%>">
							<%=objRS.Fields("Name")%>
						</a>
					</td>
					<td align="center">
						<%=objRS.Fields("Title")%>
					</td>
					<td align="center">
						<%=objRS.Fields("PostDate")%>
					</td>
				</tr>
				<%
	objRS.MoveNext
	Loop
	%>
			</table>
		</div>
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
