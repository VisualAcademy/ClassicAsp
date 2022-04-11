<%
Dim objCon: Dim objRs: Dim strSql
'[!]페이징 처리 관련
Dim intPage: intPage = Request("Page")
If intPage = "" Then	'/넘겨져 온 페이지 값이 없으면 1페이지부터 출력
	intPage = 1
End If
'[!]이전/다음
Dim intPageCount: intPageCount = 0 '//총 페이지 수
Set objCon = Server.CreateObject("ADODB.Connection")
	objCon.Open("Provider=SQLOLEDB.1;Password=Basic;Persist Security Info=True;User ID=Basic;Initial Catalog=Basic;Data Source=(local)")
		strSql = "Select * From Basic Order By Num Desc"
		Set objRs = Server.CreateObject("ADODB.RecordSet")
			'[!] 한 페이지에 5개씩 보여주기
			objRs.PageSize = 5
			'[!] 페이징 처리시 반드시 3번째 인자(커서타입) : 1 또는 3
			objRs.Open strSql, objCon, 1
				'출력
				If objRs.BOF Or objRs.EOF Then
				Else
					'[!]이전/다음
					intPageCount = objRs.PageCount
					'[!] 그럼, 끊어져 읽어온 것 중 몇번째부터 보여줄건지
					objRs.AbsolutePage = intPage
					Dim i: i = 1	'//초기식
					Do Until objRs.EOF Or i > objRs.PageSize '//조건식
						Response.Write(objRs.Fields("Num") & "<br>")
						objRs.MoveNext()
						i = i + 1 '//증감식
					Loop
				End If
			objRs.Close()
		Set objRs = Nothing
	objCon.Close()
Set objCon = Nothing
%>

<hr>

<%
Call AdvancedPaging(intPage, intPageCount)'/고급 페이징 함수 호출
Sub AdvancedPaging(PageNo, NumPage)'PageNo:현재페이지, NumPage:전체페이지
%>
<!--이전 10개, 다음 10개 페이징 처리 시작-->
	<FONT style="font-size: 9pt;">
	   <font color="#c0c0c0">[</font>
	<%     If PageNo > 10 Then %>
	   <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?Page=<%= ((PageNo - 1) \ 10) * 10 %>">◀</a>
	<%     Else %>
	   <font color="#c0c0c0">◁</font>
	<%     End If %>               
	   <font color="#c0c0c0">|</font>
	<%     For i = (((PageNo - 1) \ 10) * 10 + 1) To ((((PageNo - 1) \ 10) + 1) * 10)
	   If i > NumPage Then
			Exit For
	   End If
	   If i = Int(PageNo) Then
	%>
	   <b><%= i %></b> <font color="#c0c0c0">|</font> 
	<%     Else %>
	   <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?Page=<%= i %>"><%= i %></a> <font color="#c0c0c0">|</font> 
	<%     End If %>
	<%     Next %>
	<%     If CInt(i) < CInt(NumPage) Then %>
	   <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?Page=<%= ((PageNo - 1) \ 10) * 10 + 11 %>">▶</a>
	<%     Else %>
	   <font color="#c0c0c0">▷</font>
	<%     End If %>
	   <font color="#c0c0c0">]</font>
	</FONT>
<!--이전 10개, 다음 10개 페이징 처리 종료-->
<%
End Sub
%> 


<hr>

<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post">
페이지 : 
<input type="text" name="Page" value="<%=intPage%>" size="3" style="text-align:center;"> 
<input type="submit" value="이동">
</form>

<hr>

<%
If intPage > 1 Then
%>
	<a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?Page=<%=intPage-1%>">
	[이전]
	</a>
<%
Else
%>
	<span style="color:gray;">[이전]</span>
<%
End If
%>
(현재:<%=intPage%>/총:<%=intPageCount%>) 
<%
If CInt(intPage) <> CInt(intPageCount) Then
%>
	<a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?Page=<%=intPage+1%>">
	[다음]
	</a>
<%
Else
%>
	<span style="color:gray;">[다음]</span>
<%
End If
%>