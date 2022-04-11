<%
Dim objCon: Dim objRs: Dim strSql
'[!]페이징 처리 관련
Dim intPage: intPage = Request("Page")
If intPage = "" Then	'/넘겨져 온 페이지 값이 없으면 1페이지부터 출력
	intPage = 1
End If
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

<form action="페이징연습.asp" method="post">
페이지 : 
<input type="text" name="Page" size="3" style="text-align:center;"> 
<input type="submit" value="이동">
</form>
