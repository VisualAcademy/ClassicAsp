<%
'[1] 변수 선언
Option Explicit
Dim objCon
Dim strSql
Dim objRs
'[2] 데이터베이스의 인스턴스 생성 : Connection 오브젝트
Set objCon = Server.CreateObject("ADODB.Connection")
'[3] 데이터베이스 열기
objCon.Open("Provider=SQLOLEDB.1;Password=Memo;Persist Security Info=True;User ID=Memo;Initial Catalog=Memo;Data Source=(local)")
'[4] SQL문 실행
strSql = "Select * From Memo"
Set objRs = objCon.Execute(strSql)
If objRs.BOF Or objRs.EOF Then
	Response.Write("입력된 데이터가 없습니다.")
Else
	Call ShowRecordSet(objRs)
End If
'[5] 데이터베이스 닫기
objCon.Close()
'[6] 데이터베이스의 인스턴스 해제
Set objCon = Nothing 
%>
<%
Sub ShowRecordSet(objRs)
	While Not objRs.EOF
		Response.Write(objRs.Fields("Name") & Space(1) & _
		objRs("Title") & Space(1) & _
		objRs.Fields(4) & "<br>")
		objRs.MoveNext()
	WEnd
End Sub 
%>

