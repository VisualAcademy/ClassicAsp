<%
'[1] 변수선언
Dim objCon
Dim strSql
Dim Name: Name = Request("Name")
Dim Email: Email = Request("Email")
Dim Title: Title = Request("Title")
Dim PostIP: PostIP = Request.ServerVariables("REMOTE_HOST")
Dim Content: Content = Request("Content")
Dim Password: Password = Request("Password")
Dim Encoding: Encoding = Request("Encoding")
Dim Homepage: Homepage = Request("Homepage")
'[!] 작은(홑)따옴표 처리 : '(작은따옴표 하나->''(작은따옴표 두개)
Name = Replace(Name, "'", "''")
Title = Replace(Title, "'", "''")
Content = Replace(Content, "'", "''")
'[2] 커넥션 객체 생성
Set objCon = Server.CreateObject("ADODB.Connection")
'[3] 데이터베이스 연결
objCon.Open("Provider=SQLOLEDB.1;Password=Basic;Persist Security Info=True;User ID=Basic;Initial Catalog=Basic;Data Source=(local)")
'[4] 미리 실행할 SQL문 지정
'홍길동 -> " & 대체할변수 & "
strSql = "Insert Basic(Name,Email,Title,PostDate,PostIP,Content,Password,ReadCount,Encoding,Homepage) Values('" & Name & "','" & Email & "','" & Title & "',GetDate(),'" & PostIP & "', '" & Content & "', '" & Password & "', 0,'" & Encoding & "', '" & Homepage & "')"
'[5] SQL문 실행
objCon.Execute(strSql)
'[6] 닫기
objCon.Close()
'[7] 해제
Set objCon = Nothing
'[8] 이동
Response.Redirect("./BoardList.asp")
%>
<center>데이터가 저장되었습니다.</center>