<%
'변수 선언
Dim objCon: Dim strSql
Dim Name: Name = Request("Name")
Dim Email: Email = Request("Email")
Dim Title: Title = Request("Title")
Dim PostIP: PostIP = Request.ServerVariables("REMOTE_HOST")
Dim Content: Content = Request("Content")
Dim Password: Password = Request("Password")
Dim Encoding: Encoding = Request("Encoding")
Dim Homepage: Homepage = Request("Homepage")
Dim FileName: FileName = Request("FileName")
Dim FileSize: FileSize = 100
'[!] 작은따옴표 처리
Name = Replace(Name, "'", "''")
Title = Replace(Title, "'", "''")
Content = Replace(Content, "'", "''")
'커넥션 객체 생성
Set objCon = Server.CreateObject("ADODB.Connection")
'커넥션 객체 오픈
objCon.Open(Application("ConnectionString"))
'미리 실행할 SQL문 지정 : 순수한데이터 -> " & 변수 & "
strSql = "Insert Upload(Name,Email,Title,PostDate,PostIP,Content,Password,ReadCount,Encoding,Homepage,ModifyDate,ModifyIP,FileName,FileSize,DownCount) Values('" & Name & "','" & Email & "','" & Title & "',GetDate(),'" & PostIP & "','" & Content & "',	'" & Password & "',0,'" & Encoding & "','" & Homepage & "',Null,Null,'" & FileName & "'," & FileSize & ",0)"
'SQL문 실행
objCon.Execute(strSql)
'커넥션 객체 닫기
objCon.Close()
'커넥션 객체 해제
Set objCon = Nothing
'리스트 페이지로 이동
Response.Redirect("BoardList.asp")
%>
<center>데이터가 입력되었습니다.</center>