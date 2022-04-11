<html>
<body bgColor=#ffffff>

<% If Request("userid") = "" then %>
<p> 폼 데이터 입력 </p>
<hr>

<form action="form_return.asp" method="post"> 
   사용자 ID : <input type="text" name="userid" size=20><br>
   비밀번호 : <input type="password" name="password" size=20><p>
   <input type="submit" value="보내기">
   <input type="reset" value="초기화">
</form>

<% Else %>
<p> 입력한 결과 </p>
<hr>

안녕하세요. <%= Request("UserID") %> 님.<br>
사용자ID : <%= Request("UserID") %> <br> 
비밀번호 : <%= Request("Password") %><br>
<%
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(Server.MapPath(".")+"\password.dat",1)

If Request("UserID") = Trim(objFile.ReadLine) AND Request("Password") = Trim(objFile.ReadLine) Then
	Session("id") = "Admin"
Else
	Session("id") = "Guest"
End IF

If Session("id") = "Admin" Then
	Response.Write("관리자군요...")
Else
	Response.Write("지금 손님으로 접속중입니다.")	
End If
%>
<% End if %>
</body>    
</html>
