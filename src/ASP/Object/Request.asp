<%
' 아이디 받기
Dim UserID
Dim Password
UserID = Request.Form("UserID")
Password = Request("Password")
%>
넘겨져 온 아이디는 <%= UserID %>이고,
암호는 <%= Password %><br />
IP주소는 <%= Request.ServerVariables("REMOTE_ADDR") %><br />
IP주소는 <%= Request.ServerVariables("REMOTE_HOST") %><br />
