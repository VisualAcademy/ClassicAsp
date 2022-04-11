<% If Request("userid") = "" Then %>

전송할 데이터 입력<br>
<hr>
<form action="" method="post"> 
   사용자 ID : <input type="text" name="UserID" size=20><br>
   비밀번호 : <input type="password" name="Password" size=20><p>
   <input type="submit" value="보내기">
   <input type="reset" value="초기화">
</form>

<% Else %>

전송된 데이터 출력<br>
<hr>
안녕하세요. <%= Request.Form("UserID") %> 님.<br>
사용자ID : <%= Request.QueryString("UserID") %> <br> 
비밀번호 : <%= Request("Password") %>

<% End If %>