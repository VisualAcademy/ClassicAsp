<html>
<body bgcolor=White>

<% if request("mode") = "out" then %>
<p> 입력한 결과 </p>
<hr>

안녕하세요. <%= Request("UserID") %> 님.<br>
사용자ID : <%= Request("UserID") %> <br> 
비밀번호 : <%= Request("Password") %>

<% else %>

<p> 폼 데이터 입력 </p>
<hr>

<form action="Request_Hidden_Form.asp" method="Get"> 

   사용자 ID : <input type="text" name="userid" size=20><br>
   비밀번호 : <input type="password" name="password" size=20><p>

   <input type="submit" value="보내기">
   <input type="reset" value="초기화">

   <input type="hidden" name="mode" value="out">
</form>

<% end if %>
</body>    
</html>

