<% if request("userid") = "" then %>

<p> 폼 데이터 입력 </p>
<hr>

<form action="" method="post"> 
   사용자 ID : <input type="text" name="userid" size=20><br>
   비밀번호 : <input type="password" name="password" size=20><p>
   <input type="submit" value="보내기">
   <input type="reset" value="초기화">
</form>

<% else %>

<p> 입력한 결과 </p>
<hr>

안녕하세요. <%= Request("UserID") %> 님.<br>
사용자ID : <%= Request("UserID") %> <br> 
비밀번호 : <%= Request("Password") %>

<% end if %>
