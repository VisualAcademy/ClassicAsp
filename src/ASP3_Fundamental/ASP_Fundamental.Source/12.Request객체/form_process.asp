<p> 입력한 결과 페이지 </p>
<!-- post방식으로 전송시에는 Form컬렉션으로 받는다. -->
안녕하세요. <%= Request.Form("UserID") %> 님.<hr>

<!-- get방식으로 전송시에는 QueryString컬렉션으로 받는다. -->
사용자ID : <%= Request.QueryString("UserID") %> <br> 

<!-- Request객체로 직접 받는다(주로 사용) -->
비밀번호 : <%= Request("Password") %>
