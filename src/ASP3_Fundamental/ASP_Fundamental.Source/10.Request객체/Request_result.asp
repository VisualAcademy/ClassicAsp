<HTML>
<BODY>
<font face="돋움" size="2">
<P>&nbsp;</P>
<center>
<h2>Request로 받은 값들</h2>
<p>&nbsp;</P>

이름은 <%=Request.Form("username")%> 입니다.<p>
나이는 <%=Request.Form("age")%> 입니다.<p>
전화번호는 <%=Request("tel")%> 입니다.<p>

'Request
'.form()  -- 컬렉션...=(배열) : 키값을 사용해서 값을 반환...(post방식으로 전송시)
'.QueryString() : get방식으로 전송시(<A>)
'Request객체는 두개의 컬렉션을 사용하지않고 바로 정보를 받을 수 있다.
- Request("넘겨온이름")

</center></font>
</BODY>
</HTML>
