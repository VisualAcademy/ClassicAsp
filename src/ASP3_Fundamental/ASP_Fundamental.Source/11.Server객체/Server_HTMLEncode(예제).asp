<h4> Server 객체, HTMLEncode 메서드 </h4>
<hr>

<h5> 글자 색을 바꾸려면 ... </h5>
<p>
<% = Server.HTMLEncode("<FONT>") %> 태그에서 Color 속성을 바꾸어 줍니다. 
다음은 그 예를 보인 것입니다.
<p>
<% = Server.HTMLEncode("<font color=blue> 안녕하세요. </font>") %> <br>
-- <font color=blue> 안녕하세요. </font>

<h5> ScriptTimeout를 알아내려면 ... </h5>
<p>
<% = Server.HTMLEncode("<% ... %\>") %> 태그를 이용해서 브라우저로 보이는 방법입니다.
<p>
ScriptTimeout 시간은 <% = Server.HTMLEncode("<%=Server.ScriptTimeout %\>")%> 초 입니다. 
<br>
-- ScriptTimeout 시간은 <% = Server.ScriptTimeout %> 초 입니다.

