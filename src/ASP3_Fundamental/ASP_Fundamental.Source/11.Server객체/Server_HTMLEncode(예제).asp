<h4> Server ��ü, HTMLEncode �޼��� </h4>
<hr>

<h5> ���� ���� �ٲٷ��� ... </h5>
<p>
<% = Server.HTMLEncode("<FONT>") %> �±׿��� Color �Ӽ��� �ٲپ� �ݴϴ�. 
������ �� ���� ���� ���Դϴ�.
<p>
<% = Server.HTMLEncode("<font color=blue> �ȳ��ϼ���. </font>") %> <br>
-- <font color=blue> �ȳ��ϼ���. </font>

<h5> ScriptTimeout�� �˾Ƴ����� ... </h5>
<p>
<% = Server.HTMLEncode("<% ... %\>") %> �±׸� �̿��ؼ� �������� ���̴� ����Դϴ�.
<p>
ScriptTimeout �ð��� <% = Server.HTMLEncode("<%=Server.ScriptTimeout %\>")%> �� �Դϴ�. 
<br>
-- ScriptTimeout �ð��� <% = Server.ScriptTimeout %> �� �Դϴ�.

