<%
Response.Write "���� ��ư�� üũ �ڽ� �Է� �� ���� <hr>"
Response.Write Request("Radio") & "<br>"

For Each strCheck In Request("Check")
	Response.Write strCheck & "<br>"
Next    
%>