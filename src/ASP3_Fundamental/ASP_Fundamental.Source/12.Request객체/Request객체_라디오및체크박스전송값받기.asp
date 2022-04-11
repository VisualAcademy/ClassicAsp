<%
Response.Write "라디오 버튼과 체크 박스 입력 폼 예제 <hr>"
Response.Write Request("Radio") & "<br>"

For Each strCheck In Request("Check")
	Response.Write strCheck & "<br>"
Next    
%>