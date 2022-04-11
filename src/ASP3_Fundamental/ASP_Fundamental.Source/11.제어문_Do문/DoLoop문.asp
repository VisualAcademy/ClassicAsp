<!-- while ... wend 반복문 -->

<%
	count = 0
	
	while  count < 10

		response.write "+" & count
		count = count + 1

	wend
%>
<p>

<!-- do while .. loop 반복문 -->

<%
	count = 0
	
	do while  count < 10

		response.write "+" & count
		count = count + 1

	loop
%>
<p>

<!-- do .. loop while 반복문 -->

<%
	count = 0
	
	do 

		response.write "+" & count
		count = count + 1

	loop while count < 10
%>
<p>

<!-- do until .. loop  반복문 -->

<%
	count = 0
	
	do until count = 10

		response.write "+" & count
		count = count + 1

	loop 
%>
<p>

<!-- do .. loop until 반복문 -->

<%
	count = 0
	
	do 

		response.write "+" & count
		count = count + 1

	loop until count = 10
%>
<p>