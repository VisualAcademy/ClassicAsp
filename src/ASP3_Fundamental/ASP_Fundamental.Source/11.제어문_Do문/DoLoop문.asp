<!-- while ... wend �ݺ��� -->

<%
	count = 0
	
	while  count < 10

		response.write "+" & count
		count = count + 1

	wend
%>
<p>

<!-- do while .. loop �ݺ��� -->

<%
	count = 0
	
	do while  count < 10

		response.write "+" & count
		count = count + 1

	loop
%>
<p>

<!-- do .. loop while �ݺ��� -->

<%
	count = 0
	
	do 

		response.write "+" & count
		count = count + 1

	loop while count < 10
%>
<p>

<!-- do until .. loop  �ݺ��� -->

<%
	count = 0
	
	do until count = 10

		response.write "+" & count
		count = count + 1

	loop 
%>
<p>

<!-- do .. loop until �ݺ��� -->

<%
	count = 0
	
	do 

		response.write "+" & count
		count = count + 1

	loop until count = 10
%>
<p>