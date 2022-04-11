
<table border = 1>

<%	
For Each key in Request.ServerVariables
%>
	<tr>
	<td>
		<%
		Response.Write key 
		%>
	</td>
	<td>
	<%
	If Request.ServerVariables(key) <> "" Then
		Response.Write Request.ServerVariables(key)
	Else
		Response.Write "&nbsp;"
	End If
	%>
	</td>
	</tr>
<%
Next 
%>

</table>