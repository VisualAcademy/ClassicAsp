<%
Sub AdvancedPaging(PageNo, NumPage)'PageNo:����������, NumPage:��ü������
%>
	<!--���� 10��, ���� 10�� ����¡ ó�� ����-->
		<FONT style="font-size: 9pt;">
			<font color="#c0c0c0">[</font>
		<%	If PageNo > 10 Then %>
			<a href="./boardlist.asp?Page=<%= ((PageNo - 1) \ 10) * 10 %>">��</a>
		<%	Else %>
			<font color="#c0c0c0">��</font>
		<%	End If %>			
			<font color="#c0c0c0">|</font>
		<%	For i = (((PageNo - 1) \ 10) * 10 + 1) To ((((PageNo - 1) \ 10) + 1) * 10)
			If i > NumPage Then
				Exit For
			End If
			If i = Int(PageNo) Then
		%>
			<b><%= i %></b> <font color="#c0c0c0">|</font> 
		<%	Else %>
			<a href="./boardlist.asp?Page=<%= i %>"><%= i %></a> <font color="#c0c0c0">|</font> 
		<%	End If %>
		<%	Next %>
		<%	If CInt(i) < CInt(NumPage) Then %>
			<a href="./boardlist.asp?Page=<%= ((PageNo - 1) \ 10) * 10 + 11 %>">��</a>
		<%	Else %>
			<font color="#c0c0c0">��</font>
		<%	End If %>
			<font color="#c0c0c0">]</font>
		</FONT>
	<!--���� 10��, ���� 10�� ����¡ ó�� ����-->
<%
End Sub
%>