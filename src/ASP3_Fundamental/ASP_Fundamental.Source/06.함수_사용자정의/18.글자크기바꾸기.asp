<%
Sub ConvertFont(strColor, intSize)
%>

<font color="<%=strColor%>" size="<%=intSize%>">
글자 모양이 바뀔까여?
</font>

<%
End Sub
%>

<%

Call ConvertFont("red", 6)

%>


