<% '�Ű��������� �޾Ƽ� ó���ϴ� �Լ�
SUB PrintHTML2(printsize,printcolor) 
%>

     <HR>
     <font size = "<%=printsize%>" color="<%=printcolor%>">
     ����� ����� ���� �Լ� ��ºκ� - �Ű��������� ���� ���
     </font>
     <HR>

<% 
END SUB 
%>

<%
 str_size = 4
str_color = "red"
' CALL PrintHTML2(str_size,str_color)
PrintHTML2 str_size,str_color   '�Ű������� �ִ� ���� ��ȣ ����.
%>