<%
Option Explicit
Dim value1,value2,value3,value4,objDB

value1=100   '������
value2="�ȳ��ϼ���"   '���ڿ�
value3=empty   '����
Set objDB = Server.CreateObject("ADODB.Connection")   '��ü�� ���� ����
%>  
value1=100<br>
value2="�ȳ��ϼ���"<br>
value3=empty<br>
Set objDB = Server.CreateObject("ADODB.Connection")<br>
<hr>
[1] IsEmpty(value1) : <%=IsEmpty(value1)%><br>
[2] IsEmpty(value2) : <%=IsEmpty(value2)%><br>
[3] IsEmpty(value3) : <%=IsEmpty(value3)%><br>
[4] IsEmpty(value4) : <%=IsEmpty(value4)%><br>
[5] value3�� ���� : <%=value3%><br>
[6] value4�� ���� : <%=value4%><br>
[7] value3 + value1 : <%=value3+value1%><br>
[8] IsNumeric(value1) : <%=IsNumeric(value1)%><br>
[9] IsObject(objDB) : <%=IsObject(objDB)%>
