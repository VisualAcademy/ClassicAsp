<%@ Language=VBScript %>
<%
'-------------------------------------
' ��������� �Լ� ���� - �ۼ��� : �ڿ���
'-------------------------------------
%>

<% '�Լ� �⺻ ��� : ���� ���ν���
SUB Hi() 

	 Response.Write("�������ν��� ���")
	 
END SUB 
%>

<% '�Ű��������� �޾Ƽ� ó���ϴ� �Լ�
SUB PrintHTML2(printsize, printcolor) 
%>

	<font size = "<%=printsize%>" color="<%=printcolor%>">
	����� ����� ���� �Լ� ��ºκ� - �Ű��������� ���� ���
	</font>

<% 
END SUB 
%>

<% '���ϰ��� ��ȯ�����ִ� �Լ� : ���� ��ȯ��ų �� �Լ��̸��� ���� �ִ´�. : �Լ�
Function ConvertToWeekDay(str_week) 

     select case str_week
          case 1 : ConvertToWeekDay = "�Ͽ���"
          case 2 : ConvertToWeekDay = "������"
          case 3 : ConvertToWeekDay = "ȭ����"
          case 4 : ConvertToWeekDay = "������"
          case 5 : ConvertToWeekDay = "�����"
          case 6 : ConvertToWeekDay = "�ݿ���"
          case 7 : ConvertToWeekDay = "�����"
     end select
     
End Function
%>

����� ���� �Լ� ����
<HR>
<%
'[1] �Ű�����(�Ķ����) �� ���ϰ��� ���� �Լ� ȣ��
CALL Hi() 
'Hi()
'Hi
%>
<HR>
<%
'[2] �Ű������� �ִ� �Լ� ȣ��
str_size = 4
str_color = "red"
CALL PrintHTML2(str_size,str_color) 
%>
<HR>
<% 
'[3] �Ű����� �� ���ϰ��� �ִ� �Լ� ȣ��
str_week = WeekDay(Now)
Response.Write(ConvertToWeekDay(str_week)) 
%>
