<%
	Option Explicit
	Dim intSince
	'�����Լ� : ��¥ ���� �Լ�
	'Now(), Date(), Time()
	Response.Write Now()   '2001-08-21 ���� 6:16:33 �� ���� ����.
	Response.Write("<br>")
	Response.Write Now   '2001-08-21 ���� 6:16:33 �� ���� ����.
	Response.Write("<br>")
	Response.Write Date()   '2001-08-21
	Response.Write("<br>")
	Response.Write Time()   '���� 6:17:53
	Response.Write("<br>")

	'��¥ ���� �Լ�
	'Year(), Month(), Day(), WeekDay(), Hour(), Minute(), Second()
	Response.Write(Year(Now()))   '��
	Response.Write("<br>")
	Response.Write(Month(Now()))   '��
	Response.Write("<br>")
	Response.Write(Day(Now()))   '��
	Response.Write("<br>")
	Response.Write(WeekDay(Now()))   '����(1:�Ͽ��� ~ 7:�����)
	Response.Write("<br>")
	Response.Write(Hour(Now()))   '�ð�
	Response.Write("<br>")
	Response.Write(Minute(Now()))   '��
	Response.Write("<br>")
	Response.Write(Second(Now()))   '��
	Response.Write("<br>")

	'��¥�� ���� ���ڰ��� ��ȯ(DateSerial()).
	intSince = Date() - DateSerial(2002,01,01)
	Response.Write("���ذ� ���� " & intSince &" ���� �������ϴ�.")

   'ũ�������� ������ ��ĥ�� ���Ҵ�����???
	intSince = DateSerial(2002,12,25) - Date()
	Response.Write("<br>")
	Response.Write("ũ�������� ������ " & intSince &"���� ���ҽ��ϴ�.")

%>
	FormatDateTime(Date,1)�� ��� : <%=FormatDateTime(Date,1)%><br>
	FormatDateTime(Date,2)�� ��� : <%=FormatDateTime(Date,2)%><br>
	FormatDateTime(Date,3)�� ��� : <%=FormatDateTime(Date,3)%><br>
	FormatDateTime(Date,4)�� ��� : <%=FormatDateTime(Date,4)%><p>
