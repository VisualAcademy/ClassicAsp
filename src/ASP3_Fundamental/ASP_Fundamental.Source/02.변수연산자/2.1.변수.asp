<%
	'[1] ���� ����� ���� : ���� ���������� ������ �ϳ��� �������� �ʰ� ����ϸ� ������ �߻���Ų��.
	'�ּ��� ������ ù�ٿ� �;��Ѵ�.
	Option Explicit

	'[2] ���� ���� : Dim�̶�� Ű���带 ����ؼ� �����Ѵ�.
	'�밡���� ǥ��� �Ǵ� �Ľ�Įǥ����� ����ؼ� �������� ���Ѵ�.
	Dim strHi   '������ ����...(String)
	Dim intSu   '������ ����...(Integer)
	Dim objDB   '��ü�� ����...(Object)
	Dim dateBirthday   '��¥�� ����...
	Dim strSQL, strDB   '������ ������ ���ÿ� �����Ϸ��� ,(��ǥ)�� ����ؼ� �����Ѵ�.

	'[3] ���� �� ����
	strHi= "�ȳ��ϼ���"   'ASP���� ���ڿ� ������ "(ū����ǥ)�� ���´�.
	intSu = 100
	dateBirthday = #08-21-2001#   '��¥���� �����Ϸ��� ��¥ ���ڿ� �յڷ� #(Sharp) ��ȣ�� ���δ�.

	'[4] ���� �� ���
	Response.Write("strHi : " & strHi & "<br>" )
	Response.Write "intSu+intSu : " & intSu+intSu & "<br>"
	Response.Write "dateBirthday : " & dateBirthday & "<br>"
	
	'[5] ���� ���� �� ���(ASP�����Լ��� TypeName() ��� : �ش� ������ ������ ǥ���Ѵ�.)
	Response.Write(strHi & ": &nbsp;" & TypeName(strHi) & "<br>")
	Response.Write(intSu & ":&nbsp;" & TypeName(intSu+intSu) & "<br>")
	Response.Write(dateBirthday & ":&nbsp;" & TypeName(dateBirthday))
%>