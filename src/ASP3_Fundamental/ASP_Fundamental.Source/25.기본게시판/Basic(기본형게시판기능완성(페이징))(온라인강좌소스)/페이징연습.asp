<%
Dim objCon: Dim objRs: Dim strSql
'[!]����¡ ó�� ����
Dim intPage: intPage = Request("Page")
If intPage = "" Then	'/�Ѱ��� �� ������ ���� ������ 1���������� ���
	intPage = 1
End If
Set objCon = Server.CreateObject("ADODB.Connection")
	objCon.Open("Provider=SQLOLEDB.1;Password=Basic;Persist Security Info=True;User ID=Basic;Initial Catalog=Basic;Data Source=(local)")
		strSql = "Select * From Basic Order By Num Desc"
		Set objRs = Server.CreateObject("ADODB.RecordSet")
			'[!] �� �������� 5���� �����ֱ�
			objRs.PageSize = 5
			'[!] ����¡ ó���� �ݵ�� 3��° ����(Ŀ��Ÿ��) : 1 �Ǵ� 3
			objRs.Open strSql, objCon, 1
				'���
				If objRs.BOF Or objRs.EOF Then
				Else
					'[!] �׷�, ������ �о�� �� �� ���°���� �����ٰ���
					objRs.AbsolutePage = intPage
					Dim i: i = 1	'//�ʱ��
					Do Until objRs.EOF Or i > objRs.PageSize '//���ǽ�
						Response.Write(objRs.Fields("Num") & "<br>")
						objRs.MoveNext()
						i = i + 1 '//������
					Loop
				End If
			objRs.Close()
		Set objRs = Nothing
	objCon.Close()
Set objCon = Nothing
%>

<form action="����¡����.asp" method="post">
������ : 
<input type="text" name="Page" size="3" style="text-align:center;"> 
<input type="submit" value="�̵�">
</form>
