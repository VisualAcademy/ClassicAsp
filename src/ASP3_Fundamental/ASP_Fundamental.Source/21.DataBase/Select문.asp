<%
	'[1] ������ ����� ����
	Option Explicit

	'[2] ���� ����
	Dim objDB, objRS, strSQL

	'[3] �����ͺ��̽��� �ν��Ͻ�(��ü)�� ����
	Set objDB = Server.CreateObject("ADODB.Connection")

	'[4] �����ͺ��̽� ����(Open() �޼ҵ� ��� : DSN, UID, PWD)
	objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd;")

	'[5] ���ڵ���� �ν��Ͻ� ����
	Set objRS = Server.CreateObject("ADODB.RecordSet")

	'[6] SQL�� �ۼ�
	strSQL = "Select * From my_memo Order By Num Desc"

	'[7] ���ڵ�� ����(Open()�޼ҵ� : SQL��, DB������(Ŀ�ؼǰ�ü), Ŀ��Ÿ��, ��Ÿ��, �ɼ�)
	objRS.Open strSQL, objDB

	'[8] ���ڵ� �� ó�� ���μ��� ����
	If objRS.BOF OR objRS.EOF Then
		Response.Write("<center>�Էµ� �ڷᰡ �����ϴ�.</center>")
	Else
		Response.Write("<div align='center'>")
		Response.Write("<table border='1'>")
		Response.Write("<tr><td>��ȣ</td><td>�̸�</td><td>�̸���</td><td>����</td><td>�ۼ���</td></tr>")
		
		Do Until objRS.EOF
			Response.Write("<tr>")
			Response.Write("<td>" & objRS.Fields("Num") & "</td>")
			Response.Write("<td>" & objRS.Fields("Name") & "</td>")
			Response.Write("<td>" & objRS.Fields("Email") & "</td>")
			Response.Write("<td>" & objRS.Fields("Title") & "</td>")
			Response.Write("<td>" & objRS.Fields("PostDate") & "</td>")
			Response.Write("</tr>")
			objRS.MoveNext	
		Loop

		Response.Write("</table>")
		Response.Write("</div>")
	End If

	'[9] ���ڵ�� �ݱ�
	objRS.Close

	'[10] �����ͺ��̽� �ݱ�
	objDB.Close

	'[11] ���ڵ�� ��ü�� �ν��Ͻ� ����
	Set objRS = Nothing

	'[12] �����ͺ��̽� ��ü�� �ν��Ͻ� ����
	Set objDB = Nothing
%>


