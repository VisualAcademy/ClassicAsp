<%
'���� ���� �� ������Ʈ�� �ޱ�
Dim objCon: Dim strSql: Dim objRs
Dim Num: Num = Request("Num")
Dim Password: Password = Request("Password")
'Ŀ�ؼ� ����
Set objCon = Server.CreateObject("ADODB.Connection")
'����
objCon.Open(Application("ConnectionString"))
	'SQL�� ���� : �Ѱ��� �� ��ȣ�� �ش��ϴ� �н�����/��ȣ���� 
	strSql = "Select * From Upload Where Num = " & Num & _
			 " And Password = '" & Password & "'"
	'���ڵ�� ��ü�� �ޱ� : ������ ���ÿ� ��� ����
	Set objRs = objCon.Execute(strSql)'//objRs.Open()
	If objRs.BOF Or objRS.EOF Then
		'�ڷ� ������ : �ڹٽ�ũ��Ʈ ���
%>
		<script language="javascript">
		window.alert("��ŷ�Ϸ���?");
		window.history.go(-1);
		</script>
<%		
		'Response.End()
	Else
	
'-----------------------------------------------
'�̹� ���ε�� ������ ���� �����ϴ� ����
'-----------------------------------------------
Call DeleteArticleNow()
Sub DeleteArticleNow()
	Dim objFSO, objFile	'//FSO��ü�� ���� ����
	If objRS.Fields("FileName") <> "" Then '//���ε�� ������ �ִٸ�
		'//FSO��ü�� �ν��Ͻ� ����
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		'//��������� files�������ִ� ���Ͽ� ���� ������ ���ͼ�...
		Set objFile = objFSO.GetFile(Server.MapPath(".") + "\files\" & objRS.Fields("FileName"))
		'//�ش� ������ ����
		objFile.Delete
		'//��ü ����
		Set objFile = Nothing
		Set objFSO = Nothing
	End If
End Sub
'----------------------------------------------

		'���� ó�� �� ����Ʈ�� �̵�
		strSql = "Delete Upload Where Num = " & Num
		objCon.Execute(strSql)
		Response.Redirect("BoardList.asp")
	End If
	'���ڵ��(�޸𸮻��� ���̺� ���) �ݱ�
	objRs.Close()
	'���ڵ�� ��ü ����
	Set objRs = Nothing
'�ݱ�
objCon.Close()
'Ŀ�ؼ� ����
Set objCon = Nothing
'Response.Redirect("BoardList.asp")
%>








