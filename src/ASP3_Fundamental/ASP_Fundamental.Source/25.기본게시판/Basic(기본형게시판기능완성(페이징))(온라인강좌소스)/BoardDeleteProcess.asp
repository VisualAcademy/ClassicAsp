<%
Dim objCon: Dim objCmd: Dim objRs
Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=Basic;Persist Security Info=True;User ID=Basic;Initial Catalog=Basic;Data Source=(local)")
	Set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = objCon
	'//�Ѱ��� �� ��ȣ�� ��й�ȣ�� �´� �����Ͱ� �ִ��� �˻�
	objCmd.CommandText = "Select Password From Basic Where Num = " & Request("Num") & " And Password = '" & Request("Password") & "'"
	    Set objRs = objCmd.Execute()	
		If objRs.BOF Or objRs.EOF Then	'�´� �����Ͱ� ���ٸ�...
%>
		<SCRIPT LANGUAGE="JavaScript">
		window.alert("��й�ȣ�� ���� �ʽ��ϴ�.");
		history.go(-1);
		</SCRIPT>
<%
		Else	'�´� �����Ͱ� �ִٸ� ���� �� ����Ʈ�� �̵�
			objCmd.CommandText = "Delete Basic Where Num = " & Request("Num")
			objCmd.Execute()'//���� ��ɽ���
			Response.Redirect("./BoardList.asp")'//����Ʈ�� �̵�
		End If	
		objRs.Close()
		Set objRs = Nothing
	Set objCmd = Nothing
objCon.Close()
Set objCon = Nothing
%>