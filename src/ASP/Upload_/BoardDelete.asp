<script language="javascript">
function DeleteCheck(){
	var valPassword = 
		window.document.DeleteForm.Password.value;
	if(valPassword == ""){
		window.alert("��й�ȣ�� �־��ּ���.");
		window.document.DeleteForm.Password.focus();
		return false;//��ũ��Ʈ ����
	}
	window.document.DeleteForm.submit();//����
}
</script>
<%=Request("Num")%>�� ���� �����Ϸ��� ��й�ȣ�� �ʿ��մϴ�.<br>
<form action="BoardDeleteProcess.asp" method="post"
name="DeleteForm" onsubmit="return DeleteCheck();">
<input type="hidden" name="Num" value="<%=Request("Num")%>">
��й�ȣ : <input type="password" name="Password">
<br>
<input type="submit" value="����">
</form>