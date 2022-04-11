<script language="javascript">
function DeleteCheck(){
	var valPassword = 
		window.document.DeleteForm.Password.value;
	if(valPassword == ""){
		window.alert("비밀번호를 넣어주세요.");
		window.document.DeleteForm.Password.focus();
		return false;//스크립트 종료
	}
	window.document.DeleteForm.submit();//전송
}
</script>
<%=Request("Num")%>번 글을 삭제하려면 비밀번호가 필요합니다.<br>
<form action="BoardDeleteProcess.asp" method="post"
name="DeleteForm" onsubmit="return DeleteCheck();">
<input type="hidden" name="Num" value="<%=Request("Num")%>">
비밀번호 : <input type="password" name="Password">
<br>
<input type="submit" value="삭제">
</form>