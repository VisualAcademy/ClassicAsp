<html>
<head>
<script language="javascript">

	function check()
	{
		if(document.myform.userid.value == "")
		{
			alert("아뒤 넣어라!");
			document.myform.userid.focus();
			document.myform.userid.select();
			return true;
		}
		if(document.myform.userid.value.length < 3 || document.myform.userid.value.length >15)
		{
			alert("아뒤가 허접하당.");
			document.myform.userid.focus();
			document.myform.userid.select();
			return true;
		}

		document.myform.submit();
	}


</script>
</head>
<body bgcolor="#fcffef">

<p> 폼 이용한 입력 페이지 </p>
<hr>

<form name="myform" action="form_process.asp" method="post"> 

   사용자 ID : <input type="text" name="userid" size=20><br>
   비밀번호 : <input type="password" name="password" size=20><p>

   <input type="button" value="보내기" onclick="check();">
   <input type="reset" value="초기화">
 
</form>
   
</body>
</html>
