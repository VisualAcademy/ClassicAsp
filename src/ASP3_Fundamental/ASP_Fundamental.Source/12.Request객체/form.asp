<html>
<head>
<script language="javascript">

	function check()
	{
		if(document.myform.userid.value == "")
		{
			alert("�Ƶ� �־��!");
			document.myform.userid.focus();
			document.myform.userid.select();
			return true;
		}
		if(document.myform.userid.value.length < 3 || document.myform.userid.value.length >15)
		{
			alert("�Ƶڰ� �����ϴ�.");
			document.myform.userid.focus();
			document.myform.userid.select();
			return true;
		}

		document.myform.submit();
	}


</script>
</head>
<body bgcolor="#fcffef">

<p> �� �̿��� �Է� ������ </p>
<hr>

<form name="myform" action="form_process.asp" method="post"> 

   ����� ID : <input type="text" name="userid" size=20><br>
   ��й�ȣ : <input type="password" name="password" size=20><p>

   <input type="button" value="������" onclick="check();">
   <input type="reset" value="�ʱ�ȭ">
 
</form>
   
</body>
</html>
