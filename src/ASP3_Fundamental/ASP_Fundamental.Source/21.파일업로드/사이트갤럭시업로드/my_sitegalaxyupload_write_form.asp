<%
' -----------------------------------------------------------
' Title : Using SiteGalaxyUpload Component
' Program Name : my_sitegalaxyupload_write_form.asp
' Description : ���Ͽø��� �⺻ ��
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: 
' Support: 
' -----------------------------------------------------------
%>
<HTML>
<HEAD>
<TITLE>����Ʈ�����þ��ε� ������Ʈ ����</TITLE>
<script language="javascript">
	<!--
		function check()
		{
			if(window.document.uploadform.FileName.value == "") {
				window.alert("'����'�� ÷���� �ֽʽÿ�.\n");
				window.document.uploadform.FileName.focus();
				window.document.uploadform.FileName.select();
				return false;
			}
			return true;
		} // end function
	//-->
</script>
</HEAD>
<BODY BGCOLOR="#FFFFFF">
	<div align="center">	
	<h3>����Ʈ ������ ���ε� ������Ʈ�� ����ؼ� ���� �ø���</h3>
	<form name="uploadform" method="post" action="./my_sitegalaxyupload_write_process.asp"  enctype="multipart/form-data" onsubmit="return check();">
	<input type="file" name="FileName">
	<input type="submit" value="���ε�">
	</form>
	</div>
</BODY>
</HTML>