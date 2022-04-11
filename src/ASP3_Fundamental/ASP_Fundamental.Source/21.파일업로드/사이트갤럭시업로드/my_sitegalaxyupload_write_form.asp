<%
' -----------------------------------------------------------
' Title : Using SiteGalaxyUpload Component
' Program Name : my_sitegalaxyupload_write_form.asp
' Description : 파일올리기 기본 폼
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: 
' Support: 
' -----------------------------------------------------------
%>
<HTML>
<HEAD>
<TITLE>사이트갤럭시업로드 컴포넌트 사용법</TITLE>
<script language="javascript">
	<!--
		function check()
		{
			if(window.document.uploadform.FileName.value == "") {
				window.alert("'파일'을 첨부해 주십시오.\n");
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
	<h3>사이트 갤럭시 업로드 컴포넌트를 사용해서 파일 올리기</h3>
	<form name="uploadform" method="post" action="./my_sitegalaxyupload_write_process.asp"  enctype="multipart/form-data" onsubmit="return check();">
	<input type="file" name="FileName">
	<input type="submit" value="업로드">
	</form>
	</div>
</BODY>
</HTML>