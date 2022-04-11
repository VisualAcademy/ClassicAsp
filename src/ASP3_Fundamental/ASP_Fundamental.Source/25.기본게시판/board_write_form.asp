<%
'--------------------------------------------------
' Title : MyBoard Education Version
' Program Name : board_write_form.asp
' Program Description : 폼태그를 이용한 게시판에 글쓰기
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>
<html>
<head>
	<title>MyBoard 글쓰기</title>

<script language="Javascript">
	<!--

		function check()   //MyMemo와 다른 방식으로 접근(onSubmit 이벤트핸들러 사용)
		{
			var valTitle = document.WriteForm.Title.value;
			var valContent = document.WriteForm.Content.value;
			var valName = document.WriteForm.Name.value;
			var valPassword = document.WriteForm.Password.value;
			
			if(valName == "") {
				alert("'작성자'를 입력해 주십시오.\n");
				document.WriteForm.Name.focus();
				document.WriteForm.Name.select();
				return false;
			}
				
			if(valTitle == "" || valTitle == " ") {
				alert("'제목'을 입력해 주십시오.\n");
				document.WriteForm.Title.focus();
				document.WriteForm.Title.select();
				return false;
			}

			if(valContent == "") {
				alert("'내용'을 입력해 주십시오.\n");
				document.WriteForm.Content.focus();
				document.WriteForm.Content.select();
				return false;
			}
			
			if(valPassword == "") {
				alert("'암호'를 입력해 주십시오.\n");
				document.WriteForm.Password.focus();
				document.WriteForm.Password.select();
				return false;
			}

			return true;

		} // end function

	//-->
</script>

</head>
<body>

<div align="center">
<table border="1">
<form action="./board_write_process.asp" method="Post" name="WriteForm" OnSubmit="return check();">
	<tr>
		<td align="center" colspan="3">▣게시판 글쓰기▣</td>
	</tr>
	<tr>
		<td align="center">
			작성자 :
		</td>
		<td colspan="2">
			<input type="text" name="Name" size="15" maxlength="15" value="" >
		</td>
	</tr>
	<tr>
		<td align="center">
			E-mail :
		</td>
		<td colspan="2">
			<input type="text" name="Email" size="30" maxlength="50" value="">
		</td>
	</tr>
	<tr>
		<td align="center">
			제&nbsp;&nbsp;&nbsp;목:
		</td>
		<td colspan="2">
			<input type="text" name="Title" size="60" maxlength="70" value="">
		</td>
	</tr>
	<tr>
		<td align="center">
			내&nbsp;&nbsp;&nbsp;용:
		</td>
		<td colspan="2">
			<textarea name="Content" cols="60" rows="10"></textarea>
		</td>
	</tr>
	<tr>
		<td colspan="3" align="center">
		<tt>
			■▣□●◎○◆◈◇♥♡♠♤♣♧※†‡☞☜★☆◑◐◀◁▶▷ <br>
			☏☎♨§㈜㉿㏘㏂™㏇♬♪♩↕♀♂】【∴∑▼▽▲△『』「」 
		</tt>
		</td>
	</tr>
	<tr>
		<td align="center">
			Encoding:
		</td>
		<td colspan="2">
			<select name="Encoding">
				<option selected>Text
				<option>HTML
			</select>
		</td>
	</tr>
	<tr>
		<td colspan="3">
			&nbsp;Encoding Type :<br>
			&nbsp;&nbsp;&nbsp;-	'Text' 는 HTML 코드를 실행하지 않고 내용 그대로를 보여줍니다.<br>
			&nbsp;&nbsp;&nbsp;-	'HTML' 는 HTML 코드를 실행한 결과를 보여줍니다.
		</td>
	</tr>
	<tr>
		<td valign="middle" align="center">
			암&nbsp;&nbsp;&nbsp;호:
		</td>
		<td>
			<input type="password" name="Password" size="10" maxlength="10" >
		</td>
		<td align="center">
			<input type="submit" value="작성 완료">
		</td>
	</tr>
	<tr>
		<td align="center" colspan="3">
			<a href="./board_list.asp">[게시판 리스트 보기]</a>
		</td>
	</tr>	
</form>
</table>
</div>

</body>
</html>