<%
'--------------------------------------------------
' Title : Myado Education Version
' Program Name : ado_modify.asp
' Program Description : 폼태그를 이용한 입력 내용 수정
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>

<!--METADATA TYPE= "typelib" NAME= "ADODB Type Library" FILE="C:\Program Files\Common Files\SYSTEM\ADO\msado15.dll" -->

<%
'변수의 명시적 선언
Option Explicit
Dim objDB, objRS, strSQL, Num
Dim strDB

'넘겨져온 레코드 넘버(Num)값을 변수에 저장
Num = Request.QueryString("Num")   '주의 : QueryString 컬렉션으로 받아야 함.

'--------------------------------------------------------------------
'[1] 레코드 셋 객체의 인스턴스 생성
Set objRS = Server.CreateObject("Adodb.RecordSet")


'[2] 데이터베이스 연결 문자열 생성
strDB = "Provider=SQLOLEDB.1;Password=my_database_pwd;Persist Security Info=True;User ID=my_database_id;Initial Catalog=my_database;Data Source=ZZAZIPKIDOTCOM"

'[3] SQL문 생성
strSQL = "Select * From my_ado Where Num = " & Num
'Response.Write strSQL

'[4] 레코드 셋 열기
objRs.Open strSQL, strDB, adOpenStatic, adLockPessimistic, adCmdText
'--------------------------------------------------------------------
%>
<html>
<head>
	<title>Myado 수정하기</title>

<script language="Javascript">
	<!--

		function check()   //MyMemo 및 MyGuest와 다른 방식으로 접근(onSubmit 이벤트핸들러 사용)
		{
			// 변수에 해당 객체 구문 대입
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

			return true;   //조건을 만족하면 submit

		} // end function

	//-->
</script>

</head>
<body>

<div align="center">
<table border="1">
<form action="./ado_modify_process.asp" method="Post" name="WriteForm" OnSubmit="return check();">
	<input type="hidden" name="num" value="<%=Num%>">
	<tr>
		<td align="center" colspan="3">▣게시판 글 수정하기▣</td>
	</tr>
	<tr>
		<td align="center">
			작성자 :
		</td>
		<td colspan="2">
			<input type="text" name="Name" size="15" maxlength="15" value="<%=objRS.Fields("Name")%>">
		</td>
	</tr>
	<tr>
		<td align="center">
			E-mail :
		</td>
		<td colspan="2">
			<input type="text" name="Email" size="30" maxlength="50" value="<%=objRS.Fields("Email")%>">
		</td>
	</tr>
	<tr>
		<td align="center">
			제&nbsp;&nbsp;&nbsp;목:
		</td>
		<td colspan="2">
			<input type="text" name="Title" size="60" maxlength="70" value="<%=objRS.Fields("Title")%>">
		</td>
	</tr>
	<tr>
		<td align="center">
			내&nbsp;&nbsp;&nbsp;용:
		</td>
		<td colspan="2">
			<textarea name="Content" cols="60" rows="10"><%=objRS.Fields("Content")%></textarea>
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
				<option <% If objRS.Fields("Encoding") = "Text" Then %>selected<% End If %>>Text
				<option <% If objRS.Fields("Encoding") = "HTML" Then %>selected<% End If %>>HTML
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
			<input type="submit" value="수정 완료">
		</td>
	</tr>
	<tr>
		<td align="center" colspan="3">
			<a href="./ado_list.asp">[게시판 리스트 보기]</a>
		</td>
	</tr>	
</form>
</table>
</div>

</body>
</html>

<%
objRS.Close

Set objRS = Nothing
%>
