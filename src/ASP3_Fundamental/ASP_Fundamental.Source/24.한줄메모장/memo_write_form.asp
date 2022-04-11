<%
'--------------------------------------------------------------------
' Title : MyMemo Education Version
' Program Name : memo_write_form.asp
' Program Description : 메모글 저장하는 입력폼
' Include Files : None
' Last Update : 2001.08.25. 23:20
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------------------------
%>
<HTML>
	<HEAD>
		<TITLE>MY메모장 입력 페이지</TITLE>
		<SCRIPT language="javascript">
function write_form_check()
{
   //이름값 입력 체크
   if(window.document.write_form.name.value == "")
   {
      window.alert("이름을 넣어주세요.");   //메시지 박스 출력.
	  window.document.write_form.name.focus();   //해당 입력양식에 포커스.
	  window.document.write_form.name.select();   //해당 입력양식값을 선택.
	  return false;   //현재 함수를 호출한 곳으로 false값 반환.
   }
   //이름 길이값 체크
   if(window.document.write_form.name.value.length < 2 || window.document.write_form.name.value.length > 5)
   {
      window.alert("이름을 2자리이상 5자리 이하로 넣어주세요.");   //메시지 박스 출력.
	  window.document.write_form.name.focus();   //해당 입력양식에 포커스.
	  window.document.write_form.name.select();   //해당 입력양식값을 선택.
	  return false;   //현재 함수를 호출한 곳으로 false값 반환.
   }
   //이메일값 입력 체크
   if(window.document.write_form.email.value == "")
   {
      window.alert("이메일을 넣어주세요.");   //메시지 박스 출력.
	  window.document.write_form.email.focus();   //해당 입력양식에 포커스.
	  window.document.write_form.email.select();   //해당 입력양식값을 선택.
	  return false;   //현재 함수를 호출한 곳으로 false값 반환.
   }
   //이메일 길이값 입력 체크
   if(window.document.write_form.email.value.length < 10 || window.document.write_form.email.value.length > 30)
   {
      window.alert("이메일을 정확하게 넣어주세요.");   //메시지 박스 출력.
	  window.document.write_form.email.focus();   //해당 입력양식에 포커스.
	  window.document.write_form.email.select();   //해당 입력양식값을 선택.
	  return false;   //현재 함수를 호출한 곳으로 false값 반환.
   }
   //메모값 입력 체크
   if(window.document.write_form.title.value == "")
   {
      window.alert("메모를 남겨주세요.");   //메시지 박스 출력.
	  window.document.write_form.title.focus();   //해당 입력양식에 포커스.
	  window.document.write_form.title.select();   //해당 입력양식값을 선택.
	  return false;   //현재 함수를 호출한 곳으로 false값 반환.
   }
   //메모 길이값 입력 체크
   if(window.document.write_form.title.value.length < 1 || window.document.write_form.title.value.length > 50)
   {
      window.alert("메모의 길이를 50자이하로 넣어주세요.");   //메시지 박스 출력.
	  window.document.write_form.title.focus();   //해당 입력양식에 포커스.
	  window.document.write_form.title.select();   //해당 입력양식값을 선택.
	  return false;   //현재 함수를 호출한 곳으로 false값 반환.
   }
   window.document.write_form.submit();   //체크사항 만족시 전송.
}
		</SCRIPT>
	</HEAD>
	<BODY bgcolor="#FFFFFF">
		<DIV align="center">
			<TABLE border="1" width="100%">
				<FORM name="write_form" method="post" action="./memo_write_process.asp">
					<TR>
						<TD colspan="2" align="center" align="center">
							▣메모 남기기▣
						</TD>
					</TR>
					<TR>
						<TD align="center">
							이&nbsp;&nbsp;&nbsp;름
						</TD>
						<TD align="center">
							<INPUT type="text" name="name" size="30">
						</TD>
					</TR>
					<TR>
						<TD align="center">
							이메일
						</TD>
						<TD align="center">
							<INPUT type="text" name="email" size="30">
						</TD>
					</TR>
					<TR>
						<TD align="center">
							메&nbsp;&nbsp;&nbsp;모
						</TD>
						<TD align="center">
							<INPUT type="text" name="title" size="30">
						</TD>
					</TR>
					<TR>
						<TD colspan="2" align="center">
							<INPUT type="button" value="글쓰기" onclick="write_form_check();">&nbsp;&nbsp; 
							<INPUT type="reset" value="취&nbsp;&nbsp;소">
						</TD>
					</TR>
				</FORM>
			</TABLE>
		</DIV>
	</BODY>
</HTML>
