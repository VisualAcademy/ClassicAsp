<%
'=================================
' MySendMail Education Version
' Copyright (C) 2000 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'=================================
%>
<% 
SUB MainSendMail()
%>

<html>
<head>
<title>MySendMail</title>
<script language="javascript">
<!--
function mailcheck()
{
	//������ ��
	if (document.mailform.fromname.value =="") {
		alert("�̸��� �����ּ���");
		return false;
	}
	//������ �� ����
	if (document.mailform.fromemail.value =="") {
		alert("�̸����� �����ּ���");
		return false;
	}
	//�޴� �� 
	if (document.mailform.toemail.value =="") {
		alert("�̸����� �����ּ���");
		return false;
	}
document.mailform.submit();
}
//-->
</script>
</head>
<body>
	<table border cellspacing="0" width="100%" bgcolor="<%=main_bgcolor%>" bordercolordark="white" bordercolorlight="#A56C03">
			<tr>
				<td height="17" align="<%=TopicAlign(main_topic_align)%>" bgcolor="<%=main_list_color%>" >
				<form name="mailform" method="post" action="sendmail.asp">
				<font color="<%=main_lst_topic_color%>">
				�� SEND MAIL ��
				</font>
				</a>
			
				</td>
			</tr>
			<tr>
				<td align="right" valign="top" class="text1" onmouseover="this.style.backgroundColor='#FFFFCC'" onmouseout="this.style.backgroundColor=''">

				�غ������� �̸� : <input type="������ ��" name="fromname" size="17" style="border:1 solid" class=notice>

				</td>
			</tr>
			<tr>
				<td align="right" valign="top" class="text1" onmouseover="this.style.backgroundColor='#FFFFCC'" onmouseout="this.style.backgroundColor=''">

				�غ������� ���� : <input type="text" name="fromemail" size="17" style="border: 1 solid" class=notice>

				</td>
			</tr>
			<tr>
				<td align="right" valign="top" class="text1" onmouseover="this.style.backgroundColor='#FFFFCC'" onmouseout="this.style.backgroundColor=''">

				�ع޴��� ���� : <input type="text" name="toemail" size="17" style="border: 1 solid" class=notice>

				</td>
			</tr>
			<tr>
				<td align="right" valign="top" class="text1" onmouseover="this.style.backgroundColor='#FFFFCC'" onmouseout="this.style.backgroundColor=''">

				������ : <input type="text" name="subject" size="24" style="border: 1 solid">

				</td>
			</tr>
															
			<tr>
				<td align="center" valign="top" class="text1" onmouseover="this.style.backgroundColor='#FFFFCC'" onmouseout="this.style.backgroundColor=''">

				<textarea name="content" rows="3" width =100% cols="32" style="border: 1 solid"></textarea>

				</td>
			</tr>
			<tr>
				<td align = "center" valign="top" class="text1" onmouseover="this.style.backgroundColor='#FFFFCC'" onmouseout="this.style.backgroundColor=''">
				<p align="center">
				<input type="button" value="����������" onclick="mailcheck();" style="border:1 solid; height:18">

				&nbsp;&nbsp;

				<input type=reset value="���������" onclick=history.go(-1) style="border:1 solid; height:18;"><br>
				<input type=checkbox name="flagHTML" value="1">HTML ����
				</p>
				</td>
				</form>					
			</tr>
	</table>
</body>    
</html>

<%
End sub
%>
