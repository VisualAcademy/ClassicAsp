<%
'������ ����� ���� ����
Option Explicit

'���� ����
Dim objMail, FromName, FromEmail, ToEmail, Subject, MailBody, flagHTML

'�Ѱ����� �� Form Value�� ������ ������ ����
FromName = Request("fromname")
FromEmail = Request("fromemail")
ToEmail = Request("toemail")
Subject = Request("subject")
MailBody = Request("content")
flagHTML = Request("flagHTML")

'CDO��ü�� �ν��Ͻ� ����
Set objMail = Server.CreateObject("CDONTS.NewMail")

'CDONTS NewMail ��ü�� �� ���� �����Ѵ�.
'Mail���� ������ ����� �̸�(���⼭�� <>�ȿ� �̸����ּҵ� &�� �߰��ߴ�.)
'objMail.From = FromName & "<" & FromEmail & ">"   '�Ʒ��ڵ�� ���� �ǹ�
objMail.Value("From") = FromName & "<" & FromEmail & ">"

'Mail���� ������ ����� �̸��� �ּ�
objMail.To = ToEmail

'Mail�� ����
objMail.Subject = Subject

'HTML�� �������� Text�� �������� ����
'HTML ������ Mail �ϰ��(�������� ������ �Ϲ� TEXT ����)
If flagHTML = "1" Then
  objMail.BodyFormat = 0
  objMail.MailFormat = 0
End If

'Mail�� ����
objMail.Body = MailBody

'���������� �� ��ü�� ������ �����Ѵ�.
objMail.Send()

'��ü�� �ν��Ͻ� ����
Set objMail = Nothing
%>

<html>
<head>
<title>SEND MAIL</title>
<style type="text/css">
<!--
    A:link    {font: 10pt GulimChe,����; color: #0000ff; text-decoration: none}
    A:active  {font: 10pt GulimChe,����; color: #ff1493; text-decoration: none}
    A:visited {font: 10pt GulimChe,����;  text-decoration: none}
    A:hover   {text-decoration:  color: #ff1493;}
      td        {font: 10pt GulimChe,����; color: black;}
     .bold     {font: bold 13pt ����; color: white;}
-->
</style>
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red">
<br>
<br>
<br>
<br>

<div align="center"><table border cellpadding="5" cellspacing="0" width="600"
 bordercolordark="white" bordercolorlight="#A56C04">
    <tr>
        <td height="17" align="center" bgcolor="#E7F0BB"><p align="center">�� SEND MAIL ��</td>
    </tr>
    <tr>
        <td width="964">
			<p align="center"><blink><b><%=ToEmail%></b>�Բ� MAIL ���ۿϷ�.</blink><br> </p>
		</td>
    </tr>
    <tr>
        <td width="964"><p align="center">
			<a href="./MainPage.asp"><blink>2�� �Ŀ� ���� �������� �̵��մϴ�.</blink></a>
			</p>
        </td>
    </tr>
</table></div>
<div align="center"><table border cellpadding="3" cellspacing="0" width="600"
 bgcolor="#FFCC00" bordercolordark="white" bordercolorlight="#D3B670" style="border-width:1px; border-bottom-color:black; border-top-color:black; border-right-color:black; border-left-color:black; border-bottom-style:solid; border-left-style:solid; border-top-style:solid; border-right-style:solid;">
    <tr>
        <td height="15" align="center" bgcolor="#E4D6C0" class="small"><p><a
             href="http://www.redplus.net/" target="_blank" title="PYJ HOMEPAGE">PYJ-LOVE 
            ActiveMainPage COPYRIGHT 2000 BY RED+.NET ALL RIGHTS RESERVED.</a> 
        </td>
    </tr>
</table></div>

<script language="JavaScript">
window.setTimeout("location.href='./MainPage.asp'", 2000);
</script>

</body>

</html>