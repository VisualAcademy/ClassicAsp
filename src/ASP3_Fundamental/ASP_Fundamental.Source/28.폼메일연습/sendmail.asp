<%
'변수의 명시적 선언 정의
Option Explicit

'변수 선언
Dim objMail, FromName, FromEmail, ToEmail, Subject, MailBody, flagHTML

'넘겨져온 각 Form Value를 가져와 변수에 저장
FromName = Request("fromname")
FromEmail = Request("fromemail")
ToEmail = Request("toemail")
Subject = Request("subject")
MailBody = Request("content")
flagHTML = Request("flagHTML")

'CDO객체의 인스턴스 생성
Set objMail = Server.CreateObject("CDONTS.NewMail")

'CDONTS NewMail 객체의 각 값을 지정한다.
'Mail에서 보내는 사람의 이름(여기서는 <>안에 이메일주소도 &로 추가했다.)
'objMail.From = FromName & "<" & FromEmail & ">"   '아래코드와 같은 의미
objMail.Value("From") = FromName & "<" & FromEmail & ">"

'Mail에서 보내는 사람의 이메일 주소
objMail.To = ToEmail

'Mail의 제목
objMail.Subject = Subject

'HTML로 전송할지 Text로 전송할지 결정
'HTML 형식의 Mail 일경우(지정하지 않으면 일반 TEXT 형식)
If flagHTML = "1" Then
  objMail.BodyFormat = 0
  objMail.MailFormat = 0
End If

'Mail의 내용
objMail.Body = MailBody

'최종적으로 각 객체의 값들을 전송한다.
objMail.Send()

'객체의 인스턴스 해제
Set objMail = Nothing
%>

<html>
<head>
<title>SEND MAIL</title>
<style type="text/css">
<!--
    A:link    {font: 10pt GulimChe,굴림; color: #0000ff; text-decoration: none}
    A:active  {font: 10pt GulimChe,굴림; color: #ff1493; text-decoration: none}
    A:visited {font: 10pt GulimChe,굴림;  text-decoration: none}
    A:hover   {text-decoration:  color: #ff1493;}
      td        {font: 10pt GulimChe,굴림; color: black;}
     .bold     {font: bold 13pt 굴림; color: white;}
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
        <td height="17" align="center" bgcolor="#E7F0BB"><p align="center">■ SEND MAIL ■</td>
    </tr>
    <tr>
        <td width="964">
			<p align="center"><blink><b><%=ToEmail%></b>님께 MAIL 전송완료.</blink><br> </p>
		</td>
    </tr>
    <tr>
        <td width="964"><p align="center">
			<a href="./MainPage.asp"><blink>2초 후에 메인 페이지로 이동합니다.</blink></a>
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