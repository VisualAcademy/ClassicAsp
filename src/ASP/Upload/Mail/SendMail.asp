<%
Set objNewMessage = Server.CreateObject("CDO.Message")

objNewMessage.From = "redplus@redplus.net"     '//보내는 사람
objNewMessage.To = "redplus@redplus.net"     '//받는 사람
objNewMessage.Subject = "2003에서 보내기 테스트"     '//제목
objNewMessage.HtmlBody = "<b>HTML로 변환전송</b>"     '//내용

objNewMessage.Send()                                        '//전송

Set objNewMessage = Nothing
%>
<center>메일 전송 완료</center> 