<%
'[1] AdRotator 컴포넌트의 인스턴스 생성
Set objAd = Server.CreateObject("MSWC.AdRotator")

'[2] AdRotator.txt 파일 호출(광고 보이기)
Response.Write(objAd.GetAdvertisement("/Sample/AdRotator/AdRotator.txt"))

'[3] 객체 해제
Set objAd = Nothing
%>