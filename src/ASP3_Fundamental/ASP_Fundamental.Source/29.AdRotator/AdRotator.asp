<%
'[1] AdRotator ������Ʈ�� �ν��Ͻ� ����
Set objAd = Server.CreateObject("MSWC.AdRotator")

'[2] AdRotator.txt ���� ȣ��(���� ���̱�)
Response.Write(objAd.GetAdvertisement("/Sample/AdRotator/AdRotator.txt"))

'[3] ��ü ����
Set objAd = Nothing
%>