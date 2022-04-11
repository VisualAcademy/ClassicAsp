<%
  Application.Lock 
  application("count") = application("count") + 1
  Application.UnLock 
%>
<HTML>
<HEAD>
</HEAD>
<BODY>
<br><center><font face="돋움" size="2">
<h2>어플리케이션 증가</h2>
<P>&nbsp;</P>
applicaton("count')의 값은 : <%=application("count")%>
</font></center>
</BODY>
</HTML>
