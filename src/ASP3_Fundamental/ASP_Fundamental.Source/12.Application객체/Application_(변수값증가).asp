<%
  Application.Lock 
  application("count") = application("count") + 1
  Application.UnLock 
%>
<HTML>
<HEAD>
</HEAD>
<BODY>
<br><center><font face="����" size="2">
<h2>���ø����̼� ����</h2>
<P>&nbsp;</P>
applicaton("count')�� ���� : <%=application("count")%>
</font></center>
</BODY>
</HTML>
