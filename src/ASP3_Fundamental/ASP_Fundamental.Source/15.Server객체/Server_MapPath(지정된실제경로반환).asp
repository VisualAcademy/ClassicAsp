
<h4> Server.MapPath ���� : ������ ���丮 ��� ���ϱ� </h4>

<!-- MapPath()�� ���� ��ũ��Ʈ�� ��θ� ��ȯ�����ش�. -->
<hr>
MapPath(".") : <%=Server.MapPath(".")%><br>
���� ASP�� �ִ� ���丮 (.) : <%= Server.MapPath(".") %><br>
<hr>
MapPath("\") : <%=Server.MapPath("\")%><br>
Ȩ ���丮 ��� (\) : <%= Server.MapPath("\") %><br>
<hr>
MapPath("..") : <%=Server.MapPath("..")%><br>
ASP�� �ִ� ���� ���丮 �̸� (..) : <%= Server.MapPath("..") %><br>
<hr>
MapPath("/") : <%=Server.MapPath("/")%><br>
Ȩ ���丮 ��� (/) : <%= Server.MapPath("/") %><br>
<hr>
MapPath("/MyTest") : <%=Server.MapPath("/MyTest")%><br>
���� ���丮 ���丮 ���(/asp3) : <%= Server.MapPath("/asp3") %><br>