
<h4> Server.MapPath 예제 : 물리적 디렉토리 경로 구하기 </h4>

<!-- MapPath()는 현재 스크립트의 경로를 반환시켜준다. -->
<hr>
MapPath(".") : <%=Server.MapPath(".")%><br>
현재 ASP가 있는 디렉토리 (.) : <%= Server.MapPath(".") %><br>
<hr>
MapPath("\") : <%=Server.MapPath("\")%><br>
홈 디렉토리 경로 (\) : <%= Server.MapPath("\") %><br>
<hr>
MapPath("..") : <%=Server.MapPath("..")%><br>
ASP가 있는 상위 디렉토리 이름 (..) : <%= Server.MapPath("..") %><br>
<hr>
MapPath("/") : <%=Server.MapPath("/")%><br>
홈 디렉토리 경로 (/) : <%= Server.MapPath("/") %><br>
<hr>
MapPath("/MyTest") : <%=Server.MapPath("/MyTest")%><br>
가상 디렉토리 디렉토리 경로(/asp3) : <%= Server.MapPath("/asp3") %><br>