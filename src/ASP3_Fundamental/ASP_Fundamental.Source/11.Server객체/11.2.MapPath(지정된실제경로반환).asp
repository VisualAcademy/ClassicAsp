<!-- MapPath()는 현재 스크립트의 경로를 반환시켜준다. -->

현재 ASP가 있는 디렉토리 (.) : <%= Server.MapPath(".") %><br>
<hr>
홈 디렉토리 경로 (\) : <%= Server.MapPath("\") %><br>
<hr>
홈 디렉토리 경로 (/) : <%= Server.MapPath("/") %><br>
<hr>
한단계 상위 디렉토리 이름 (..) : <%= Server.MapPath("..") %><br>
<hr>
가상 디렉토리 디렉토리 경로(/asp) : <%= Server.MapPath("/asp") %><br>