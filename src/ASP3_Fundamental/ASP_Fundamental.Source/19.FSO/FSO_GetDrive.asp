<%
'[0] 변수선언
Option Explicit
Dim objFSO   '드라이브의 정보를 얻어올 객체의 인스턴스용 변수.
Dim objCdrive   'C드라이브의 모든 정보를 얻어올 핸들링 객체변수.

'[1] 객체 생성
'FileSystemObject객체의 인스턴스 생성 : FSO객체를 objFSO란 이름으로 사용하겠다를 지정.
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

'[2] 핸들링 얻어오기
'C드라이브의 핸들링 얻어오기 : C드라이브의 정보를 받아서 
'objCDrive란 이름으로 사용하겠다를 지정.
Set objCDrive = objFSO.GetDrive("C:\")

'[3] 프로그램 처리
%>
[1] C드라이브의 전체 용량 : <%=objCdrive.TotalSize%><br>
[2] C드라이브의 남은 용량 : <%=objCdrive.FreeSpace%><br>
[3] C드라이브의 파일시스템 : <%=objCdrive.FileSystem%><br>
<%
'[4] 객체 해제
Set objCDrive = Nothing
Set objFSO = Nothing
%>