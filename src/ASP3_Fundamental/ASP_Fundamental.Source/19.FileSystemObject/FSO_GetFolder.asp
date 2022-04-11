<%
	'Option Explicit

	Dim objFSO   '드라이브의 정보를 얻어올 객체의 인스턴스용 변수.
	Dim objFolder   '특정 폴더의 모든 정보를 얻어올 핸들링 객체변수.

	'FileSystemObject객체의 인스턴스 생성
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	'C:\InetPub의 핸들링 얻어오기
	Set objFolder = objFSO.GetFolder("C:\InetPub")
%>

[1] C:\InetPub의 용량 : <%=objFolder.Size%><br>


