<%
	'Option Explicit

	Dim objFso   '드라이브의 정보를 얻어올 객체의 인스턴스용 변수.
	Dim objFolder   '특정 폴더의 모든 정보를 얻어올 핸들링 객체변수.

	'FileSystemObject객체의 인스턴스 생성
	Set objFso = Server.CreateObject("Scripting.FileSystemObject")

	'C:\InetPub의 핸들링 얻어오기
	'Set objFolder = objFso.GetFolder("C:\InetPub")
	Set objFolder = objFso.GetFolder("C:\Temp")
%>
[1] C:\Temp의 용량 : <%=objFolder.Size%><br>
<%
Set objFolder = Nothing
Set objFso = Nothing
%>