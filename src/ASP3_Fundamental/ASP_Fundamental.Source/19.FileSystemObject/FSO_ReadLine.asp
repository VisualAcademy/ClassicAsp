<%
	'[1] 변수선언
	Option Explicit
	Dim objFSO, objFile

	'[2] FSO객체의 인스턴스 생성
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	'[3] 파일 객체 생성 : 핸들링 읽어오기 : 읽기전용모드 1로
	Set objFile = objFSO.OpenTextFile("C:\test1.txt",1)

	'[4] 파일 읽어오기 : ReadLine메소드는 한줄씩 읽어온다.
	Response.Write objFile.ReadLine & "<br>"
	Response.Write objFile.ReadLine & "<br>"
	Response.Write objFile.ReadLine & "<br>"

	'[5] 객체 해제
	Set objFile = Nothing
	Set objFSO = Nothing
%>
