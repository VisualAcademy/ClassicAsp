<%
	'[1]변수 선언
	Option Explicit
	Dim objFso   'FSO객체의 인스턴스 선언
	Dim objFile   'FSO객체의 File객체용 핸들링 변수

	'[2]FSO 객체의 인스턴스(실체)를 생성한다.(파일 관련 객체의 설계도는 Scripting.FileSystemObject)
	Set objFso = Server.CreateObject("Scripting.FileSystemObject")

	'[3]파일생성 : CreateTextFile("파일이름및경로",덮어쓰기유무,유니코드저장유무)
	Set objFile = objFso.CreateTextFile("C:\Temp\test1.txt",true,false)
	Set objFile = objFso.CreateTextFile("C:\Temp\test2.txt",true,false)
	Set objFile = objFso.CreateTextFile("C:\Temp\test3.txt",true,false)

	'파일 존재유무 체크 : FileExists -> 존재하면 True, 그렇지않으면 False값 반환 메소드.
	If objFso.FileExists("C:\Temp\test1.txt") Then
		Response.Write("test1.txt 파일이 만들어졌습니다.<br>")
	End If
	If objFso.FileExists("C:\Temp\test2.txt") Then
		Response.Write("test2.txt 파일이 만들어졌습니다.<br>")
	End If
	If objFso.FileExists("C:\Temp\test3.txt") Then
		Response.Write("test3.txt 파일이 만들어졌습니다.<br>")
	End If

	'[4]FSO객체의 인스턴스 해제(ASP가 알아서 해제 시켜주지만, 기본적으로 작성해라...)
	Set objFile = Nothing
	Set objFso = Nothing
%>
