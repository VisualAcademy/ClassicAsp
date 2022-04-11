<%
	'[1] 변수선언
	Option Explicit
	Dim objFSO, objFile

	'[2] objFSO란 이름으로 파일 객체를 사용.
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	'[3] 텍스트 파일 읽어오기 : OpenTextFile메소드("경로및파일이름",읽기/쓰기모드결정,덮어쓰기유무,유니코드)
	'파일객체의 핸들링 얻어오기
	Set objFile = objFSO.OpenTextFile("c:\test1.txt",8,true) 
	'Set objFile = objFSO.OpenTextFile("c:\test1.txt",8,true,false) 
%>
<% 
	'[4] 텍스트 파일에 한줄씩 쓰기 : WriteLine메소드 : 자동으로 다음칸으로 포커스 이동.
	objFile.WriteLine("메롱~~~")
	objFile.WriteLine("이 글은 두번째 라인에 쓰여집니다.")
	objFile.WriteLine("이 글은 세번째 라인에 쓰여집니다.")

	'[5] 선언된 객체 해제
	Set objFile = Nothing   '객체 해제
	Set objFSO = Nothing	'객체 해제
%>
<center>글쓰기 완료</center>

