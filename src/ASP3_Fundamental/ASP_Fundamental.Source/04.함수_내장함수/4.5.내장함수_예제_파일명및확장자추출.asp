<%@ Language=VBScript %>
<%
'이 예제는 자료실 구축시 파일의 이름과 확장자 등의 정보를 구할 때 사용되는 로직이다.
%>
<%
'변수 선언
Option Explicit

Dim strFilePath   '파일의 전체경로를 저장
Dim strFileName   '파일의 이름과 확장자를 저장
Dim strName       '파일의 이름만 저장
Dim strExt        '파일의 확장자만 저장

'파일 경로 및 파일명 지정
strFilePath = "C:\PYJ\PYJ.LOVE.exe" 

'파일명, 확장자 추출하기.
strFileName = Mid(strFilePath, InstrRev(strFilePath, "\") + 1 )   '파일명.확장자 추출

'strName = Mid(strFileName, 1, Instr(strFileName, ".")-1)   '파일명 추출
strName = Mid(strFileName, 1, InstrRev(strFileName, ".")-1)   '파일명 추출

strExt = Mid(strFileName, InstrRev(strFileName,".")+1) '확장자 추출

'추출된 정보 출력하기.
Response.Write("파일의 전체경로 : " & strFilePath & "<br>")

Response.Write("파일명.확장자 : " & strFileName & "<br>")

Response.Write("파일명 : " & strName & "<br>")

Response.Write("확장자 : " & strExt & "<br>")
%>