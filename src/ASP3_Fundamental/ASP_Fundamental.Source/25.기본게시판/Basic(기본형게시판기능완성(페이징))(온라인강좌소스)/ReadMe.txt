※ 기본형 게시판 설치 설명서

1. 설치/테스트 환경
	Windows Server 2003 + SQL Server 2000 + ASP3.0
2. 설치 순서
	- Basic.zip(현재 게시판 전체 소스)를 아무데나 압축을 푼다.
		예) c:\inetpub\wwwroot\Basic\ ...
	- Basic 폴더를 웹공유(가상 디렉터리) 설정한다.
		예) http://localhost/Basic/
	- SQL Server EM를 열고 Basic 또는 다른 이름으로 데이터베이스를 만든다.
	- 해당 데이터베이스를 선택한 후 QA로 접속 후 Basic 폴더에 있는 Basic.sql 파일을 열고 Basic 테이블을 설치한다.
	- SQL Server EM에서 보안 > 로그인 사용자에서 아이디 : Basic, 비밀번호 : Basic으로 로그인 사용자를 하나 만든다.
	- Basic 폴더에서 빈 텍스트 파일 하나를 만들고 파일명을 Basic.udl로 변경한 후, 해당 파일을 더블클릭하여, 앞서 만든 데이터베이스에 연결할 수 있는 데이터베이스 연결문자열을 설정한다.
	- udl파일을 메모장으로 연 후 데이터베이스 연결문자열을 복사한 후 각각의 ASP 소스파일을 열어 objCon.Open() 안의 소스를 변경한다.
	- 웹브라우저를 열고 http://localhost/Basic/BoardList.asp를 실행한 후 기본형 게시판을 테스트한다.(Basic 폴더에 default.asp 추가)
	- 기타 게시판 디자인을 변경하여 맘대로 쓴다.
3. 기타
	- 100% 공짜입니다. 공부하시는데 참고가 되었으면 합니다.
