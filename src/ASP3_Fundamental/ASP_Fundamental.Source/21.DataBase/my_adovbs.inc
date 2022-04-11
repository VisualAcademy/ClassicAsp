<%
'--------------------------------------------------------------------
' Title : Microsoft ADO에 대한 설명
' Program Name : my_adovbs.inc
' Program Description : ADO 상수 사용법
' Include Files : ADO constants include file for VBScript
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------------------------


'--------------------------------------------------------------------
' [1] ADO 상수
'	ADO 객체의 프로퍼티와 메소드에서 사용되는 미리 정의된 많은 상수.

' [2] ADO 상수를 정의해 놓은 파일의 위치
'	C:\Program Files\Common Files\System\ado

' [3] ADO 상수를 포함하는 코드(ASP파일에 사용)
'	<!-- #include file="./adovbs.inc" -->

'[4] 타입라이브러리에 대한 참조를 사용한 ADO 상수를 포함하는 코드
'(아래 코드는 주로 global.asa파일에 포함시키면 애플리케이션의 모든 페이지에서 ADO 상수 사용 가능)
'	<!-- METADATA TYPE="typelib" FILE="C:\program files\common files\system\ado\msado25.dll"-->

'[5] OLE DB 연결 문자열 생성
'	해당 폴더에 oledb.txt파일을 만들어 놓고, 확장자를 udl로 변경한다.
'   oledb.udl 파일의 등록정보를 열고 지시에 따라 연결문자열을 생성한다.

'[6] 데이터 다루기(RecordSet 객체 사용)
'	레코드 셋 객체의 인스턴스 생성
'	Set objRS = Server.CreateObject("ADODB.RecordSet")

'	'레코드 셋 객체의 Open() 메소드를 사용해서 레코드 셋 가져오기
'	objRS.Open [Source], [ActiveConnection], [CursorType], [LockType], [Options]
		
'		[Source] : SQL 문자열
'		[ActiveConnection] :Connection객체의 인스턴스 또는 OLE DB 연결 문자열
'		[CursorType] : 사용할 커서 타입. 기본값은 adFowardOnly
'		[LockType] : 사용할 락 타입. 기본값은 adLockReadOnly
'		[Options] : 데이터 제공자에게 Source 인수가 무엇인지 알려준다.
'--------------------------------------------------------------------


'--------------------------------------------------------------------
'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

'---- LockTypeEnum Values ----
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

'---- CommandTypeEnum Values ----
Const adCmdUnknown = &H0008
Const adCmdText = &H0001
Const adCmdTable = &H0002
Const adCmdStoredProc = &H0004
Const adCmdFile = &H0100
Const adCmdTableDirect = &H0200
'--------------------------------------------------------------------
%>