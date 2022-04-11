<%
'--------------------------------------------------------------------
' Title : Microsoft ADO�� ���� ����
' Program Name : my_adovbs.inc
' Program Description : ADO ��� ����
' Include Files : ADO constants include file for VBScript
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------------------------


'--------------------------------------------------------------------
' [1] ADO ���
'	ADO ��ü�� ������Ƽ�� �޼ҵ忡�� ���Ǵ� �̸� ���ǵ� ���� ���.

' [2] ADO ����� ������ ���� ������ ��ġ
'	C:\Program Files\Common Files\System\ado

' [3] ADO ����� �����ϴ� �ڵ�(ASP���Ͽ� ���)
'	<!-- #include file="./adovbs.inc" -->

'[4] Ÿ�Զ��̺귯���� ���� ������ ����� ADO ����� �����ϴ� �ڵ�
'(�Ʒ� �ڵ�� �ַ� global.asa���Ͽ� ���Խ�Ű�� ���ø����̼��� ��� ���������� ADO ��� ��� ����)
'	<!-- METADATA TYPE="typelib" FILE="C:\program files\common files\system\ado\msado25.dll"-->

'[5] OLE DB ���� ���ڿ� ����
'	�ش� ������ oledb.txt������ ����� ����, Ȯ���ڸ� udl�� �����Ѵ�.
'   oledb.udl ������ ��������� ���� ���ÿ� ���� ���Ṯ�ڿ��� �����Ѵ�.

'[6] ������ �ٷ��(RecordSet ��ü ���)
'	���ڵ� �� ��ü�� �ν��Ͻ� ����
'	Set objRS = Server.CreateObject("ADODB.RecordSet")

'	'���ڵ� �� ��ü�� Open() �޼ҵ带 ����ؼ� ���ڵ� �� ��������
'	objRS.Open [Source], [ActiveConnection], [CursorType], [LockType], [Options]
		
'		[Source] : SQL ���ڿ�
'		[ActiveConnection] :Connection��ü�� �ν��Ͻ� �Ǵ� OLE DB ���� ���ڿ�
'		[CursorType] : ����� Ŀ�� Ÿ��. �⺻���� adFowardOnly
'		[LockType] : ����� �� Ÿ��. �⺻���� adLockReadOnly
'		[Options] : ������ �����ڿ��� Source �μ��� �������� �˷��ش�.
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