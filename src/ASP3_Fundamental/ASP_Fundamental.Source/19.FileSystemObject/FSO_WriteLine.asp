<%
	'[1] ��������
	Option Explicit
	Dim objFSO, objFile

	'[2] objFSO�� �̸����� ���� ��ü�� ���.
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	'[3] �ؽ�Ʈ ���� �о���� : OpenTextFile�޼ҵ�("��ι������̸�",�б�/���������,���������,�����ڵ�)
	'���ϰ�ü�� �ڵ鸵 ������
	Set objFile = objFSO.OpenTextFile("c:\test1.txt",8,true) 
	'Set objFile = objFSO.OpenTextFile("c:\test1.txt",8,true,false) 
%>
<% 
	'[4] �ؽ�Ʈ ���Ͽ� ���پ� ���� : WriteLine�޼ҵ� : �ڵ����� ����ĭ���� ��Ŀ�� �̵�.
	objFile.WriteLine("�޷�~~~")
	objFile.WriteLine("�� ���� �ι�° ���ο� �������ϴ�.")
	objFile.WriteLine("�� ���� ����° ���ο� �������ϴ�.")

	'[5] ����� ��ü ����
	Set objFile = Nothing   '��ü ����
	Set objFSO = Nothing	'��ü ����
%>
<center>�۾��� �Ϸ�</center>

