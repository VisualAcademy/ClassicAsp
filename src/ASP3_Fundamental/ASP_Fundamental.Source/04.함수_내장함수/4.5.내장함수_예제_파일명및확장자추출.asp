<%@ Language=VBScript %>
<%
'�� ������ �ڷ�� ����� ������ �̸��� Ȯ���� ���� ������ ���� �� ���Ǵ� �����̴�.
%>
<%
'���� ����
Option Explicit

Dim strFilePath   '������ ��ü��θ� ����
Dim strFileName   '������ �̸��� Ȯ���ڸ� ����
Dim strName       '������ �̸��� ����
Dim strExt        '������ Ȯ���ڸ� ����

'���� ��� �� ���ϸ� ����
strFilePath = "C:\PYJ\PYJ.LOVE.exe" 

'���ϸ�, Ȯ���� �����ϱ�.
strFileName = Mid(strFilePath, InstrRev(strFilePath, "\") + 1 )   '���ϸ�.Ȯ���� ����

'strName = Mid(strFileName, 1, Instr(strFileName, ".")-1)   '���ϸ� ����
strName = Mid(strFileName, 1, InstrRev(strFileName, ".")-1)   '���ϸ� ����

strExt = Mid(strFileName, InstrRev(strFileName,".")+1) 'Ȯ���� ����

'����� ���� ����ϱ�.
Response.Write("������ ��ü��� : " & strFilePath & "<br>")

Response.Write("���ϸ�.Ȯ���� : " & strFileName & "<br>")

Response.Write("���ϸ� : " & strName & "<br>")

Response.Write("Ȯ���� : " & strExt & "<br>")
%>