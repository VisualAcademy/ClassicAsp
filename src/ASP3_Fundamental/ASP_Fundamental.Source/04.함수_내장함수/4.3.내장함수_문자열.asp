<%
Option Explicit
Dim strTest

strTest = " Abc Def Ghi ihG feD cbA "
%>

strTest�� ����� ���ڿ� : <%=strTest%> <br>
strTest�� ����� ���ڿ��� ���� : <%=Len(strTest)%> <br>
strTest�� ����� ���ڿ��� ��� �빮�ڷ� ��ȯ : <%=UCase(strTest)%> <br>
strTest�� ����� ���ڿ��� ��� �ҹ��ڷ� ��ȯ : <%=LCase(strTest)%> <br>
strTest�� ����� ���ڿ� �� Def�� ������ġ : <%=InStr(strTest,"Def")%> <br>
strTest�� ����� ���ڿ� �� f�� ��ġ(�տ���) : <%=InStr(strTest,"f")%> <br>
strTest�� ����� ���ڿ� �� f�� ��ġ(�ڿ���) : <%=InStrRev(strTest,"f")%> <br>
strTest�� ����� ���ڿ� �� ���ʿ��� 4�� �˻�(���ʿ���) : <%=Left(strTest,4)%> <br>
strTest�� ����� ���ڿ� �� �����ʿ��� 4�� �˻�(���ʿ���) : <%=Right(strTest,4)%> <br>
strTest�� ����� ���ڿ� �� 8��°���� 10��°����(3����)�� ���ڿ� �˻� : <%=Mid(strTest,8,3)%> <br>
strTest�� ����� ���ڿ� �� 8��°�� ������ ���ڿ� �˻� : <%=Mid(strTest,8)%> <br>
strTest�� ����� ���ڿ� �� ���ʰ����� ���� : <%=LTrim(strTest)%> <br>
strTest�� ����� ���ڿ� �� ���ʰ����� ����(���ڿ� ����) : <%=Len(LTrim(strTest))%> <br>
strTest�� ����� ���ڿ� �� �����ʰ����� ���� : <%=RTrim(strTest)%> <br>
strTest�� ����� ���ڿ� �� �����ʰ����� ����(���ڿ� ����) : <%=Len(RTrim(strTest))%> <br>
strTest�� ����� ���ڿ� �� ���ʰ����� ���� : <%=Trim(strTest)%> <br>
strTest�� ����� ���ڿ� �� ���ʰ����� ����(���ڿ� ����) : <%=Len(Trim(strTest))%> <br>