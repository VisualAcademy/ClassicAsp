<%
    ' �������� ���� ĳ�ø� ������ �ʴ´�.      
    'Response.Expires = 0

    ' ��Ű �� �о����
    ' ������ �湮�� �ð��� Cookie �ݷ��ǿ��� �о�´�.
    LastVisit = Request.Cookies("LastVisit")
 
    ' ��Ű �� �����ϱ�, ���� ��¥/�ð� ����
    ' �� ������ ������ �湮���� �� ������ ������ �湮 �ð�/��¥
    Response.Cookies("LastVisit") = FormatDateTime( Now )
%>

<h3> Cookie �̿� </h3>
<hr>
<%           
    If LastVisit = "" Then
        ' ó�� ���� ��� ��Ű ������ ���� 
        Response.Write "ó�� �� ���� ȯ���մϴ�."
    Else
        ' �̹� ���� ���� �ִ� ��� ������ �湮 ��¥/�ð��� ������
        Response.Write "����� " & LastVisit  & "�� �湮�� ���� �ֽ��ϴ�."
    End If 
%>
<p><A href="Response_Cookies(�������湮).asp"> ������ �ٽ� �о���� </A>

