<html>
<body>

<!----------------------------------------------------------->
<h4> Session ��ü, Contents �ݷ��� </h4>
<hr><p>

<form action="session.asp?mode=add" method="post">
    1. Session ������ �߰��մϴ�.<br>
    key : (<input type="text" name="key" size=10>,
    <input type="text" name="value" size=10>)

    <input type="submit" value="�߰�">	
</form>

<form action="session.asp?mode=delete" method="post">
    2. Session ������ �����մϴ�.<br>
    key : <input type="text" name="key">

    <input type="submit" value="����">	
</form>

<form action="Session.asp?mode=deleteall" method="post">
    3. ��� Session ������ ����ϴ�.

    <input type="submit" value="��� ����">	
</form>

<hr>
<!----------------------------------------------------------->
<%
    key = request("key")

    select case Request("mode")

    case "add"
        Session(key) = request("value")

    case "delete"
        Session.Contents.Remove(key)

    case "deleteall"
        Session.Contents.RemoveAll()

    end select

    ' Session ���� ����Ʈ �����ֱ�, Contents �ݷ��� ����
    For Each item in Session.Contents

        objItem = Session.Contents(item)

        ' �������� ��ü���� Ȯ��     
        If IsObject( objItem ) Then
            Response.Write "��ü : " & item & "<BR>"

        ' �������� �迭���� Ȯ��     
        ElseIf IsArray( objItem ) Then
            Response.Write "�迭 : " & item & "<BR>"

            varray = Session.Contents(item)
            For i = 0 To UBound(varray)
                Response.Write "(" & i & ") " & vArray(i) & "<BR>"
            Next

        ' �������� ������ ���
        Else
            Response.Write "���� : Session('" & item & "') : '" _
                           & Session.Contents(item) & "'<BR>"
        End If
    Next
%>
</body>    
</html>

