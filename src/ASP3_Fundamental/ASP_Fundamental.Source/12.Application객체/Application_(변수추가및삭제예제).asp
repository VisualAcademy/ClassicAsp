<html>
<body>

<!----------------------------------------------------------->
<h4> Application ��ü, Contents �ݷ��� </h4>
<hr><p>

<form action="appvar.asp?mode=add" method="post">
    1. Application ������ �߰��մϴ�.<br>
    key : (<input type="text" name="key" size=10>,
    <input type="text" name="value" size=10>)

    <input type="submit" value="�߰�">	
</form>

<form action="appvar.asp?mode=remove" method="post">
    2. Application ������ �����մϴ�.<br>
    key : <input type="text" name="key">

    <input type="submit" value="����">	
</form>

<form action="appvar.asp?mode=removeall" method="post">
    3. ��� Application ������ ����ϴ�.

    <input type="submit" value="��� ����">	
</form>
<hr>

<!----------------------------------------------------------->
<%
    key = request("key")

    select case Request("mode")

    case "add" ' �߰�
        Application.Lock
        Application(key) = request("value")
        Application.Unlock

    case "remove" ' ����
        Application.Lock
        Application.Contents.Remove(key)
        Application.Unlock

    case "removeall" ' ��� ����
        Application.lock
        Application.Contents.RemoveAll()
        Application.Unlock

    end select

    ' Application ���� ����Ʈ �����ֱ�, Contents �ݷ��� ����
    For Each item in Application.Contents

        objItem = Application.Contents(item)

        ' ��ü ���� Ȯ��
        If IsObject( objItem ) Then
            Response.Write "��ü : " & item & "<BR>"

        ' �迭 ���� Ȯ��
        ElseIf IsArray( objItem ) Then
            Response.Write "�迭 : " & item & "<BR>"

            varray = Application.Contents(item)

            For i = 0 To UBound(varray)
            Response.Write "(" & i & ") " & vArray(i) & "<BR>"
            Next
        Else
            Response.Write "���� : Application('" & item & "') : '" _
                        & Application.Contents(item) & "'<BR>"
        End If
    Next
%>

</body>    
</html>


