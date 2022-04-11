<html>
<body>

<!----------------------------------------------------------->
<h4> Session 객체, Contents 콜렉션 </h4>
<hr><p>

<form action="session.asp?mode=add" method="post">
    1. Session 변수를 추가합니다.<br>
    key : (<input type="text" name="key" size=10>,
    <input type="text" name="value" size=10>)

    <input type="submit" value="추가">	
</form>

<form action="session.asp?mode=delete" method="post">
    2. Session 변수를 삭제합니다.<br>
    key : <input type="text" name="key">

    <input type="submit" value="삭제">	
</form>

<form action="Session.asp?mode=deleteall" method="post">
    3. 모든 Session 변수를 지웁니다.

    <input type="submit" value="모두 삭제">	
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

    ' Session 변수 리스트 보여주기, Contents 콜렉션 루핑
    For Each item in Session.Contents

        objItem = Session.Contents(item)

        ' 아이템이 객체인지 확인     
        If IsObject( objItem ) Then
            Response.Write "객체 : " & item & "<BR>"

        ' 아이템이 배열인지 확인     
        ElseIf IsArray( objItem ) Then
            Response.Write "배열 : " & item & "<BR>"

            varray = Session.Contents(item)
            For i = 0 To UBound(varray)
                Response.Write "(" & i & ") " & vArray(i) & "<BR>"
            Next

        ' 아이템이 변수인 경우
        Else
            Response.Write "변수 : Session('" & item & "') : '" _
                           & Session.Contents(item) & "'<BR>"
        End If
    Next
%>
</body>    
</html>

