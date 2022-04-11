<html>
<body>

<!----------------------------------------------------------->
<h4> Application 객체, Contents 콜렉션 </h4>
<hr><p>

<form action="appvar.asp?mode=add" method="post">
    1. Application 변수를 추가합니다.<br>
    key : (<input type="text" name="key" size=10>,
    <input type="text" name="value" size=10>)

    <input type="submit" value="추가">	
</form>

<form action="appvar.asp?mode=remove" method="post">
    2. Application 변수를 삭제합니다.<br>
    key : <input type="text" name="key">

    <input type="submit" value="삭제">	
</form>

<form action="appvar.asp?mode=removeall" method="post">
    3. 모든 Application 변수를 지웁니다.

    <input type="submit" value="모두 삭제">	
</form>
<hr>

<!----------------------------------------------------------->
<%
    key = request("key")

    select case Request("mode")

    case "add" ' 추가
        Application.Lock
        Application(key) = request("value")
        Application.Unlock

    case "remove" ' 삭제
        Application.Lock
        Application.Contents.Remove(key)
        Application.Unlock

    case "removeall" ' 모두 삭제
        Application.lock
        Application.Contents.RemoveAll()
        Application.Unlock

    end select

    ' Application 변수 리스트 보여주기, Contents 콜렉션 루핑
    For Each item in Application.Contents

        objItem = Application.Contents(item)

        ' 객체 인지 확인
        If IsObject( objItem ) Then
            Response.Write "객체 : " & item & "<BR>"

        ' 배열 인지 확인
        ElseIf IsArray( objItem ) Then
            Response.Write "배열 : " & item & "<BR>"

            varray = Application.Contents(item)

            For i = 0 To UBound(varray)
            Response.Write "(" & i & ") " & vArray(i) & "<BR>"
            Next
        Else
            Response.Write "변수 : Application('" & item & "') : '" _
                        & Application.Contents(item) & "'<BR>"
        End If
    Next
%>

</body>    
</html>


