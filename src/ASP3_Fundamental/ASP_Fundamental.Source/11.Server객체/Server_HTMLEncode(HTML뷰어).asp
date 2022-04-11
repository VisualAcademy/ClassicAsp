<html>

<!-- METADATA TYPE="typelib" FILE="c:\WinNT\System32\scrrun.dll" -->
<body>

<h4> HTML 뷰어 </h4>
<hr>

<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
<h5>파일 열기 : </h5>
<input type="file" name="filename" value="">
<input type="submit" name="load" value=" 파일 열기 "><br>
</form>

<%
    ' 파일 이름과 경로 구하기
    if Len(Request("filename")) then 
        fname = Request("filename")
    else
        fname = Server.MapPath("first.htm")
    end if
%>
<br><h5> 파일 내용 : <% = fname %> </h5>
<hr>
<%
    On Error Resume Next
    
    ' 파일 시스템 객체 생성
    Set fs = Server.CreateObject("Scripting.FileSystemObject")

    ' 파일 열기
    Set file = fs.OpenTextFile( fname, ForReading)

    ' 객체 생성 확인
    if IsObject( file ) then 
        ' 끝까지 읽는다.
        Do While Not file.AtEndOfStream   
            strLine = file.ReadLine     ' 파일에서 한 줄씩 읽어온다.
            Response.Write Server.HTMLEncode(strLine) & "<br>"
        Loop
        file.Close
    else
        Response.Write "파일 열기에 실패했습니다."
    end if
%>
</body>
</html>

