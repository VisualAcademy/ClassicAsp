<html>

<!-- METADATA TYPE="typelib" FILE="c:\WinNT\System32\scrrun.dll" -->
<body>

<h4> HTML ��� </h4>
<hr>

<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
<h5>���� ���� : </h5>
<input type="file" name="filename" value="">
<input type="submit" name="load" value=" ���� ���� "><br>
</form>

<%
    ' ���� �̸��� ��� ���ϱ�
    if Len(Request("filename")) then 
        fname = Request("filename")
    else
        fname = Server.MapPath("first.htm")
    end if
%>
<br><h5> ���� ���� : <% = fname %> </h5>
<hr>
<%
    On Error Resume Next
    
    ' ���� �ý��� ��ü ����
    Set fs = Server.CreateObject("Scripting.FileSystemObject")

    ' ���� ����
    Set file = fs.OpenTextFile( fname, ForReading)

    ' ��ü ���� Ȯ��
    if IsObject( file ) then 
        ' ������ �д´�.
        Do While Not file.AtEndOfStream   
            strLine = file.ReadLine     ' ���Ͽ��� �� �پ� �о�´�.
            Response.Write Server.HTMLEncode(strLine) & "<br>"
        Loop
        file.Close
    else
        Response.Write "���� ���⿡ �����߽��ϴ�."
    end if
%>
</body>
</html>

