
<%
    ' ��� ���� ���� �Է����� �˾Ƴ���.
    FormNo = Request.QueryString("FormNo")
    	
    Select Case FormNo

    Case 1
        Response.Write "�ؽ�Ʈ, ��ȣ, �ؽ�Ʈ ���� �Է� �� <hr>"

        Response.Write "Text = " & Request("text") & "<br>"
        Response.Write "Pass = " & Request("pass") & "<br>"
        Response.Write "Area = " & Request("area") & "<br>"

    Case 2
        Response.Write "���� ��ư�� ýũ �ڽ� �Է� �� ���� <hr>"
        Response.Write Request("Radio") & "<br>"

        for each strCheck in Request("Check")
    	    Response.Write strCheck & "<br>"
        next    

    Case 3
        Response.Write "����Ʈ�� ����Ʈ �ڽ� �Է� �� ���� <hr>"
        Response.Write Request("List") & "<br>"

        for each strCheck in Request("ListBox")
    	    Response.Write strCheck & "<BR>"
        next

    End Select
%>

