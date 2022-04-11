
<%
    ' 어느 폼에 대한 입력인지 알아낸다.
    FormNo = Request.QueryString("FormNo")
    	
    Select Case FormNo

    Case 1
        Response.Write "텍스트, 암호, 텍스트 영역 입력 폼 <hr>"

        Response.Write "Text = " & Request("text") & "<br>"
        Response.Write "Pass = " & Request("pass") & "<br>"
        Response.Write "Area = " & Request("area") & "<br>"

    Case 2
        Response.Write "라디오 버튼과 첵크 박스 입력 폼 예제 <hr>"
        Response.Write Request("Radio") & "<br>"

        for each strCheck in Request("Check")
    	    Response.Write strCheck & "<br>"
        next    

    Case 3
        Response.Write "리스트와 리스트 박스 입력 폼 예제 <hr>"
        Response.Write Request("List") & "<br>"

        for each strCheck in Request("ListBox")
    	    Response.Write strCheck & "<BR>"
        next

    End Select
%>

