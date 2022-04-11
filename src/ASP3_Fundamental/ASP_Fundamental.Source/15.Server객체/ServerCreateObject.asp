<%
    Dim objBrowser

    ' 객체 생성	
	'Set 객체이름 = Server.CreateObject("설계도")
    Set objBrowser = Server.CreateObject("MSWC.BrowserType") 

    ' 객체 생성 확인
    If IsObject( objBrowser ) Then

        ' 생성한 객체의 프로퍼티와 메서드 이용  %>
        사용하고 있는 브라우저는 <%=objBrowser.Browser%> <%=objBrowser.Version%> 입니다. <P> 
<%  End If %>

