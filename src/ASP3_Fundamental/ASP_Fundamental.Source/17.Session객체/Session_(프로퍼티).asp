<h4>Session 타임아웃 </h4>
<hr>
    Session.SessionID(OS맘대루) : <%=Session.SessionID%> <br>
    Session.Timeout(세션유지시간) : <%=Session.Timeout%> 분 <br>

    Session.CodePage(언어) : <%= Session.CodePage%><br>
    Session.LCID(지역) : <%= Session.LCID %>

    <% Session.Timeout = 20 %>
    <br> 세션 타임 아웃을 <%= Session.Timeout %> 분으로 바꿈
