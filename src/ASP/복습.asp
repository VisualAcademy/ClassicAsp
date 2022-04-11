<%
Dim a
Dim b
a=5
b=5
    Response.Write "a=5이고 b=5입니다.<br><br>"
    Response.Write "a=b는"&(a=b)&"입니다.<br>"
    Response.Write "a<>b는"&(a<>b)&"입니다.<br>"
    Response.Write "a<b는"& (a < b) &"입니다.<br>"
    Response.Write "a>b는"& (a > b) &"입니다.<br>"
    Response.Write "a<=b는"&(a<=b)&"입니다.<br>"
    Response.Write "a>=b는"&(a>=b)&"입니다.<br>"
%>
