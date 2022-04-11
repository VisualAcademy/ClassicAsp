
<% 
' My ASP formatting number sample
mynumber=123.4567
Response.write "<hr>" & mynumber & "<br>"
Response.write "formatnumber(mynumber,0)" & "<br>"
Response.write formatnumber(mynumber,0) & "<hr>"
Response.write "formatnumber(mynumber,2)" & "<br>"
Response.write formatnumber(mynumber,2) & "<hr>"
Response.write "formatnumber(mynumber,6)" & "<br>"
Response.write formatnumber(mynumber,6) & "<hr>"

mynumber=.4567
Response.write mynumber & "<br>"
'0 means means no leading zeroes
Response.write "formatnumber(mynumber,2,0)" & "<br>"
Response.write formatnumber(mynumber,2,0) & "<hr>"
'1 means means pad with leading zeroes
'Response.write "formatnumber(mynumber,2,1)" & "<br>"
'Response.write formatnumber(mynumber,2,1) & "<hr>"

'mynumber=-123.4567
'Response.write mynumber & "<br>"
'0 means means no parentheses for negative numbers
'Response.write "formatnumber(mynumber,2,0,0)" & "<br>"
'Response.write formatnumber(mynumber,2,0,0) & "<hr>"
'1 means means yes parentheses for negative numbers
'Response.write "formatnumber(mynumber,2,0,1)" & "<br>"
'Response.write formatnumber(mynumber,2,0,1) & "<hr>"
%>
