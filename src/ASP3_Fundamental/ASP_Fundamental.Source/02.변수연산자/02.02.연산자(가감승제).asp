<%
Option Explicit
Dim intA
Dim intB

intA = 10
intB = 3

Response.Write("intA = " & intA & "<br>")
Response.Write("intB = " & intB & "<br>")
Response.Write("<hr color='red'>")

Response.Write("intA + intB = " & (intA + intB) & "<br>")
Response.Write("intA - intB = " & (intA - intB) & "<br>")
Response.Write("intA * intB = " & (intA * intB) & "<br>")
Response.Write("intA / intB = " & (intA / intB) & "<br>")
Response.Write("intA MOD intB = " & (intA MOD intB) & "<br>")
Response.Write("intA ^ intB = " & (intA ^ intB) & "<br>")
%>
