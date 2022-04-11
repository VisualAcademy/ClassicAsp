<%
Dim Vote_ID: Vote_ID = Request("Vote_ID")
Dim Vote_Item: Vote_Item = Request("Vote_Item")
Dim objCon, strSql

strSql = "Update Vote_Item Set Selected = Selected + 1 " & _
    "Where UID In (" & Vote_Item & ")"

Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open("Provider=SQLOLEDB.1;Password=1234;Persist Security Info=True;User ID=ASPSample;Initial Catalog=ASPSample;Data Source=(local)\SQLEXPRESS")
objCon.Execute(strSql)
objCon.Close()
Set objCon = Nothing

'ÀÌµ¿
Response.Redirect("Vote_Result.asp")
%>