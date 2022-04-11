<%
'--------------------------------------------------------------------
' Title : ADO - Command객체
' Program Name : Command객체사용.asp
' Program Description : Command객체의 모든 멤버 사용
' Include Files : None
' Copyright (C) 2002 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------------------------
%>
<%
Dim objCmd, objRs

Set objCmd = Server.CreateObject("ADODB.Command")

'[0]Command명을 설정(Connection오브젝트에서 커맨드를 실행할 때 식별)
objCmd.Name = "objCmdName"

'[1]Command 오브젝트의 접속명을 지정(OLE DB연결문자열 작성).
objCmd.ActiveConnection = "Provider=SQLOLEDB.1;Password=Memo;Persist Security Info=True;User ID=Memo;Initial Catalog=Memo;Data Source=(local)"

'[2]CommnadText 속성에 넣을 Command의 종류를 설정(기본값 : 8(adCmdUnknown))
objCmd.CommandType = 8

'[3]Command 문자열을 설정.
objCmd.CommandText = "Select * From Memo"

'[4]Command 실행 경과 시간을 설정.(기본값 : 30초)
objCmd.CommandTimeout = 10

'[5]Command 문자열 또는 스트림(Stream)을 나타내는 GUID를 설정.
Response.Write(objCmd.Dialect&"<br>")

'[6]Command를 실행하기 전에 SQL문을 컴파일할 것인지 설정(True | False)
objCmd.Prepared = True

'[7]Parameter 오브젝트를 생성(Append메소드와 함께 사용)
'Set objPara = Command오브젝트.CreateParameter([Name],[Type],[Direction],[Size],[Value])
Set objPara = objCmd.CreateParameter("PrmID",3,1,1)
'파라미터를 추가
objCmd.Parameters.Append objPara

'[8]Command를 실행
'[Set objRs =] Command오브젝트.Execute([Records],[Para],[Type])
Set objRs = objCmd.Execute

'[9]Command 오브젝트의 상태(0:닫힘,1:열림,...)
Response.Write(objCmd.State)

Response.Write("<div align=center><table border=1>")
Do Until objRs.EOF
	Response.Write("<tr>")
	For i = 0 To objRs.Fields.Count - 1
		Response.Write("<td>" & objRs(i) & "</td>")
	Next
	Response.Write("</tr>")
	objRs.MoveNext
Loop
Response.Write("</table></div>")

'[10] Command를 취소
objCmd.Cancel

objRs.Close

Set objRs = Nothing
Set objCmd = Nothing
%>