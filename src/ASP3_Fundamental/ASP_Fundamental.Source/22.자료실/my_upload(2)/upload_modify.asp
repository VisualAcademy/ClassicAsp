<%
'--------------------------------------------------
' Title : Myupload Education Version
' Program Name : upload_modify.asp
' Program Description : ���±׸� �̿��� �Է� ���� ����
' Include Files : None
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------------------
%>
<%
'������ ����� ����
Option Explicit
Dim objDB, objRS, strSQL, Num

'�Ѱ����� ���ڵ� �ѹ�(Num)���� ������ ����
Num = Request.QueryString("Num")   '���� : QueryString �÷������� �޾ƾ� ��.

'Connection��ü�� �ν��Ͻ� ����
Set objDB = Server.CreateObject("ADODB.Connection")

'�����ͺ��̽� ����
objDB.Open("dsn=my_database_dsn;uid=my_database_id;pwd=my_database_pwd;")

'���ڵ� �� ��ü�� �ν��Ͻ� ����
Set objRS = Server.CreateObject("Adodb.RecordSet")

'SQL�� ����
strSQL = "Select * From my_upload Where Num = " & Num
'Response.Write strSQL

'���ڵ� �� ����
objRS.Open strSQL, objDB
%>
<html>
<head>
	<title>Myupload �����ϱ�</title>

<script language="Javascript">
	<!--

		function check()   //MyMemo �� MyGuest�� �ٸ� ������� ����(onSubmit �̺�Ʈ�ڵ鷯 ���)
		{
			// ������ �ش� ��ü ���� ����
			var valTitle = document.WriteForm.Title.value;
			var valContent = document.WriteForm.Content.value;
			var valName = document.WriteForm.Name.value;
			var valPassword = document.WriteForm.Password.value;
			
			if(valName == "") {
				alert("'�ۼ���'�� �Է��� �ֽʽÿ�.\n");
				document.WriteForm.Name.focus();
				document.WriteForm.Name.select();
				return false;
			}
				
			if(valTitle == "" || valTitle == " ") {
				alert("'����'�� �Է��� �ֽʽÿ�.\n");
				document.WriteForm.Title.focus();
				document.WriteForm.Title.select();
				return false;
			}

			if(valContent == "") {
				alert("'����'�� �Է��� �ֽʽÿ�.\n");
				document.WriteForm.Content.focus();
				document.WriteForm.Content.select();
				return false;
			}
			
			if(valPassword == "") {
				alert("'��ȣ'�� �Է��� �ֽʽÿ�.\n");
				document.WriteForm.Password.focus();
				document.WriteForm.Password.select();
				return false;
			}

			return true;   //������ �����ϸ� submit

		} // end function

	//-->
</script>

</head>
<body>

<div align="center">
<table border="1">
<form action="./upload_modify_process.asp" method="Post" name="WriteForm" OnSubmit="return check();">
	<input type="hidden" name="num" value="<%=Num%>">
	<tr>
		<td align="center" colspan="3">�ðԽ��� �� �����ϱ��</td>
	</tr>
	<tr>
		<td align="center">
			�ۼ��� :
		</td>
		<td colspan="2">
			<input type="text" name="Name" size="15" maxlength="15" value="<%=objRS.Fields("Name")%>">
		</td>
	</tr>
	<tr>
		<td align="center">
			E-mail :
		</td>
		<td colspan="2">
			<input type="text" name="Email" size="30" maxlength="50" value="<%=objRS.Fields("Email")%>">
		</td>
	</tr>
	<tr>
		<td align="center">
			��&nbsp;&nbsp;&nbsp;��:
		</td>
		<td colspan="2">
			<input type="text" name="Title" size="60" maxlength="70" value="<%=objRS.Fields("Title")%>">
		</td>
	</tr>
	<tr>
		<td align="center">
			��&nbsp;&nbsp;&nbsp;��:
		</td>
		<td colspan="2">
			<textarea name="Content" cols="60" rows="10"><%=objRS.Fields("Content")%></textarea>
		</td>
	</tr>
	<tr>
		<td colspan="3" align="center">
		<tt>
			��á�ܡݡۡߢ¡ޢ������������آӢԢѢСڡ٢ŢĢ������� <br>
			�΢Ϣ͡עߢޢ����ݢܢۢաϡΡ����Ţ�����⡺������ 
		</tt>
		</td>
	</tr>
	<tr>
		<td align="center">
			Encoding:
		</td>
		<td colspan="2">
			<select name="Encoding">
				<option <% If objRS.Fields("Encoding") = "Text" Then %>selected<% End If %>>Text
				<option <% If objRS.Fields("Encoding") = "HTML" Then %>selected<% End If %>>HTML
			</select>
		</td>
	</tr>
	<tr>
		<td colspan="3">
			&nbsp;Encoding Type :<br>
			&nbsp;&nbsp;&nbsp;-	'Text' �� HTML �ڵ带 �������� �ʰ� ���� �״�θ� �����ݴϴ�.<br>
			&nbsp;&nbsp;&nbsp;-	'HTML' �� HTML �ڵ带 ������ ����� �����ݴϴ�.
		</td>
	</tr>
	<tr>
		<td valign="middle" align="center">
			��&nbsp;&nbsp;&nbsp;ȣ:
		</td>
		<td>
			<input type="password" name="Password" size="10" maxlength="10" >
		</td>
		<td align="center">
			<input type="submit" value="���� �Ϸ�">
		</td>
	</tr>
	<tr>
		<td align="center" colspan="3">
			<a href="./upload_list.asp">[�Խ��� ����Ʈ ����]</a>
		</td>
	</tr>	
</form>
</table>
</div>

</body>
</html>