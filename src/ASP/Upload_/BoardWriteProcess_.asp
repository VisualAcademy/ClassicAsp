<%
'���� ����
Dim objCon: Dim strSql
Dim Name: Name = Request("Name")
Dim Email: Email = Request("Email")
Dim Title: Title = Request("Title")
Dim PostIP: PostIP = Request.ServerVariables("REMOTE_HOST")
Dim Content: Content = Request("Content")
Dim Password: Password = Request("Password")
Dim Encoding: Encoding = Request("Encoding")
Dim Homepage: Homepage = Request("Homepage")
Dim FileName: FileName = Request("FileName")
Dim FileSize: FileSize = 100
'[!] ��������ǥ ó��
Name = Replace(Name, "'", "''")
Title = Replace(Title, "'", "''")
Content = Replace(Content, "'", "''")
'Ŀ�ؼ� ��ü ����
Set objCon = Server.CreateObject("ADODB.Connection")
'Ŀ�ؼ� ��ü ����
objCon.Open(Application("ConnectionString"))
'�̸� ������ SQL�� ���� : �����ѵ����� -> " & ���� & "
strSql = "Insert Upload(Name,Email,Title,PostDate,PostIP,Content,Password,ReadCount,Encoding,Homepage,ModifyDate,ModifyIP,FileName,FileSize,DownCount) Values('" & Name & "','" & Email & "','" & Title & "',GetDate(),'" & PostIP & "','" & Content & "',	'" & Password & "',0,'" & Encoding & "','" & Homepage & "',Null,Null,'" & FileName & "'," & FileSize & ",0)"
'SQL�� ����
objCon.Execute(strSql)
'Ŀ�ؼ� ��ü �ݱ�
objCon.Close()
'Ŀ�ؼ� ��ü ����
Set objCon = Nothing
'����Ʈ �������� �̵�
Response.Redirect("BoardList.asp")
%>
<center>�����Ͱ� �ԷµǾ����ϴ�.</center>