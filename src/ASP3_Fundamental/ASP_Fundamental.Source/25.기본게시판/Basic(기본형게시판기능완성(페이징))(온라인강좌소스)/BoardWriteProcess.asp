<%
'[1] ��������
Dim objCon
Dim strSql
Dim Name: Name = Request("Name")
Dim Email: Email = Request("Email")
Dim Title: Title = Request("Title")
Dim PostIP: PostIP = Request.ServerVariables("REMOTE_HOST")
Dim Content: Content = Request("Content")
Dim Password: Password = Request("Password")
Dim Encoding: Encoding = Request("Encoding")
Dim Homepage: Homepage = Request("Homepage")
'[!] ����(Ȭ)����ǥ ó�� : '(��������ǥ �ϳ�->''(��������ǥ �ΰ�)
Name = Replace(Name, "'", "''")
Title = Replace(Title, "'", "''")
Content = Replace(Content, "'", "''")
'[2] Ŀ�ؼ� ��ü ����
Set objCon = Server.CreateObject("ADODB.Connection")
'[3] �����ͺ��̽� ����
objCon.Open("Provider=SQLOLEDB.1;Password=Basic;Persist Security Info=True;User ID=Basic;Initial Catalog=Basic;Data Source=(local)")
'[4] �̸� ������ SQL�� ����
'ȫ�浿 -> " & ��ü�Һ��� & "
strSql = "Insert Basic(Name,Email,Title,PostDate,PostIP,Content,Password,ReadCount,Encoding,Homepage) Values('" & Name & "','" & Email & "','" & Title & "',GetDate(),'" & PostIP & "', '" & Content & "', '" & Password & "', 0,'" & Encoding & "', '" & Homepage & "')"
'[5] SQL�� ����
objCon.Execute(strSql)
'[6] �ݱ�
objCon.Close()
'[7] ����
Set objCon = Nothing
'[8] �̵�
Response.Redirect("./BoardList.asp")
%>
<center>�����Ͱ� ����Ǿ����ϴ�.</center>