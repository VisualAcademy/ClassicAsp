--�⺻�� �Խ��� �ۼ��� �����ͺ��̽� ����
--Drop Database Basic
Create Database Basic
Go

--Basic�����ͺ��̽��� �̵�
Use BoardView

--Basic���̺� ����
--Drop Table dbo.Basic
Create Table dbo.Basic
(
     Num Int Identity(1, 1) Not Null Primary Key,  --��ȣ
     Name VarChar(25) Not Null,                    --�̸�
     Email VarChar(100) Null,                     --�̸���     
     Title VarChar(150) Not Null,                  --����
     PostDate DateTime Default GetDate() Not Null, --�ۼ���     
     PostIP VarChar(15) Not Null,                  --�ۼ�IP
     Content Text Not Null,                        --����
     [Password] VarChar(20) Not Null,                --��й�ȣ
     ReadCount Int Default 0,                    --��ȸ��
     Encoding VarChar(10) Not Null,              --���ڵ�(HTML/Text)
     Homepage VarChar(100) Null,                 --Ȩ������
     ModifyDate DateTime Null,                   --������     
     ModifyIP VarChar(15) Null                   --����IP
)
Go

--Write ���������� ���Ǵ� ���� ����
Insert Basic(Name,Email,Title,PostDate,PostIP,Content,Password,
ReadCount,Encoding,Homepage) 
Values('ȫ�浿','h@h.com','ȫ�浿�ٳ఩�ϴ�.',
GetDate(),'192.168.0.25', '��¼�� ��¼��...', '1234', 0,'Text', 
'http://www.h.com/')

--List ���������� ���Ǵ� ���� ����
Select * From Basic Order By Num Desc

--View ���������� ���Ǵ� ���� ����
Select * From Basic Where Num = 1

--Modify ���������� ���Ǵ� ���� ����
Update Basic Set Name = '��λ�', Email = 'b@b.com', Title = '��λ��� ������', Content = '���� ������ ����������...', Encoding = 'HTML', Homepage = 'http://www.b.com/', ModifyDate = GetDate(), ModifyIP = '192.168.0.25' Where Num = 1
--Delete
Delete Basic Where Num = 1
--Search
Select * From Basic 
Where 
	Name Like '%�̸�%' Or Title Like '%����˻�%' Or 
	Content Like '%����˻�%'



