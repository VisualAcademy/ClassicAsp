-- �⺻ �Խ���(Basic)�� ���̺� ����
Create Table dbo.Basic
(
	Num Int Identity(1, 1) Not Null Primary Key,--��ȣ
	Name VarChar(25) Not Null,				--�̸�
	Email VarChar(100) Null, 				--�̸���	
	Title VarChar(150) Not Null,				--����
	PostDate DateTime Default GetDate() Not Null,--�ۼ���	
	PostIP VarChar(15) Not Null,				--�ۼ�IP
	Content Text Not Null,					--����
	Password VarChar(20) Not Null,			--��й�ȣ
	ReadCount Int Default 0,				--��ȸ��
	Encoding VarChar(10) Not Null,	--���ڵ�(HTML/Text)
	Homepage VarChar(100) Null,				--Ȩ������
	ModifyDate DateTime Null,				--������	
	ModifyIP VarChar(15) Null				--����IP
)
Go
-- ���� ���ù�
Update Basic Set Name = 'ȫ', Email = 'h@h.com', Title = '����', Content = '����', Encoding = 'Text', Homepage = 'http://', ModifyDate = GetDate(), ModifyIP = '127.0.0.1' Where Num = 2

Select * From Basic Where Num = 1
--�Է� ���ù�
Insert Basic Values('ȫ�浿', 'h@h.com', '�ȳ�', GetDate(), '127.0.0.1', '����', '1234', 0, 'Text', 'http://', GetDate(), '127.0.0.1')
Go

Update Basic Set Name = 'aaaa', Email = 'aaa', Title = 'aaaaa', Content = 'aaaa', Encoding = 'Text', Homepage = 'aaaa', ModifyDate = GetDate(), ModifyIP = '::1' Where Num = 130
Select * From Basic
Go
