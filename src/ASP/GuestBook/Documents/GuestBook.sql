Create Table dbo.gBook
(
	UID Int Identity(1, 1) Not Null Primary Key, --�Ϸ�
	[User_Name] VarChar(20) Not Null, --�ۼ���
	Comment VarChar(4000) Not Null, --����
	RegDate SmallDateTime Default(GetDate()), --�ۼ���
	User_IP VarChar(15) Not Null, --IP�ּ�
	Pwd VarChar(12) Not Null	--��ȣ
)
Go

Insert gBook Values('ȫ�浿', '�ȳ�', GetDate(), '127.0.0.1', '1234')

Select * From gBook
