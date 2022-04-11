Create Table dbo.Vote
(
	UID Int Identity(1, 1) Not Null Primary Key, --�Ϸù�ȣ
	Title VarChar(300) Not Null, --��������
	Allow_Multi Bit Not Null, --���߼��ÿ���
	RegDate SmallDateTime Not Null Default(GetDate())--���������
)
Go
Insert Vote Values('�����ִ� ����', 0, Default)
Select * From Vote
Create Table dbo.Vote_Item
(
	UID Int Identity(1, 1) Not Null, -- �Ϸù�ȣ
	Vote_ID Int Not Null, -- Vote ���̺��� UID��
	Title VarChar(200) Not Null, -- ���׸��� ����
	Selected Int Not Null Default(0) -- ��ǥ��
)
Go
Insert Vote_Item Values(1, 'C#', 0)
Insert Vote_Item Values(1, 'ASP.NET', 0)
Insert Vote_Item Values(1, 'ASP', 0)
Select * From Vote_Item

-- ���ο� �������
Insert Vote Values('������� �����?', 1, Default)
Insert Vote_Item Values(2, '��', 0)
Insert Vote_Item Values(2, '�ƴϿ�', 0)
Insert Vote_Item Values(2, '�۽��', 0)


