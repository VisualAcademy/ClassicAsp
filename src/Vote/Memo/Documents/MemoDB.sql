-- ���� �޸��� ���̺� ����
Create Table dbo.Memos
(
	Num Int
		Identity(1, 1) Not Null Primary Key,	-- ��ȣ
	Name VarChar(25) Not Null,					-- �̸�
	Email VarChar(50) Null,						-- �̸���
	Title VarChar(150) Not Null,				-- ���ٸ޸�(����)
	PostDate SmallDateTime Default(GetDate()),	-- �ۼ���
	PostIP VarChar(15) Null						-- IP�ּ�	
)
Go
