
Create Table dbo.Upload
(
	Num Int Identity(1, 1) Not Null Primary Key,  --��ȣ
	Name VarChar(25) Not Null,                    --�̸�
	Email VarChar(100) Null,                     --�̸���     
	Title VarChar(150) Not Null,                  --����
	PostDate DateTime Default GetDate() Not Null,   --�ۼ���   
	PostIP VarChar(15) Not Null,                    --�ۼ�IP
	Content Text Not Null,                         --����
	Password VarChar(20) Not Null,                 --��й�ȣ
	ReadCount Int Default 0,                    --��ȸ��
	Encoding VarChar(10) Not Null,          --���ڵ�(HTML/Text)
	Homepage VarChar(100) Null,                   --Ȩ������
	ModifyDate DateTime Null,                    --������     
	ModifyIP VarChar(15) Null,                    --����IP
	FileName VarChar(255) Null,		--���ϸ�
	FileSize Int Null,			--���ϻ�����
	DownCount Int Default 0			--�ٿ�Ƚ��
)
Go
