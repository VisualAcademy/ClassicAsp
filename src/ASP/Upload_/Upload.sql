use upload
go

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

--���� ������
--6.1. �Է� ��(BoardWrite.asp)
--6.2. �Է� ó��(BoardWriteProcess.asp)
Insert Upload(Name,Email,Title,PostDate,PostIP,Content, 
	Password,ReadCount,Encoding,Homepage,ModifyDate,
	ModifyIP,FileName,FileSize,DownCount)
Values ('Name','Email','Title',GetDate(),'127.0.0.1','Content', 
	'Password',0,'Text','Homepage',Null,
	Null,'FileName',100,0)
--6.3. ���(����Ʈ;BoardList.asp)
Select * From Upload Order By Num Desc
--6.4. �������(BoardView.asp)
Update Upload Set ReadCount = ReadCount + 1 Where Num = 1--��ȸ��
Select * From Upload Where Num = 1
--6.5. ���� ��(BoardModify.asp)
Select * From Upload Where Num = 1
--6.6. ���� ó��(BoardModifyProcess.asp)
Select Password From Upload Where Num = 1 And Password = 'Password'
Update Upload
Set
	Name = 'Name', 
	Email = 'Email',
	Title = 'Title',
	Content = 'Content',
	Encoding = 'HTML',
	Homepage = 'Homepage',
	ModifyDate = GetDate(),
	ModifyIP = '127.0.0.1'
Where Num = 1 And Password = 'Password'
--6.7. ���� ��(BoardDelete.asp)
--6.8. ���� ó��(BoardDeleteProcess.asp)
Select Password From Upload Where Num = 1 And Password = 'Password'
Delete Upload Where Num = 1
--6.9. �˻� ��(BoardSearch.asp)
--6.10. �˻� ó��(BoardSearchList.asp)
Select * From Upload 
Where 
	Name Like 'Name' 
	Or 
	Title Like 'Title' 
	Or 
	Content Like 'Content'
Order By Num Desc
--6.11. �ٿ�ε�(BoardDown.asp)
Update Upload Set DownCount = DownCount + 1 Where Num = 1

















