* SQL Server 2000

1. ��ġ/����

2. DB ���� : my_database

3. Table���� :	my_ado

���̺� ����

CREATE TABLE [dbo].[my_ado] (
	[Num] [int] IDENTITY (1, 1) NOT NULL ,
	[Name] [varchar] (25) NULL ,
	[Email] [varchar] (100) NULL ,
	[Title] [varchar] (140) NULL ,
	[PostDate] [datetime] NULL ,
	[ModifyDate] [datetime] NULL ,
	[ReadCount] [int] NULL ,
	[PostIP] [varchar] (15) NULL ,
	[ModifyIP] [varchar] (15) NULL ,
	[Encoding] [varchar] (10) NULL ,
	[Passwd] [varchar] (20) NULL ,
	[Content] [text] NULL 
) 

4. ����� ���� 
	UID : my_database_id
	PWD : my_database_pwd

5. OLE DB���� : script������ �ִ� .UDL���� ����...

6. ���ǿ� �⺻ �Խ��� �ٿ� -> my_ado��� ��Ī���� ������͸� ��´�.

7. ASP ���� : http://localhost/my_ado/ado_list.asp

