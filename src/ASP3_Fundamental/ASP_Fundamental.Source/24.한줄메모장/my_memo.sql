CREATE TABLE [dbo].[my_memo] (
	[Num] [int] IDENTITY (1, 1) NOT NULL ,
	[Name] [varchar] (25) NULL ,
	[Email] [varchar] (100) NULL ,
	[Title] [varchar] (140) NULL ,
	[PostDate] [datetime] NULL 
) ON [PRIMARY]
GO

