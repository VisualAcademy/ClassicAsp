CREATE TABLE [dbo].[my_upload] (
	[Num] [int] IDENTITY (1, 1) NOT NULL ,
	[Name] [varchar] (25) NULL ,
	[Email] [varchar] (100) NULL ,
	[Title] [varchar] (140) NULL ,
	[FileName] [varchar] (255) NULL ,
	[FileSize] [int] NULL ,
	[PostDate] [datetime] NULL ,
	[ModifyDate] [datetime] NULL ,
	[ReadCount] [int] NULL ,
	[DownCount] [int] NULL ,
	[PostIP] [varchar] (15) NULL ,
	[ModifyIP] [varchar] (15) NULL ,
	[Encoding] [varchar] (10) NULL ,
	[Passwd] [varchar] (20) NULL ,
	[Ref] [int] NULL ,
	[Step] [int] NULL ,
	[RefOrder] [int] NULL ,
	[Content] [text] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

