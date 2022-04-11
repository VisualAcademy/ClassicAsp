CREATE TABLE [dbo].[my_board] (
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
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

