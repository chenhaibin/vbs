if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Logs]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Logs]
GO

CREATE TABLE [dbo].[Logs] (
	[EntryID] [char] (10) COLLATE Latin1_General_CI_AS NULL ,
	[EventClass] [char] (10) COLLATE Latin1_General_CI_AS NULL ,
	[oDateTime] [datetime] NULL ,
	[EventID] [char] (10) COLLATE Latin1_General_CI_AS NULL ,
	[Source] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[Server] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[oDescription] [varchar] (255) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

