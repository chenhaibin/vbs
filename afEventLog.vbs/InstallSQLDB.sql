IF EXISTS (SELECT name FROM master.dbo.sysdatabases WHERE name = N'EventLogs')
	DROP DATABASE [EventLogs]
GO

CREATE DATABASE [EventLogs]  ON (NAME = N'EventLogs_Data', FILENAME = N'c:\program files\Microsoft SQL Server\MSSQL\Data\EventLogs_data.mdf' , SIZE = 10, FILEGROWTH = 10%) LOG ON (NAME = N'EventLogs_Log', FILENAME = N'c:\program files\Microsoft SQL Server\MSSQL\Data\EventLogs_Log.ldf' , SIZE = 10, FILEGROWTH = 10%)
 COLLATE Latin1_General_CI_AS
GO

exec sp_dboption N'EventLogs', N'autoclose', N'false'
GO

exec sp_dboption N'EventLogs', N'bulkcopy', N'false'
GO

exec sp_dboption N'EventLogs', N'trunc. log', N'true'
GO

exec sp_dboption N'EventLogs', N'torn page detection', N'true'
GO

exec sp_dboption N'EventLogs', N'read only', N'false'
GO

exec sp_dboption N'EventLogs', N'dbo use', N'false'
GO

exec sp_dboption N'EventLogs', N'single', N'false'
GO

exec sp_dboption N'EventLogs', N'autoshrink', N'true'
GO

exec sp_dboption N'EventLogs', N'ANSI null default', N'false'
GO

exec sp_dboption N'EventLogs', N'recursive triggers', N'false'
GO

exec sp_dboption N'EventLogs', N'ANSI nulls', N'false'
GO

exec sp_dboption N'EventLogs', N'concat null yields null', N'false'
GO

exec sp_dboption N'EventLogs', N'cursor close on commit', N'false'
GO

exec sp_dboption N'EventLogs', N'default to local cursor', N'false'
GO

exec sp_dboption N'EventLogs', N'quoted identifier', N'false'
GO

exec sp_dboption N'EventLogs', N'ANSI warnings', N'false'
GO

exec sp_dboption N'EventLogs', N'auto create statistics', N'true'
GO

exec sp_dboption N'EventLogs', N'auto update statistics', N'true'
GO

if( ( (@@microsoftversion / power(2, 24) = 8) and (@@microsoftversion & 0xffff >= 724) ) or ( (@@microsoftversion / power(2, 24) = 7) and (@@microsoftversion & 0xffff >= 1082) ) )
	exec sp_dboption N'EventLogs', N'db chaining', N'false'
GO

use [EventLogs]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Logs]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Logs]
GO

CREATE TABLE [dbo].[Logs] (
	[EntryID] [char] (10) COLLATE Latin1_General_CI_AS NULL ,
	[EventClass] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[oDateTime] [datetime] NULL ,
	[EventID] [char] (10) COLLATE Latin1_General_CI_AS NULL ,
	[Source] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[Server] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[oDescription] [varchar] (1500) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

