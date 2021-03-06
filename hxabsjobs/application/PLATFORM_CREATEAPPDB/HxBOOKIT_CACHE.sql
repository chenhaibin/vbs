/****** Object:  Stored Procedure dbo.usp_getLookupClassData    Script Date: 04/11/2003 14:30:24 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_getLookupClassData]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_getLookupClassData]
GO

/****** Object:  Stored Procedure dbo.usp_getLookupData    Script Date: 04/11/2003 14:30:24 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_getLookupData]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_getLookupData]
GO

/****** Object:  Table [dbo].[TCOMP_MESSAGE]    Script Date: 04/11/2003 14:30:24 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TCOMP_MESSAGE]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TCOMP_MESSAGE]
GO

/****** Object:  Table [dbo].[TINTERVIEWTYPE]    Script Date: 04/11/2003 14:30:24 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TINTERVIEWTYPE]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TINTERVIEWTYPE]
GO

/****** Object:  Table [dbo].[TINTERVIEWTYPE_APPT_TYPE]    Script Date: 04/11/2003 14:30:24 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TINTERVIEWTYPE_APPT_TYPE]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TINTERVIEWTYPE_APPT_TYPE]
GO

/****** Object:  Table [dbo].[TINTERVIEWTYPE_ENCLOSURE]    Script Date: 04/11/2003 14:30:24 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TINTERVIEWTYPE_ENCLOSURE]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TINTERVIEWTYPE_ENCLOSURE]
GO

/****** Object:  Table [dbo].[TINTERVIEWTYPE_PREREQUISITE]    Script Date: 04/11/2003 14:30:24 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TINTERVIEWTYPE_PREREQUISITE]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TINTERVIEWTYPE_PREREQUISITE]
GO

/****** Object:  Table [dbo].[TLOOKUPCLASS]    Script Date: 04/11/2003 14:30:24 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TLOOKUPCLASS]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TLOOKUPCLASS]
GO

/****** Object:  Table [dbo].[TLOOKUPTABLE]    Script Date: 04/11/2003 14:30:24 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TLOOKUPTABLE]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TLOOKUPTABLE]
GO

/****** Object:  Table [dbo].[TSLOT]    Script Date: 04/11/2003 14:30:24 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TSLOT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TSLOT]
GO

/****** Object:  Table [dbo].[TCOMP_MESSAGE]    Script Date: 04/11/2003 14:30:25 ******/
CREATE TABLE [dbo].[TCOMP_MESSAGE] (
	[COMP_MESSAGE_ID] [decimal](10, 0) NOT NULL ,
	[COMP_MESSAGE] [varchar] (500) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[TINTERVIEWTYPE]    Script Date: 04/11/2003 14:30:25 ******/
CREATE TABLE [dbo].[TINTERVIEWTYPE] (
	[INTERVIEWTYPE_ID] [int] NOT NULL ,
	[COMP_MESSAGE_ID] [int] NULL ,
	[LETTERTYPE_ID] [int] NULL ,
	[CLOSED_ROOM_REQUIRED] [char] (1) COLLATE Latin1_General_CI_AS NULL ,
	[DEFAULT_DURATION_IN_SLOTS] [int] NULL ,
	[INTERVIEWTYPE_CHANNEL_NUMBER] [int] NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[TINTERVIEWTYPE_APPT_TYPE]    Script Date: 04/11/2003 14:30:25 ******/
CREATE TABLE [dbo].[TINTERVIEWTYPE_APPT_TYPE] (
	[INTERVIEWTYPE_ID] [int] NOT NULL ,
	[APPT_TYPE_ID] [int] NOT NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[TINTERVIEWTYPE_ENCLOSURE]    Script Date: 04/11/2003 14:30:25 ******/
CREATE TABLE [dbo].[TINTERVIEWTYPE_ENCLOSURE] (
	[INTERVIEWTYPE_ID] [int] NOT NULL ,
	[ENCLOSURETYPE_ID] [int] NOT NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[TINTERVIEWTYPE_PREREQUISITE]    Script Date: 04/11/2003 14:30:25 ******/
CREATE TABLE [dbo].[TINTERVIEWTYPE_PREREQUISITE] (
	[INTERVIEWTYPE_ID] [int] NOT NULL ,
	[PREREQUISITE_ID] [int] NOT NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[TLOOKUPCLASS]    Script Date: 04/11/2003 14:30:26 ******/
CREATE TABLE [dbo].[TLOOKUPCLASS] (
	[CLASS_ID] [decimal](10, 0) NOT NULL ,
	[CLASS_NAME] [varchar] (50) COLLATE Latin1_General_CI_AS NOT NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[TLOOKUPTABLE]    Script Date: 04/11/2003 14:30:26 ******/
CREATE TABLE [dbo].[TLOOKUPTABLE] (
	[CLASS_ID] [decimal](10, 0) NOT NULL ,
	[LOOKUP_CODE] [decimal](10, 0) NOT NULL ,
	[LOOKUP_DECODE] [varchar] (100) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[TSLOT]    Script Date: 04/11/2003 14:30:26 ******/
CREATE TABLE [dbo].[TSLOT] (
	[SLOT_ID] [decimal](5, 0) NOT NULL ,
	[SLOT_TIME] [char] (5) COLLATE Latin1_General_CI_AS NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[TINTERVIEWTYPE] WITH NOCHECK ADD 
	CONSTRAINT [PK_TINTERVIEWTYPE] PRIMARY KEY  CLUSTERED 
	(
		[INTERVIEWTYPE_ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[TINTERVIEWTYPE_APPT_TYPE] WITH NOCHECK ADD 
	CONSTRAINT [PK_TINTERVIEWTYPE_APPT_TYPE] PRIMARY KEY  CLUSTERED 
	(
		[INTERVIEWTYPE_ID],
		[APPT_TYPE_ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[TINTERVIEWTYPE_ENCLOSURE] WITH NOCHECK ADD 
	CONSTRAINT [PK_TINTERVIEWTYPE_ENCLOSURE] PRIMARY KEY  CLUSTERED 
	(
		[INTERVIEWTYPE_ID],
		[ENCLOSURETYPE_ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[TINTERVIEWTYPE_PREREQUISITE] WITH NOCHECK ADD 
	CONSTRAINT [PK_TINTERVIEWTYPE_PREREQUISITE] PRIMARY KEY  CLUSTERED 
	(
		[INTERVIEWTYPE_ID],
		[PREREQUISITE_ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[TLOOKUPCLASS] WITH NOCHECK ADD 
	CONSTRAINT [PK_TLOOKUPCLASS] PRIMARY KEY  CLUSTERED 
	(
		[CLASS_ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[TLOOKUPTABLE] WITH NOCHECK ADD 
	CONSTRAINT [PK_TLOOKUPTABLE] PRIMARY KEY  CLUSTERED 
	(
		[CLASS_ID],
		[LOOKUP_CODE]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[TSLOT] WITH NOCHECK ADD 
	CONSTRAINT [PK_TSLOT] PRIMARY KEY  CLUSTERED 
	(
		[SLOT_ID]
	)  ON [PRIMARY] 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.usp_getLookupClassData    Script Date: 04/11/2003 14:30:26 ******/
CREATE PROCEDURE usp_getLookupClassData

(
@varTableName AS VARCHAR(50),
@intID 	AS INTEGER
)

AS

DECLARE @varSQL AS VARCHAR(400)

SELECT @varSQL = CASE @varTableName
			WHEN 'TLOOKUPTABLE' THEN 'SELECT * FROM TLOOKUPTABLE WHERE CLASS_ID = ' + CAST(@intID AS VARCHAR(10)) + ' ORDER BY LOOKUP_CODE'

			WHEN 'TINTERVIEWTYPE' THEN 'SELECT INTERVIEWTYPE_ID FROM TINTERVIEWTYPE WHERE INTERVIEWTYPE_CHANNEL_NUMBER = ' + CAST(@intID AS VARCHAR(10)) + ' ORDER BY INTERVIEWTYPE_ID'
			
			WHEN 'TINTERVIEWTYPE_COMP' THEN 'SELECT COMP_MESSAGE_ID FROM TINTERVIEWTYPE WHERE INTERVIEWTYPE_ID = ' + CAST(@intID AS VARCHAR(10)) 

			WHEN 'TINTERVIEWTYPE_PREREQUISITE' THEN 'SELECT PREREQUISITE_ID AS ''LOOKUP_CODE'',  LOOKUP_DECODE, CLASS_ID FROM TINTERVIEWTYPE_PREREQUISITE AS TINTPRE INNER JOIN TLOOKUPTABLE AS TLOOK ON TINTPRE.PREREQUISITE_ID = TLOOK.LOOKUP_CODE WHERE TINTPRE.INTERVIEWTYPE_ID = ' + CAST(@intID AS VARCHAR(10)) + ' AND CLASS_ID=6 ORDER BY PREREQUISITE_ID'

			WHEN 'TINTERVIEWTYPE_APPT_TYPE' THEN 'SELECT APPT_TYPE_ID FROM TINTERVIEWTYPE_APPT_TYPE WHERE INTERVIEWTYPE_ID = ' + CAST(@intID AS VARCHAR(10))  + ' ORDER BY APPT_TYPE_ID'

			WHEN 'TINTERVIEWTYPE_ENCLOSURE' THEN 'SELECT ENCLOSURETYPE_ID AS ''LOOKUP_CODE'',  LOOKUP_DECODE FROM TINTERVIEWTYPE_ENCLOSURE AS TINT
											INNER JOIN TLOOKUPTABLE AS TLOOK
											ON TINT.INTERVIEWTYPE_ID = TLOOK.LOOKUP_CODE
											WHERE TINT.INTERVIEWTYPE_ID =  ' + CAST(@intID AS VARCHAR(10)) + '  AND TLOOK.CLASS_ID = 4
											ORDER BY ENCLOSURETYPE_ID'

			WHEN 'TCOMP_MESSAGE' THEN 'SELECT COMP_MESSAGE FROM TCOMP_MESSAGE WHERE COMP_MESSAGE_ID = ' + CAST(@intID AS VARCHAR(10)) + ' ORDER BY COMP_MESSAGE_ID'

			END

--EXEC sp_executesql @varSQL


EXECUTE(@varSQL)
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

/****** Object:  Stored Procedure dbo.usp_getLookupData    Script Date: 04/11/2003 14:30:26 ******/
CREATE PROCEDURE usp_getLookupData

(
@varTableName 	AS VARCHAR(255)
)

AS

IF (@varTableName = 'TSLOT')
	SELECT * FROM TSLOT

IF (@varTableName = 'TLOOKUPCLASS')
	SELECT * FROM TLOOKUPCLASS

IF (@varTableName = 'TLOOKUPTABLE')
	SELECT * FROM TLOOKUPTABLE

IF (@varTableName = 'TINTERVIEWTYPE')
	SELECT INTERVIEWTYPE_ID AS 'LOOKUP_CODE', DEFAULT_DURATION_IN_SLOTS, CLOSED_ROOM_REQUIRED, LOOKUP_DECODE, CLASS_ID FROM TINTERVIEWTYPE AS TINT
	INNER JOIN TLOOKUPTABLE AS TLOOK
	ON TINT.INTERVIEWTYPE_ID = TLOOK.LOOKUP_CODE
	WHERE TLOOK.CLASS_ID = 1
	ORDER BY INTERVIEWTYPE_ID
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

