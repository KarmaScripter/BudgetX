USE [DataModels]
GO

/****** Object:  Table [dbo].[SiteProjectCodes]    Script Date: 7/26/2021 6:55:06 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[SiteProjectCodes](
	[SiteProjectCodeId] [int] NOT NULL,
	[RcCode] [nvarchar](255) NULL,
	[DivisionName] [nvarchar](255) NULL,
	[SiteProjectCode] [nvarchar](255) NULL
) ON [PRIMARY]
GO

