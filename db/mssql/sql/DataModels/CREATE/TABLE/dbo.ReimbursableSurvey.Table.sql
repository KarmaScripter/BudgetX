USE [DataModels]
GO
/****** Object:  Table [dbo].[ReimbursableSurvey]    Script Date: 7/17/2021 9:04:22 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ReimbursableSurvey](
	[ReimbursableSurveyId] [int] NOT NULL,
	[BFY] [nvarchar](255) NULL,
	[FundCode] [nvarchar](255) NULL,
	[FundName] [nvarchar](255) NULL,
	[Amount] [float] NULL,
 CONSTRAINT [PK_ReimbursableSurvey] PRIMARY KEY CLUSTERED 
(
	[ReimbursableSurveyId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
