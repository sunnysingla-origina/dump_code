/****** Object:  Table [dbo].[tbl_MostRecentLiDocumentDataDetail]    Script Date: 11/28/2018 1:21:36 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[tbl_MostRecentLiDocumentDataDetail](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[ReferenceID] [int] NULL,
	[SubReferenceID] [int] NULL,
	[Announce] [varchar](max) NULL,
	[Program_name(s)] [varchar](max) NULL,
	[Prog#] [varchar](max) NULL,
	[Comments] [varchar](max) NULL,
	[insertDateTime] [datetime] NULL,
 CONSTRAINT [PK_tbl_MostRecentLiDocumentDataDetail] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

ALTER TABLE [dbo].[tbl_MostRecentLiDocumentDataDetail] ADD  CONSTRAINT [DF_tbl_MostRecentLiDocumentDataDetail_insertDateTime]  DEFAULT (getdate()) FOR [insertDateTime]
GO


