USE [DPI]
GO
/****** Object:  Table [dbo].[QCResultLog]    Script Date: 2021/10/29 下午 04:04:56 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[QCResultLog](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[OrderId] [nchar](10) NULL,
	[OPID] [float] NULL,
	[QCReason] [nchar](10) NULL,
	[QtyNum] [int] NULL,
	[CreateDate] [datetime] NULL,
	[LastUpdateDate] [datetime] NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[QCResultLog] ADD  CONSTRAINT [DF_QCResultLog_CreateDate]  DEFAULT (getdate()) FOR [CreateDate]
GO
ALTER TABLE [dbo].[QCResultLog] ADD  CONSTRAINT [DF_QCResultLog_LastUpdateDate]  DEFAULT (getdate()) FOR [LastUpdateDate]
GO
