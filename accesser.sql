USE [ghgl]
GO
/****** Object:  Table [dbo].[药典]    Script Date: 04/08/2017 02:06:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[药典](
	[药品类别] [varchar](20) NULL,
	[流水号] [int] NULL,
	[通用名] [varchar](100) NULL,
	[速记码] [varchar](50) NULL,
	[生成企业(总代理商)] [varchar](200) NULL,
	[药库规格] [varchar](100) NULL,
	[剂型] [varchar](50) NULL,
	[采购价] [money] NULL,
	[ID] [int] IDENTITY(1,1) NOT NULL,
 CONSTRAINT [PK_药典] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
