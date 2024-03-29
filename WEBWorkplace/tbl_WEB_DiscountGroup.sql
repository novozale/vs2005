USE [ScaDataDB]
GO
/****** Объект:  Table [dbo].[tbl_WEB_DiscountGroup]    Дата сценария: 05/29/2015 12:37:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_WEB_DiscountGroup](
	[ID] [uniqueidentifier] NOT NULL,
	[ClientCode] [nvarchar](50) NOT NULL,
	[GroupCode] [nvarchar](50) NOT NULL,
	[Discount] [decimal](18, 2) NOT NULL,
 CONSTRAINT [PK_tbl_WEB_GroupDiscount] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
