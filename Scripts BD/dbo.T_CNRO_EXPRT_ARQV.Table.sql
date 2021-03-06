USE [fluxocaixa]
GO

/****** Object:  Table [dbo].[T_CNRO_EXPRT_ARQV]    Script Date: 04/03/2018 15:25:44 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[T_CNRO_EXPRT_ARQV](
	[ID_CNRO_EXPRT] [int] NOT NULL,
	[CD_INSTT_FNCR] [varchar](50) NULL,
	[DS_INSTT_FNCR] [varchar](50) NULL,
	[CD_DCTO_RFRC_FLUXO_CAIXA] [varchar](50) NULL,
	[DS_DCTO_RFRC_FLUXO_CAIXA] [varchar](50) NULL,
	[NU_ANO_PLAN_ORIG_PROC] [int] NULL,
	[NU_CNPJ] [varchar](18) NULL,
	[TP_CNRO_EXPRT] [varchar](10) NULL
) ON [PRIMARY]
GO
