USE [fluxocaixa]
GO

/****** Object:  Table [dbo].[T_CNRIO_IMPRT_ARQV]    Script Date: 04/03/2018 15:24:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[T_CNRIO_IMPRT_ARQV](
	[ID_CNRIO_IMPRT_ARQV] [int] NOT NULL,
	[NU_CNPJ] [varchar](18) NOT NULL,
	[NU_ANO_PLAN_ORIG_PROC] [int] NOT NULL,
	[DS_CONTA_CLIE] [varchar](400) NOT NULL,
	[DS_CLSSF_PLANO_CONTA] [varchar](100) NOT NULL,
	[CD_PLANO_CONTA] [varchar](10) NOT NULL,
	[DS_PLANO_CONTA] [varchar](100) NOT NULL,
	[DS_CMNH_ARQV_ORIG] [varchar](200) NOT NULL,
	[NU_INIC_LTRA_ARQV_ORIG] [int] NOT NULL,
	[NU_FIM_LTRA_ARQV_ORIG] [int] NOT NULL,
	[CD_COL_CLSSF_PLANO_CONTA] [varchar](3) NOT NULL,
	[CD_COL_DIA] [varchar](3) NOT NULL,
	[CD_COL_DCTO_RFRC_FLUXO_CAIXA] [varchar](3) NOT NULL,
	[CD_COL_INSTT_FNCR] [varchar](3) NOT NULL,
	[CD_COL_VL_FLUXO_CAIXA] [varchar](3) NOT NULL,
	[IC_TIPO_TRANS_FLUXO_CAIXA] [varchar](3) NOT NULL,
 CONSTRAINT [PK_CNRIO_IMPRT_ARQV] PRIMARY KEY CLUSTERED 
(
	[ID_CNRIO_IMPRT_ARQV] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO