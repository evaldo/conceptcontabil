USE [fluxocaixa]
GO
/****** Object:  Table [dbo].[T_FLUXO_CAIXA]    Script Date: 10/01/2018 15:57:00 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[T_FLUXO_CAIXA](
	[ID_FLUXO_CAIXA] [int] NOT NULL,
	[ID_CLSSF_PLANO_CONTA] [int] NULL,
	[NU_CNPJ] [varchar](18) NOT NULL,
	[SK_DMSAO_TEMPO] [int] NOT NULL,
	[DT_MVMT_FLUXO_CAIXA] [date] NOT NULL,
	[DS_CLSSF_PLANO_CONTA] [varchar](100) NOT NULL,
	[CD_DCTO_RFRC_FLUXO_CAIXA] [varchar](20) NOT NULL,
	[CD_PLANO_CONTA] [varchar](10) NOT NULL,
	[DS_PLANO_CONTA] [varchar](100) NOT NULL,
	[DS_INSTT_FNCR] [varchar](50) NOT NULL,
	[VL_ENTR_FLUXO_CAIXA] [decimal](11, 2) NULL,
	[VL_SAIDA_FLUXO_CAIXA] [decimal](11, 2) NULL,
	[IC_STATUS_VALOR] [varchar](20) NOT NULL,
	[NU_MATL_INCS] [varchar](20) NULL,
	[DT_INCS] [datetime] NULL,
	[IC_TIPO_TRANS_FLUXO_CAIXA] [varchar](1) NULL,
	[DS_PLAN_ORIG_PROC] [varchar](50) NULL,
	[NU_ANO_PLAN_ORIG_PROC] [int] NOT NULL,
	[CD_CLSSF_PLANO_CONTA] [varchar](10) NULL,
	[NM_CLIE_FLUXO_CAIXA] [varchar](200) NULL,
 CONSTRAINT [PK_FLUXO_CAIXA] PRIMARY KEY CLUSTERED 
(
	[ID_FLUXO_CAIXA] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[T_FLUXO_CAIXA]  WITH CHECK ADD  CONSTRAINT [FK_FLUXO_CAIXA_DMSAO_TEMPO] FOREIGN KEY([SK_DMSAO_TEMPO])
REFERENCES [dbo].[T_DMSAO_TEMPO] ([ID_DMSAO_TEMPO])
GO
ALTER TABLE [dbo].[T_FLUXO_CAIXA] CHECK CONSTRAINT [FK_FLUXO_CAIXA_DMSAO_TEMPO]
GO
ALTER TABLE [dbo].[T_FLUXO_CAIXA]  WITH CHECK ADD  CONSTRAINT [FK_FLUXO_CAIXA_PLANO_CONTA] FOREIGN KEY([ID_CLSSF_PLANO_CONTA])
REFERENCES [dbo].[T_CLSSF_PLANO_CONTA] ([ID_CLSSF_PLANO_CONTA])
GO
ALTER TABLE [dbo].[T_FLUXO_CAIXA] CHECK CONSTRAINT [FK_FLUXO_CAIXA_PLANO_CONTA]
GO
