
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_NULLS ON 
GO

-- =============================================
-- Criando Banco de Dados
-- =============================================
IF EXISTS (SELECT * 
	   FROM   master..sysdatabases 
	   WHERE  name = 'Cinema')
	DROP DATABASE cinema	
GO

CREATE DATABASE cinema
GO

Use Cinema
Go

-- =============================================
-- Criando Tabelas VideoHall
-- =============================================
CREATE TABLE [dbo].[tb_Temp_Lotacao](
	[codFilme] [int] NULL,
	[codSala] [int] NULL,
	[sessao] [smallint] NULL,
	[horario] [datetime] NULL,
	[lotada] [bit] NOT NULL,
	[Data] [date] NULL
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_trailer](
	[descricao] [nvarchar](50) NULL,
	[arquivo] [nvarchar](max) NULL,
	[Codigo] [numeric](18, 0) IDENTITY(1,1) NOT NULL
) ON [PRIMARY]

GO

CREATE TABLE [dbo].[tb_imagens](
	[descricao] [nvarchar](50) NULL,
	[arquivo] [nvarchar](max) NULL
) ON [PRIMARY]

GO

CREATE TABLE [dbo].[tb_parametros](
	[transicao] [int] NULL,
	[intermitencia] [int] NULL,
	[velocMsg] [int] NULL,
	[vendaAndes] [int] NULL,
	[vendaDepois] [int] NULL,
	[hrLimitePeriodo] [datetime] NULL,
	[telaSessoes] [bit] NOT NULL,
	[telaFilme] [bit] NOT NULL,
	[telaPrecos] [bit] NOT NULL,
	[telaTrailer] [bit] NOT NULL,
	[telaImagem] [bit] NOT NULL,
	[mensagem] [nvarchar](max) NULL,
	[corFundT1Filme] [int] NULL,
	[corTextT1Filme] [int] NULL,
	[corFundT1Hora] [int] NULL,
	[corTextT1Hora] [int] NULL,
	[corFundT1Tutulo1] [int] NULL,
	[corTextT1Tutulo1] [int] NULL,
	[corFundT1Titulo2] [int] NULL,
	[corTextT1Titulo2] [int] NULL,
	[corFundT1Lin1] [int] NULL,
	[corTextT1Lin1] [int] NULL,
	[corFundT1Lin2] [int] NULL,
	[corTextT1Lin2] [int] NULL,
	[corFundT1Mensagem] [int] NULL,
	[corTextT1Mensagem] [int] NULL,
	[corFundT2Filme1] [int] NULL,
	[corTextT2Filme1] [int] NULL,
	[corFundT2Filme2] [int] NULL,
	[corTextT2Filme2] [int] NULL,
	[corFundT2Titulo1] [int] NULL,
	[corTextT2Titulo1] [int] NULL,
	[corFundT2Titulo2] [int] NULL,
	[corTextT2Titulo2] [int] NULL,
	[corFundT2Sessao1] [int] NULL,
	[corTextT2Sessao1] [int] NULL,
	[corFundT2Sessao2] [int] NULL,
	[corTextT2Sessao2] [int] NULL,
	[corFundT2Sala1] [int] NULL,
	[corTextT2Sala1] [int] NULL,
	[corFundT2Sala2] [int] NULL,
	[corTextT2Sala2] [int] NULL,
	[corFundT2Sessoes1] [int] NULL,
	[corTextT2Sessoes1] [int] NULL,
	[corFundT2Sessoes1L1] [int] NULL,
	[corTextT2Sessoes1L1] [int] NULL,
	[corFundT2Sessoes1L2] [int] NULL,
	[corTextT2Sessoes1L2] [int] NULL,
	[corFundT2Sessoes2] [int] NULL,
	[corTextT2Sessoes2] [int] NULL,
	[corFundT2Sessoes2L1] [int] NULL,
	[corTextT2Sessoes2L1] [int] NULL,
	[corFundT2Sessoes2L2] [int] NULL,
	[corTextT2Sessoes2L2] [int] NULL,
	[corFundT2Mensagem] [int] NULL,
	[corTextT2Mensagem] [int] NULL,
	[corFundT3Hora] [int] NULL,
	[corTextT3Hora] [int] NULL,
	[corFundT3Data] [int] NULL,
	[corTextT3Data] [int] NULL,
	[corFundT3TituloTela] [int] NULL,
	[corTextT3TituloTela] [int] NULL,
	[corFundT3Titulo] [int] NULL,
	[corTextT3Titulo] [int] NULL,
	[corFundT3Lin1] [int] NULL,
	[corTextT3Lin1] [int] NULL,
	[corFundT3Lin2] [int] NULL,
	[corTextT3Lin2] [int] NULL,
	[corFundT3Mensagem] [int] NULL,
	[corTextT3Mensagem] [int] NULL,
	[corFundLotado] [int] NULL,
	[corTextLotado] [int] NULL
) ON [PRIMARY]

GO

  INSERT INTO [tb_parametros](
           [transicao]
           ,[intermitencia]
           ,[velocMsg]
           ,[vendaAndes]
           ,[vendaDepois]
           ,[hrLimitePeriodo]
           ,[telaSessoes]
           ,[telaFilme]
           ,[telaPrecos]
           ,[telaTrailer]
           ,[telaImagem]
           ,[mensagem]
           ,[corFundT1Filme]
           ,[corTextT1Filme]
           ,[corFundT1Hora]
           ,[corTextT1Hora]
           ,[corFundT1Tutulo1]
           ,[corTextT1Tutulo1]
           ,[corFundT1Titulo2]
           ,[corTextT1Titulo2]
           ,[corFundT1Lin1]
           ,[corTextT1Lin1]
           ,[corFundT1Lin2]
           ,[corTextT1Lin2]
           ,[corFundT1Mensagem]
           ,[corTextT1Mensagem]
           ,[corFundT2Filme1]
           ,[corTextT2Filme1]
           ,[corFundT2Filme2]
           ,[corTextT2Filme2]
           ,[corFundT2Titulo1]
           ,[corTextT2Titulo1]
           ,[corFundT2Titulo2]
           ,[corTextT2Titulo2]
           ,[corFundT2Sessao1]
           ,[corTextT2Sessao1]
           ,[corFundT2Sessao2]
           ,[corTextT2Sessao2]
           ,[corFundT2Sala1]
           ,[corTextT2Sala1]
           ,[corFundT2Sala2]
           ,[corTextT2Sala2]
           ,[corFundT2Sessoes1]
           ,[corTextT2Sessoes1]
           ,[corFundT2Sessoes1L1]
           ,[corTextT2Sessoes1L1]
           ,[corFundT2Sessoes1L2]
           ,[corTextT2Sessoes1L2]
           ,[corFundT2Sessoes2]
           ,[corTextT2Sessoes2]
           ,[corFundT2Sessoes2L1]
           ,[corTextT2Sessoes2L1]
           ,[corFundT2Sessoes2L2]
           ,[corTextT2Sessoes2L2]
           ,[corFundT2Mensagem]
           ,[corTextT2Mensagem]
           ,[corFundT3Hora]
           ,[corTextT3Hora]
           ,[corFundT3Data]
           ,[corTextT3Data]
           ,[corFundT3TituloTela]
           ,[corTextT3TituloTela]
           ,[corFundT3Titulo]
           ,[corTextT3Titulo]
           ,[corFundT3Lin1]
           ,[corTextT3Lin1]
           ,[corFundT3Lin2]
           ,[corTextT3Lin2]
           ,[corFundT3Mensagem]
           ,[corTextT3Mensagem]
           ,[corFundLotado]
           ,[corTextLotado])
     VALUES     
		(5,	1000,	50,		20,				20,		convert(time,'17:00:00.000',103),	1,	1,	1,	0,	0,	'SEJAM BEM VINDOS',	33023,	16777215,
		33023,	16777215,	33023,	16777215,	12632256,	16777215,	8421504,	16777215,	12632256,	16777215,	0,	16777215,	33023,	16777215,	33023,
		16777215,	12632256,	16777215,	8421504,	16777215,	33023,	16777215,	33023,	16777215,	33023,	16777215,	8421504,	16777215,	33023,
		16777215,	12632256,	16777215,	8421504,	16777215,	33023,	16777215,	12632256,	16777215,	8421504,	16777215,	0,	16777215,	33023,
		16777215,	33023,	16777215,	33023,	16777215,	33023,	16777215,	12632256,	16777215,	8421504,	16777215,	0,	16777215,	986895,	16777215)
GO



-- =============================================
-- Criando Tabelas Cinema
-- =============================================
 
 CREATE TABLE tb_empresa (
        emp_cd               int NOT NULL,
        emp_nm               varchar(50) NOT NULL,
        emp_cnpj             char(14) NOT NULL,
        emp_inscricao        char(12) NULL,
        emp_end              varchar(50) NULL,
        emp_num_end          int NULL,
        emp_cmp_end          varchar(20) NULL,
        emp_brr_end          varchar(50) NULL,
        emp_cid_end          varchar(50) NULL,
        emp_dt_inc           datetime NOT NULL,
        emp_uf_end           varchar(2) NULL,
        emp_cep_end          char(8) NULL,
        emp_dt_alt           datetime NULL,
        emp_dt_des           datetime NULL,
        emp_tel              varchar(20) NULL,
        emp_mot_des          varchar(50) NULL,
        PRIMARY KEY (emp_cd)
 )
go
 
 
 CREATE TABLE tb_cinema (
        cin_cd               int NOT NULL,
        emp_cd               int NOT NULL,
        cin_nm               varchar(50) NOT NULL,
        cin_cnpj             char(14) NOT NULL,
        cin_inscricao        char(12) NULL,
        cin_end              varchar(50) NULL,
        cin_num_end          int NULL,
        cin_cmp_end          varchar(20) NULL,
        cin_brr_end          varchar(50) NULL,
        cin_cid_end          varchar(50) NULL,
        cin_uf_end           varchar(2) NULL,
        cin_cep_end          char(8) NULL,
        cin_tel              varchar(20) NULL,
        cin_dt_inc           datetime NOT NULL,
        cin_dt_alt           datetime NULL,
        cin_dt_des           datetime NULL,
        cin_mot_des          varchar(50) NULL,
        PRIMARY KEY (cin_cd), 
        FOREIGN KEY (emp_cd)
                              REFERENCES tb_empresa
 )
go
 
 CREATE INDEX XIF34tb_cinema ON tb_cinema
 (
        emp_cd
 )
go
 
 
 CREATE TABLE tb_sala (
        sal_cd               int IDENTITY,
        cin_cd               int NOT NULL,
        sal_desc             varchar(50) NOT NULL,
        sal_lugares          smallint NOT NULL,
        sal_dt_inc           datetime NOT NULL,
        sal_dt_alt           datetime NULL,
        sal_dt_des           datetime NULL,
        sal_mot_des          varchar(50) NULL,
        PRIMARY KEY (sal_cd), 
        FOREIGN KEY (cin_cd)
                              REFERENCES tb_cinema
 )
go
 
 CREATE INDEX XIF12tb_sala ON tb_sala
 (
        cin_cd
 )
go
 
 
 CREATE TABLE tb_poltronas (
        sal_cd               int NOT NULL,
        pol_tp_numeracao     int NULL,
        pol_num_pri_col      varchar(4) NULL,
        pol_num_filas        int NULL,
        pol_num_colunas      int NULL,
        pol_num_horiz        int NULL,
        pol_num_vert         int NULL,
        pol_poltronas        int NULL,
        pol_mat_poltr        varchar(2704) NULL,
        PRIMARY KEY (sal_cd), 
        FOREIGN KEY (sal_cd)
                              REFERENCES tb_sala
 )
go
 
 
 CREATE TABLE tb_distribuidora (
        dis_cd               int NOT NULL,
        dis_nm               varchar(50) NULL,
        PRIMARY KEY (dis_cd)
 )
go
 
 
 CREATE TABLE tbUltimoAcesso (
        ultimoAcesso         varchar(14) NOT NULL
 )
go
 
 
 CREATE TABLE tb_bol_catraca (
        bol_dt_mov           datetime NOT NULL,
        emp_cd               int NOT NULL,
        cin_cd               int NOT NULL,
        cat_cd               int NOT NULL,
        cat_nm               varchar(50) NULL,
        ctc_ini_cont         int NOT NULL,
        ctc_fim_cont         int NULL,
        PRIMARY KEY (bol_dt_mov, emp_cd, cin_cd, cat_cd)
 )
go
 
 
 CREATE TABLE tb_bol_sala (
        bol_dt_mov           datetime NOT NULL,
        emp_cd               int NOT NULL,
        cin_cd               int NOT NULL,
        sal_cd               int NOT NULL,
        sal_desc             varchar(50) NOT NULL,
        sal_lugares          smallint NOT NULL,
        PRIMARY KEY (bol_dt_mov, emp_cd, cin_cd, sal_cd)
 )
go
 
 
 CREATE TABLE tb_bol_catraca_sala (
        bol_dt_mov           datetime NOT NULL,
        emp_cd               int NOT NULL,
        cin_cd               int NOT NULL,
        sal_cd               int NOT NULL,
        cat_cd               int NOT NULL,
        PRIMARY KEY (bol_dt_mov, emp_cd, cin_cd, sal_cd, cat_cd), 
        FOREIGN KEY (bol_dt_mov, emp_cd, cin_cd, cat_cd)
                              REFERENCES tb_bol_catraca, 
        FOREIGN KEY (bol_dt_mov, emp_cd, cin_cd, sal_cd)
                              REFERENCES tb_bol_sala
 )
go
 
 CREATE INDEX XIF88tb_bol_catraca_sala ON tb_bol_catraca_sala
 (
        bol_dt_mov,
        sal_cd,
        emp_cd,
        cin_cd
 )
go
 
 CREATE INDEX XIF89tb_bol_catraca_sala ON tb_bol_catraca_sala
 (
        bol_dt_mov,
        cat_cd,
        emp_cd,
        cin_cd
 )
go
 
 
 CREATE TABLE tb_aux_boletim (
        bol_dt_mov           datetime NOT NULL,
        PRIMARY KEY (bol_dt_mov)
 )
go
 
 
 CREATE TABLE tb_bol_param (
        bol_dt_mov           datetime NOT NULL,
        emp_cd               int NOT NULL,
        cin_cd               int NOT NULL,
        par_hora_max_ses     datetime NULL,
        par_hora_limite      datetime NULL,
        par_custo_ingresso   money NULL,
        par_imposto_mun      money NULL,
        par_direitos_aut     money NULL,
        par_outros           money NULL,
        par_hora_limite23    datetime NULL,
        par_hora_limite34    datetime NULL,
        par_hora_limite45    datetime NULL,
        par_hora_limite56    datetime NULL,
        par_hora_limite12    datetime NULL,
        par_perc_meias       money NOT NULL,
        par_perc_cortesias   money NOT NULL,
        PRIMARY KEY (bol_dt_mov, emp_cd, cin_cd)
 )
go
 
 
 CREATE TABLE tb_bol_distrib (
        bol_dt_mov           datetime NOT NULL,
        dis_cd               int NOT NULL,
        dis_nm               varchar(50) NULL,
        PRIMARY KEY (bol_dt_mov, dis_cd)
 )
go
 
 
 CREATE TABLE tb_bol_filme (
        bol_dt_mov           datetime NOT NULL,
        emp_cd               int NOT NULL,
        cin_cd               int NOT NULL,
        fil_cd               int NOT NULL,
        dis_cd               int NULL,
        fil_nm               varchar(50) NOT NULL,
        fil_dt_ini           datetime NOT NULL,
        fil_dt_fim           datetime NOT NULL,
        fil_censura          smallint NOT NULL,
        fil_id_nacio         varchar(1) NOT NULL,
        PRIMARY KEY (bol_dt_mov, emp_cd, cin_cd, fil_cd), 
        FOREIGN KEY (bol_dt_mov, dis_cd)
                              REFERENCES tb_bol_distrib
 )
go
 
 CREATE INDEX XIF91tb_bol_filme ON tb_bol_filme
 (
        bol_dt_mov,
        dis_cd
 )
go
 
 
 CREATE TABLE tb_bol_cin (
        bol_dt_mov           datetime NOT NULL,
        emp_cd               int NOT NULL,
        cin_cd               int NOT NULL,
        cin_nm               varchar(50) NOT NULL,
        cin_cnpj             char(14) NOT NULL,
        cin_end              varchar(50) NULL,
        cin_num_end          int NULL,
        cin_cmp_end          varchar(20) NULL,
        cin_brr_end          varchar(50) NULL,
        cin_cid_end          varchar(50) NULL,
        cin_uf_end           varchar(2) NULL,
        cin_cep_end          char(8) NULL,
        PRIMARY KEY (bol_dt_mov, emp_cd, cin_cd)
 )
go
 
 
 CREATE TABLE tb_bol_empr (
        bol_dt_mov           datetime NOT NULL,
        emp_cd               int NOT NULL,
        emp_nm               varchar(50) NOT NULL,
        PRIMARY KEY (bol_dt_mov, emp_cd)
 )
go
 
 
 CREATE TABLE tb_boletim (
        bol_dt_mov           datetime NOT NULL,
        emp_cd               int NOT NULL,
        cin_cd               int NOT NULL,
        sal_cd               int NOT NULL,
        fil_cd               int NOT NULL,
        bol_dt_abertura      datetime NOT NULL,
        bol_dt_emissao       datetime NOT NULL,
        bol_dt_ini_per       datetime NULL,
        bol_dt_fim_per       datetime NULL,
        bol_status           varchar(1) NULL,
        PRIMARY KEY (bol_dt_mov, emp_cd, cin_cd, sal_cd, fil_cd), 
        FOREIGN KEY (bol_dt_mov, emp_cd, cin_cd)
                              REFERENCES tb_bol_param, 
        FOREIGN KEY (bol_dt_mov, emp_cd, cin_cd, fil_cd)
                              REFERENCES tb_bol_filme, 
        FOREIGN KEY (bol_dt_mov, emp_cd, cin_cd, sal_cd)
                              REFERENCES tb_bol_sala, 
        FOREIGN KEY (bol_dt_mov, emp_cd, cin_cd)
                              REFERENCES tb_bol_cin, 
        FOREIGN KEY (bol_dt_mov, emp_cd)
                              REFERENCES tb_bol_empr
 )
go
 
 
 CREATE TABLE tb_ctrl_boletim (
        bol_dt_mov           datetime NOT NULL,
        emp_cd               int NOT NULL,
        cin_cd               int NOT NULL,
        sal_cd               int NOT NULL,
        fil_cd               int NOT NULL,
        dt_envio             datetime NOT NULL,
        PRIMARY KEY (bol_dt_mov, emp_cd, cin_cd, sal_cd, fil_cd), 
        FOREIGN KEY (bol_dt_mov, emp_cd, cin_cd, sal_cd, fil_cd)
                              REFERENCES tb_boletim
 )
go
 
 
 CREATE TABLE tb_filme (
        fil_cd               int NOT NULL,
        fil_nm               varchar(50) NOT NULL,
        dis_cd               int NULL,
        fil_censura          smallint NOT NULL,
        fil_duracao          smallint NOT NULL,
        fil_durac_trai       smallint NOT NULL,
        fil_dt_ini           datetime NOT NULL,
        fil_dt_fim           datetime NOT NULL,
        fil_dt_inc           datetime NOT NULL,
        fil_dt_alt           datetime NULL,
        fil_dt_des           datetime NULL,
        fil_mot_des          varchar(50) NULL,
        fil_id_nacio         varchar(1) NOT NULL,
        PRIMARY KEY (fil_cd), 
        FOREIGN KEY (dis_cd)
                              REFERENCES tb_distribuidora
 )
go
 
 CREATE INDEX XIF90tb_filme ON tb_filme
 (
        dis_cd
 )
go
 
 
 CREATE TABLE tb_copia_filme (
        fil_cd               int NOT NULL,
        cfi_cd               int NOT NULL,
        PRIMARY KEY (fil_cd, cfi_cd), 
        FOREIGN KEY (fil_cd)
                              REFERENCES tb_filme
 )
go
 
 CREATE INDEX XIF71tb_copia_filme ON tb_copia_filme
 (
        fil_cd
 )
go
 
 
 CREATE TABLE tb_catraca (
        cat_cd               int NOT NULL,
        cin_cd               int NOT NULL,
        cat_nm               varchar(50) NULL,
        cat_dt_inc           datetime NOT NULL,
        cat_dt_alt           datetime NULL,
        cat_dt_des           datetime NULL,
        cat_mot_des          varchar(50) NULL,
        cat_hab_combo        varchar(20) NULL,
        PRIMARY KEY (cat_cd), 
        FOREIGN KEY (cin_cd)
                              REFERENCES tb_cinema
 )
go
 
 CREATE INDEX XIF60tb_catraca ON tb_catraca
 (
        cin_cd
 )
go
 
 
 CREATE TABLE tb_catraca_cont (
        cat_cd               int NOT NULL,
        ctc_dt               datetime NOT NULL,
        ctc_ini_cont         int NOT NULL,
        ctc_fim_cont         int NULL,
        PRIMARY KEY (cat_cd, ctc_dt), 
        FOREIGN KEY (cat_cd)
                              REFERENCES tb_catraca
 )
go
 
 CREATE INDEX XIF70tb_catraca_cont ON tb_catraca_cont
 (
        cat_cd
 )
go
 
 
 CREATE TABLE tb_catraca_sala (
        cat_cd               int NOT NULL,
        sal_cd               int NOT NULL,
        PRIMARY KEY (cat_cd, sal_cd), 
        FOREIGN KEY (sal_cd)
                              REFERENCES tb_sala, 
        FOREIGN KEY (cat_cd)
                              REFERENCES tb_catraca
 )
go
 
 CREATE INDEX XIF68tb_catraca_sala ON tb_catraca_sala
 (
        cat_cd
 )
go
 
 CREATE INDEX XIF69tb_catraca_sala ON tb_catraca_sala
 (
        sal_cd
 )
go
 
 
 CREATE TABLE tb_ingresso_tipo (
        igt_cd               smallint NOT NULL,
        igt_desc             varchar(30) NULL,
        PRIMARY KEY (igt_cd)
 )
go
 
 
 CREATE TABLE tb_usuario (
        usu_cd               int IDENTITY,
        cin_cd               int NULL,
        usu_nm               varchar(50) NOT NULL,
        usu_login            varchar(20) NOT NULL,
        usu_senha            char(31) NOT NULL,
        usu_dt_inc           datetime NOT NULL,
        usu_dt_alt           datetime NULL,
        usu_dt_des           datetime NULL,
        usu_mot_des          varchar(50) NULL,
        PRIMARY KEY (usu_cd), 
        FOREIGN KEY (cin_cd)
                              REFERENCES tb_cinema
 )
go
 
 CREATE UNIQUE INDEX IX_tb_usuario ON tb_usuario
 (
        usu_login,
        usu_cd
 )
go
 
 CREATE INDEX XIF19tb_usuario ON tb_usuario
 (
        cin_cd
 )
go
 
 
 CREATE TABLE tb_caixa (
        cxa_cd               int IDENTITY,
        cin_cd               int NOT NULL,
        cxa_desc             varchar(20) NOT NULL,
        cxa_dt_inc           datetime NOT NULL,
        cxa_dt_alt           datetime NULL,
        cxa_dt_des           datetime NULL,
        cxa_mot_des          varchar(50) NULL,
        PRIMARY KEY (cxa_cd), 
        FOREIGN KEY (cin_cd)
                              REFERENCES tb_cinema
 )
go
 
 CREATE INDEX XIF11tb_caixa ON tb_caixa
 (
        cin_cd
 )
go
 
 
 CREATE TABLE tb_caixa_movto (
        cxp_cd               int IDENTITY,
        cxp_dt_abertura      datetime NOT NULL,
        cxa_cd               int NOT NULL,
        usu_abertura         int NOT NULL,
        cxp_dt_fechamento    datetime NULL,
        usu_fechamento       int NULL,
        cxp_status           smallint NULL,
        cxp_talao            bit,
        PRIMARY KEY (cxp_cd), 
        FOREIGN KEY (usu_fechamento)
                              REFERENCES tb_usuario, 
        FOREIGN KEY (usu_abertura)
                              REFERENCES tb_usuario, 
        FOREIGN KEY (cxa_cd)
                              REFERENCES tb_caixa
 )
go
 
 CREATE INDEX XIF16tb_caixa_movto ON tb_caixa_movto
 (
        cxa_cd
 )
go
 
 CREATE INDEX XIF56tb_caixa_movto ON tb_caixa_movto
 (
        usu_abertura
 )
go
 
 CREATE INDEX XIF57tb_caixa_movto ON tb_caixa_movto
 (
        usu_fechamento
 )
go
 
 
 CREATE TABLE tb_operacao_tipo (
        opt_cd               int NOT NULL,
        opt_desc             varchar(30) NULL,
        opt_sinal            int NOT NULL,
        PRIMARY KEY (opt_cd)
 )
go
 
 
 CREATE TABLE tb_operacao (
        ope_cd               int IDENTITY,
        cxp_cd               int NOT NULL,
        cxa_cd               int NOT NULL,
        opt_cd               int NOT NULL,
        ope_valor            money NOT NULL,
        ope_dt_operacao      datetime NOT NULL,
        ope_dt_des           datetime NULL,
        ope_mot_des          varchar(50) NULL,
        PRIMARY KEY (ope_cd), 
        FOREIGN KEY (cxp_cd)
                              REFERENCES tb_caixa_movto, 
        FOREIGN KEY (opt_cd)
                              REFERENCES tb_operacao_tipo, 
        FOREIGN KEY (cxa_cd)
                              REFERENCES tb_caixa
 )
go
 
 CREATE INDEX XIF26tb_operacao ON tb_operacao
 (
        cxa_cd
 )
go
 
 CREATE INDEX XIF27tb_operacao ON tb_operacao
 (
        opt_cd
 )
go
 
 CREATE INDEX XIF66tb_operacao ON tb_operacao
 (
        cxp_cd
 )
go
 
 
 CREATE TABLE tb_num_talao (
        ope_cd               int NOT NULL,
        nta_seq              int NOT NULL,
        igt_cd               smallint NULL,
        nta_ini              int NULL,
        nta_fim              int NULL,
        PRIMARY KEY (ope_cd, nta_seq), 
        FOREIGN KEY (igt_cd)
                              REFERENCES tb_ingresso_tipo, 
        FOREIGN KEY (ope_cd)
                              REFERENCES tb_operacao
 )
go
 
 CREATE INDEX XIF64tb_num_talao ON tb_num_talao
 (
        ope_cd
 )
go
 
 CREATE INDEX XIF65tb_num_talao ON tb_num_talao
 (
        igt_cd
 )
go
 
 
 CREATE TABLE tb_sessao_aux (
        cxa_cd               int NOT NULL,
        sal_cd               int NOT NULL,
        fil_cd               int NOT NULL,
        sre_data             datetime NOT NULL,
        sre_horario          datetime NOT NULL,
        sea_lugares_sel      int NOT NULL,
        sea_inteiras         int NULL,
        sea_meias            int NULL,
        sea_cortesias        int NULL,
        PRIMARY KEY (cxa_cd, sal_cd, fil_cd, sre_data, sre_horario)
 )
go
 
 
CREATE TABLE [dbo].[tb_sessao_aux_p](
	[sal_cd] [int] NOT NULL,
	[fil_cd] [int] NOT NULL,
	[sre_data] [datetime] NOT NULL,
	[sre_horario] [datetime] NOT NULL,
	[sap_mat_poltr] [varchar](2704) NOT NULL,
 CONSTRAINT [PK_tb_sessao_aux_p] PRIMARY KEY CLUSTERED 
(
	[sal_cd] ASC,
	[fil_cd] ASC,
	[sre_data] ASC,
	[sre_horario] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
go

 CREATE TABLE tb_feriado (
        fer_data             datetime NOT NULL,
        fer_desc             varchar(50) NOT NULL,
        fer_dt_inc           datetime NOT NULL,
        PRIMARY KEY (fer_data)
 )
go
 
 
 CREATE TABLE tb_parametro (
        par_tmp_ses          smallint NULL,
        par_hora_max_ses     datetime NULL,
        par_hora_limite      datetime NULL,
        par_imp_cod_barra    bit,
        par_imp_lotacao      bit,
        par_imp_endereco     bit,
        par_imp_CNPJ         bit,
        par_imp_IE           bit,
        par_custo_ingresso   money NULL,
        par_imp_tck          bit,
        par_imposto_mun      money NULL,
        par_direitos_aut     money NULL,
        par_outros           money NULL,
        par_hora_limite23    datetime NULL,
        par_hora_limite34    datetime NULL,
        par_hora_limite45    datetime NULL,
        par_hora_limite56    datetime NULL,
        par_hora_limite12    datetime NULL,
        par_perc_meias       money NOT NULL,
        par_perc_cortesias   money NOT NULL,
        par_msg1             varchar(40) NULL,
        par_msg2             varchar(40) NULL,
        par_msg3             varchar(40) NULL,
        par_imp_MFIM         bit
 )
go
 
 
 CREATE TABLE tb_perfil_acesso (
        per_cd               int NOT NULL,
        per_desc             varchar(20) NOT NULL,
        PRIMARY KEY (per_cd)
 )
go
 
 
 CREATE TABLE tb_modulo (
        mod_cd               int NOT NULL,
        mod_desc             varchar(20) NOT NULL,
        PRIMARY KEY (mod_cd)
 )
go
 
 
 CREATE TABLE tb_funcao (
        mod_cd               int NOT NULL,
        fun_cd               int NOT NULL,
        fun_desc             varchar(20) NULL,
        PRIMARY KEY (mod_cd, fun_cd), 
        FOREIGN KEY (mod_cd)
                              REFERENCES tb_modulo
 )
go
 
 CREATE INDEX XIF1tb_funcao ON tb_funcao
 (
        mod_cd
 )
go
 
 
 CREATE TABLE tb_perfil_funcao (
        per_cd               int NOT NULL,
        mod_cd               int NOT NULL,
        fun_cd               int NOT NULL,
        PRIMARY KEY (per_cd, mod_cd, fun_cd), 
        FOREIGN KEY (per_cd)
                              REFERENCES tb_perfil_acesso, 
        FOREIGN KEY (mod_cd, fun_cd)
                              REFERENCES tb_funcao
 )
go
 
 CREATE INDEX XIF3tb_perfil_funcao ON tb_perfil_funcao
 (
        mod_cd,
        fun_cd
 )
go
 
 CREATE INDEX XIF4tb_perfil_funcao ON tb_perfil_funcao
 (
        per_cd
 )
go
 
 
 CREATE TABLE tb_usuario_perfil (
        usu_cd               int NOT NULL,
        per_cd               int NOT NULL,
        PRIMARY KEY (usu_cd, per_cd), 
        FOREIGN KEY (per_cd)
                              REFERENCES tb_perfil_acesso, 
        FOREIGN KEY (usu_cd)
                              REFERENCES tb_usuario
 )
go
 
 CREATE INDEX XIF6tb_usuario_perfil ON tb_usuario_perfil
 (
        usu_cd
 )
go
 
 CREATE INDEX XIF7tb_usuario_perfil ON tb_usuario_perfil
 (
        per_cd
 )
go
 
 
 CREATE TABLE tb_programacao (
        prg_cd               int IDENTITY,
        prg_dt_ini           datetime NOT NULL,
        prg_dt_fim           datetime NOT NULL,
        prg_dt_inc           datetime NOT NULL,
        prg_dt_alt           datetime NULL,
        prg_dt_des           datetime NULL,
        prg_mot_des          varchar(50) NULL,
        PRIMARY KEY (prg_cd)
 )
go
 
 
 CREATE TABLE tb_sessao (
        prg_cd               int NOT NULL,
        sal_cd               int NOT NULL,
        fil_cd               int NOT NULL,
        ses_periodo          smallint NOT NULL,
        ses_cd               smallint NOT NULL,
        ses_dia_semana       smallint NOT NULL,
        ses_horario          datetime NOT NULL,
        ses_dt_inc           datetime NOT NULL,
        ses_dt_alt           datetime NULL,
        ses_dt_des           datetime NULL,
        ses_mot_des          varchar(50) NULL,
        ses_pre_estreia      varchar(1) NOT NULL,
        PRIMARY KEY (prg_cd, sal_cd, fil_cd, ses_periodo, ses_cd, 
               ses_dia_semana), 
        FOREIGN KEY (fil_cd)
                              REFERENCES tb_filme, 
        FOREIGN KEY (sal_cd)
                              REFERENCES tb_sala, 
        FOREIGN KEY (prg_cd)
                              REFERENCES tb_programacao
 )
go
 
 CREATE INDEX XIF13tb_sessao ON tb_sessao
 (
        prg_cd
 )
go
 
 CREATE INDEX XIF14tb_sessao ON tb_sessao
 (
        sal_cd
 )
go
 
 CREATE INDEX XIF15tb_sessao ON tb_sessao
 (
        fil_cd
 )
go
 
 
 CREATE TABLE tb_prog_preco (
        ppr_cd               int IDENTITY,
        ppr_dt_ini           datetime NOT NULL,
        ppr_dt_fim           datetime NULL,
        ppr_flg_promocao     bit,
        ppr_desc             varchar(50) NOT NULL,
        ppr_patrocinador     varchar(50) NOT NULL,
        ppr_dt_inc           datetime NULL,
        ppr_dt_alt           datetime NULL,
        ppr_dt_des           datetime NULL,
        ppr_mot_des          varchar(50) NULL,
        PRIMARY KEY (ppr_cd)
 )
go
 
 
 CREATE TABLE tb_preco (
        ppr_cd               int NOT NULL,
        fil_cd               int NOT NULL,
        pre_periodo          smallint NOT NULL,
        pre_dia_semana       smallint NOT NULL,
        pre_vl_inteira_ate   money NULL,
        pre_vl_inteira_apos  money NULL,
        pre_vl_meia_ate      money NULL,
        pre_vl_meia_apos     money NULL,
        pre_dt_inc           datetime NULL,
        pre_dt_alt           datetime NULL,
        pre_dt_des           datetime NULL,
        pre_mot_des          varchar(50) NULL,
        pre_vl_inteira3      money NULL,
        pre_vl_inteira4      money NULL,
        pre_vl_inteira5      money NULL,
        pre_vl_inteira6      money NULL,
        pre_vl_meia3         money NULL,
        pre_vl_meia4         money NULL,
        pre_vl_meia5         money NULL,
        pre_vl_meia6         money NULL,
        pre_Promocao		 bit NULL,
        PRIMARY KEY (ppr_cd, fil_cd, pre_periodo, pre_dia_semana), 
        FOREIGN KEY (ppr_cd)
                              REFERENCES tb_prog_preco
 )
go
 
 CREATE INDEX XIF22tb_preco ON tb_preco
 (
        ppr_cd
 )
go
 
 
 CREATE TABLE tb_pagamento_tipo (
        pgt_cd               int NOT NULL,
        pgt_desc             varchar(20) NULL,
        PRIMARY KEY (pgt_cd)
 )
go
 
 
 CREATE TABLE tb_pagamento (
        ope_cd               int NOT NULL,
        pgt_cd               int NOT NULL,
        pag_valor            money NOT NULL,
        PRIMARY KEY (ope_cd, pgt_cd), 
        FOREIGN KEY (pgt_cd)
                              REFERENCES tb_pagamento_tipo, 
        FOREIGN KEY (ope_cd)
                              REFERENCES tb_operacao
 )
go
 
 CREATE INDEX XIF28tb_pagamento ON tb_pagamento
 (
        ope_cd
 )
go
 
 CREATE INDEX XIF29tb_pagamento ON tb_pagamento
 (
        pgt_cd
 )
go
 
 
 CREATE TABLE tb_sessao_real (
        sal_cd               int NOT NULL,
        fil_cd               int NOT NULL,
        sre_data             datetime NOT NULL,
        sre_horario          datetime NOT NULL,
        sre_lugares          smallint NOT NULL,
        sre_lugares_vendidos smallint NULL,
        sre_inteiras         int NULL,
        sre_meias            int NULL,
        sre_cortesias        int NULL,
        sre_dt_des           datetime NULL,
        sre_mot_des          varchar(50) NULL,
        usu_des              int NULL,
        sre_pre_estreia      varchar(1) NULL,
        sre_poltronas        int NULL,
        sre_mat_poltr        varchar(2704) NULL,
        PRIMARY KEY (sal_cd, fil_cd, sre_data, sre_horario), 
        FOREIGN KEY (usu_des)
                              REFERENCES tb_usuario, 
        FOREIGN KEY (fil_cd)
                              REFERENCES tb_filme, 
        FOREIGN KEY (sal_cd)
                              REFERENCES tb_sala
 )
go
 
 CREATE INDEX XIF42tb_sessao_real ON tb_sessao_real
 (
        sal_cd
 )
go
 
 CREATE INDEX XIF43tb_sessao_real ON tb_sessao_real
 (
        fil_cd
 )
go
 
 CREATE INDEX XIF51tb_sessao_real ON tb_sessao_real
 (
        usu_des
 )
go
 
 
 CREATE TABLE tb_combo (
        cbo_cd               int IDENTITY,
        cbo_nm               varchar(20) NOT NULL,
        cbo_desc             varchar(255) NOT NULL,
        cbo_dt_inc           datetime NOT NULL,
        cbo_dt_alt           datetime NULL,
        cbo_dt_des           datetime NULL,
        cbo_mot_des          varchar(50) NULL,
        PRIMARY KEY (cbo_cd)
 )
go
 
 
 CREATE TABLE tb_venda_combo (
        vcb_cd               bigint NOT NULL,
        ope_cd               int NOT NULL,
        cbo_cd               int NOT NULL,
        vcb_status           tinyint NOT NULL,
        vcb_qtde             smallint NOT NULL,
        vcb_valor            money NOT NULL,
        vcb_dt_canc          datetime NULL,
        vcb_mot_canc         varchar(50) NULL,
        PRIMARY KEY (vcb_cd), 
        FOREIGN KEY (cbo_cd)
                              REFERENCES tb_combo, 
        FOREIGN KEY (ope_cd)
                              REFERENCES tb_operacao
 )
go
 
 CREATE INDEX XIF45tb_venda_combo ON tb_venda_combo
 (
        ope_cd
 )
go
 
 CREATE INDEX XIF59tb_venda_combo ON tb_venda_combo
 (
        cbo_cd
 )
go
 
 
 CREATE TABLE tb_venda_ingresso (
        ing_cd               bigint NOT NULL,
        ope_cd               int NOT NULL,
        sal_cd               int NOT NULL,
        fil_cd               int NOT NULL,
        sre_data             datetime NOT NULL,
        sre_horario          datetime NOT NULL,
        ing_status           tinyint NOT NULL,
        ing_dt_venda         datetime NOT NULL,
        ing_dt_util          datetime NULL,
        igt_cd               smallint NOT NULL,
        ing_valor            money NOT NULL,
        ppr_cd               int NULL,
        ing_dt_canc          datetime NULL,
        ing_mot_canc         varchar(50) NULL,
        ing_num_talao        bigint NULL,
        ing_num_ing          varchar(4) NULL,
        PRIMARY KEY (ing_cd), 
        FOREIGN KEY (ppr_cd)
                              REFERENCES tb_prog_preco, 
        FOREIGN KEY (igt_cd)
                              REFERENCES tb_ingresso_tipo, 
        FOREIGN KEY (ope_cd)
                              REFERENCES tb_operacao, 
        FOREIGN KEY (sal_cd, fil_cd, sre_data, sre_horario)
                              REFERENCES tb_sessao_real
 )
go
 
 CREATE INDEX XIF50tb_venda_ingresso ON tb_venda_ingresso
 (
        sal_cd,
        fil_cd,
        sre_data,
        sre_horario
 )
go
 
 CREATE INDEX XIF52tb_venda_ingresso ON tb_venda_ingresso
 (
        ope_cd
 )
go
 
 CREATE INDEX XIF54tb_venda_ingresso ON tb_venda_ingresso
 (
        igt_cd
 )
go
 
 CREATE INDEX XIF58tb_venda_ingresso ON tb_venda_ingresso
 (
        ppr_cd
 )
go
 
 
 CREATE TABLE tb_sala_lugar (
        sal_cd               int NOT NULL,
        sal_dt_ini           datetime NOT NULL,
        sal_dt_fim           datetime NULL,
        sal_lugares          smallint NOT NULL,
        sal_mot_alt          varchar(50) NULL,
        usu_cd               int NOT NULL,
        PRIMARY KEY (sal_cd, sal_dt_ini), 
        FOREIGN KEY (usu_cd)
                              REFERENCES tb_usuario, 
        FOREIGN KEY (sal_cd)
                              REFERENCES tb_sala
 )
go
 
 CREATE INDEX XIF46tb_sala_lugar ON tb_sala_lugar
 (
        sal_cd
 )
go
 
 CREATE INDEX XIF47tb_sala_lugar ON tb_sala_lugar
 (
        usu_cd
 )
go
 
 
 CREATE TABLE tb_prog_combo (
        cbo_cd               int NOT NULL,
        pcb_dt_ini           datetime NOT NULL,
        pcb_dt_fim           datetime NULL,
        pcb_valor            money NULL,
        PRIMARY KEY (cbo_cd, pcb_dt_ini), 
        FOREIGN KEY (cbo_cd)
                              REFERENCES tb_combo
 )
go
 
 CREATE INDEX XIF53tb_prog_combo ON tb_prog_combo
 (
        cbo_cd
 )
go
 
 
 CREATE TABLE tb_bol_ope_tp (
        opt_cd               int NOT NULL,
        opt_desc             varchar(30) NULL,
        opt_sinal            int NOT NULL,
        PRIMARY KEY (opt_cd)
 )
go
 
 
 CREATE TABLE tb_bol_pag_tp (
        pgt_cd               int NOT NULL,
        pgt_desc             varchar(20) NULL,
        PRIMARY KEY (pgt_cd)
 )
go
 
 
 CREATE TABLE tb_bol_tp_ingr (
        igt_cd               smallint NOT NULL,
        igt_desc             varchar(30) NULL,
        PRIMARY KEY (igt_cd)
 )
go
 
 
 CREATE TABLE tb_bol_ingre (
        bol_dt_mov           datetime NOT NULL,
        emp_cd               int NOT NULL,
        cin_cd               int NOT NULL,
        sal_cd               int NOT NULL,
        fil_cd               int NOT NULL,
        sre_data             datetime NOT NULL,
        sre_horario          datetime NOT NULL,
        igt_cd               smallint NOT NULL,
        ing_status           tinyint NOT NULL,
        ing_dt_venda         datetime NOT NULL,
        opt_cd               int NOT NULL,
        bin_dev              char(1) NOT NULL,
        cxp_talao            bit,
        ppr_cd               int NOT NULL,
        pgt_cd               int NOT NULL,
        bin_qtde             int NULL,
        ing_valor            money NOT NULL,
        PRIMARY KEY (bol_dt_mov, emp_cd, cin_cd, sal_cd, fil_cd, 
               sre_data, sre_horario, igt_cd, ing_status, 
               ing_dt_venda, opt_cd, bin_dev, cxp_talao, ppr_cd, 
               pgt_cd), 
        FOREIGN KEY (opt_cd)
                              REFERENCES tb_bol_ope_tp, 
        FOREIGN KEY (pgt_cd)
                              REFERENCES tb_bol_pag_tp, 
        FOREIGN KEY (igt_cd)
                              REFERENCES tb_bol_tp_ingr, 
        FOREIGN KEY (bol_dt_mov, emp_cd, cin_cd, sal_cd, fil_cd)
                              REFERENCES tb_boletim
 )
go
 
 
 CREATE TABLE tb_bol_ingre2 (
        bol_dt_mov           datetime NOT NULL,
        emp_cd               int NOT NULL,
        cin_cd               int NOT NULL,
        sal_cd               int NOT NULL,
        fil_cd               int NOT NULL,
        sre_data             datetime NOT NULL,
        sre_horario          datetime NOT NULL,
        igt_cd               smallint NOT NULL,
        ing_status           tinyint NOT NULL,
        ing_dt_venda         datetime NOT NULL,
        opt_cd               int NOT NULL,
        bin_dev              char(1) NOT NULL,
        cxp_talao            bit,
        ppr_cd               int NOT NULL,
        pgt_cd               int NOT NULL,
        bin_qtde             int NOT NULL,
        ing_valor            money NOT NULL,
        PRIMARY KEY (bol_dt_mov, emp_cd, cin_cd, sal_cd, fil_cd, 
               sre_data, sre_horario, igt_cd, ing_status, 
               ing_dt_venda, opt_cd, bin_dev, cxp_talao, ppr_cd, 
               pgt_cd)
 )
go
 
 
 CREATE TABLE tb_bol_sessao (
        bol_dt_mov           datetime NOT NULL,
        emp_cd               int NOT NULL,
        cin_cd               int NOT NULL,
        sal_cd               int NOT NULL,
        fil_cd               int NOT NULL,
        sre_data             datetime NOT NULL,
        ses_horario          datetime NOT NULL,
        ses_pre_estreia      varchar(1) NOT NULL,
        PRIMARY KEY (bol_dt_mov, emp_cd, cin_cd, sal_cd, fil_cd, 
               sre_data, ses_horario), 
        FOREIGN KEY (bol_dt_mov, emp_cd, cin_cd, sal_cd, fil_cd)
                              REFERENCES tb_boletim
 )
        ON "PRIMARY"
go
 
 
 CREATE TABLE tb_bol_talao (
        bol_dt_mov           datetime NOT NULL,
        emp_cd               int NOT NULL,
        cin_cd               int NOT NULL,
        sal_cd               int NOT NULL,
        fil_cd               int NOT NULL,
        igt_cd               smallint NOT NULL,
        num_talao_ini        bigint NOT NULL,
        num_talao_fim        bigint NOT NULL,
        PRIMARY KEY (bol_dt_mov, emp_cd, cin_cd, sal_cd, fil_cd, 
               igt_cd), 
        FOREIGN KEY (igt_cd)
                              REFERENCES tb_bol_tp_ingr, 
        FOREIGN KEY (bol_dt_mov, emp_cd, cin_cd, sal_cd, fil_cd)
                              REFERENCES tb_boletim
 )
        ON "PRIMARY"
go
 
 
 CREATE TABLE tb_sis_log (
        emp_cd               int NOT NULL,
        cin_cd               int NOT NULL,
        slg_data             datetime NOT NULL,
        usu_nm               varchar(50) NOT NULL,
        slg_descricao        varchar(4000) NOT NULL,
        PRIMARY KEY (emp_cd, cin_cd, slg_data)
 )
go


-- =============================================
-- Preenchendo Dados
-- =============================================

INSERT INTO tb_parametro
           ([par_tmp_ses]
           ,[par_hora_max_ses]
           ,[par_hora_limite]
           ,[par_imp_cod_barra]
           ,[par_imp_lotacao]
           ,[par_imp_endereco]
           ,[par_imp_CNPJ]
           ,[par_imp_IE]
           ,[par_custo_ingresso]
           ,[par_imp_tck]
           ,[par_imposto_mun]
           ,[par_direitos_aut]
           ,[par_outros]
           ,[par_hora_limite23]
           ,[par_hora_limite34]
           ,[par_hora_limite45]
           ,[par_hora_limite56]
           ,[par_hora_limite12]
           ,[par_perc_meias]
           ,[par_perc_cortesias]
           ,[par_msg1]
           ,[par_msg2]
           ,[par_msg3]
           ,[par_imp_MFIM])
     VALUES
           (30,	convert(datetime,'02:00:00.000',103),convert(datetime,'01:00:00.000',103),'false','false','false','true','false',
           CONVERT(money,0),'false',CONVERT(money,0),	CONVERT(money,0),CONVERT(money,0),convert(datetime,'16:00:00.000',103),convert(datetime,'18:00:00.000',103),convert(datetime,'20:00:00.000',103),
           convert(datetime,'12:00:00.000',103),convert(datetime,'14:00:00.000',103),CONVERT(money,0),	CONVERT(money,0),'SEJAM BEM VINDOS...','','','false')
Go

Update tb_parametro Set
       [par_hora_max_ses] = convert(datetime,'1899-31-12 02:00:00',103),
       [par_hora_limite]  = convert(datetime,'1899-31-12 01:00:00',103),
       [par_hora_limite23]= convert(datetime,'1899-31-12 14:00:00',103),
       [par_hora_limite34]= convert(datetime,'1899-31-12 16:00:00',103),
       [par_hora_limite45]= convert(datetime,'1899-31-12 18:00:00',103),
       [par_hora_limite56]= convert(datetime,'1899-31-12 20:00:00',103),
       [par_hora_limite12]= convert(datetime,'1899-31-12 12:00:00',103)
Go

insert into tb_pagamento_tipo ( pgt_cd, pgt_desc ) values ( 1, 'DINHEIRO' )
GO
insert into tb_pagamento_tipo ( pgt_cd, pgt_desc ) values ( 2, 'CARTÃO DE DÉBITO' )
GO
insert into tb_pagamento_tipo ( pgt_cd, pgt_desc ) values ( 3, 'CARTÃO DE CRÉDITO' )
GO
insert into tb_pagamento_tipo ( pgt_cd, pgt_desc ) values ( 4, 'CHEQUE' )
GO

insert into tb_operacao_tipo ( opt_cd, opt_desc, opt_sinal ) values ( 1, 'VENDA', 1 )
GO
insert into tb_operacao_tipo ( opt_cd, opt_desc, opt_sinal ) values ( 2, 'DEPÓSITO FUNDO DE CAIXA', 1 )
GO
insert into tb_operacao_tipo ( opt_cd, opt_desc, opt_sinal ) values ( 3, 'DEVOLUÇÃO DE VALOR INGRESSO', -1 )
GO
insert into tb_operacao_tipo ( opt_cd, opt_desc, opt_sinal ) values ( 4, 'DEVOLUÇÃO DE VALOR COMBO', -1 )
GO
insert into tb_operacao_tipo ( opt_cd, opt_desc, opt_sinal ) values ( 5, 'SAQUE PARA DEPÓSITO', -1 )
GO
insert into tb_operacao_tipo ( opt_cd, opt_desc, opt_sinal ) values ( 6, 'SAQUE FUNDO DE CAIXA', -1 )
GO

insert into tb_ingresso_tipo ( igt_cd, igt_desc ) values ( 1, 'INTEIRA' )
GO
insert into tb_ingresso_tipo ( igt_cd, igt_desc ) values ( 2, 'MEIA' )
GO
insert into tb_ingresso_tipo ( igt_cd, igt_desc ) values ( 3, 'PRO.-INTEIRA' )
GO
insert into tb_ingresso_tipo ( igt_cd, igt_desc ) values ( 4, 'PRO.-MEIA' )
GO
insert into tb_ingresso_tipo ( igt_cd, igt_desc ) values ( 9, 'CORTESIA' )
GO

insert into tb_perfil_acesso ( per_cd, per_desc ) values ( 9, 'ADMINISTRADOR' )
GO
insert into tb_perfil_acesso ( per_cd, per_desc ) values ( 8, 'GERENTE' )
GO
insert into tb_perfil_acesso ( per_cd, per_desc ) values ( 1, 'CAIXA' )
GO

insert into tb_distribuidora ( dis_cd, dis_nm ) values ( 1, 'COLUMBIA/SONY' )
GO
insert into tb_distribuidora ( dis_cd, dis_nm ) values ( 2, 'FOX' )
GO
insert into tb_distribuidora ( dis_cd, dis_nm ) values ( 3, 'WARNER' )
GO
insert into tb_distribuidora ( dis_cd, dis_nm ) values ( 4, 'PARAMOUNT' )
GO
insert into tb_distribuidora ( dis_cd, dis_nm ) values ( 5, 'PLAY ARTE' )
GO
insert into tb_distribuidora ( dis_cd, dis_nm ) values ( 6, 'LUMIERE' )
GO
insert into tb_distribuidora ( dis_cd, dis_nm ) values ( 7, 'U.I.P.' )
GO
insert into tb_distribuidora ( dis_cd, dis_nm ) values ( 8, 'LUMIERE' )
GO
insert into tb_distribuidora ( dis_cd, dis_nm ) values ( 9, 'PARIS FILMES' )
GO
insert into tb_distribuidora ( dis_cd, dis_nm ) values ( 10, 'IMAGEM FILMES' )
GO
insert into tb_distribuidora ( dis_cd, dis_nm ) values ( 11, 'DREAMLAND' )
GO
insert into tb_distribuidora ( dis_cd, dis_nm ) values ( 12, 'UNIVERSAL' )
GO
insert into tb_distribuidora ( dis_cd, dis_nm ) values ( 13, 'ART FILMS' )
GO
insert into tb_distribuidora ( dis_cd, dis_nm ) values ( 14, 'ESTAÇÃO BOTAFOGO' )
GO
insert into tb_distribuidora ( dis_cd, dis_nm ) values ( 15, 'POLI FILMS' )
GO
insert into tb_distribuidora ( dis_cd, dis_nm ) values ( 16, 'MAIS FILMES' )
GO
insert into tb_distribuidora ( dis_cd, dis_nm ) values ( 17, 'PANDORA' )
GO
insert into tb_distribuidora ( dis_cd, dis_nm ) values ( 18, 'ALPHA FILMES' )
GO
insert into tb_distribuidora ( dis_cd, dis_nm ) values ( 19, 'DISNEY/SONY' )
GO
insert into tb_distribuidora ( dis_cd, dis_nm ) values ( 20, 'FOCCUS FILMS/IMAGEM' )
GO

INSERT INTO tb_modulo (mod_cd, mod_desc)
VALUES (1, 'Administração')
GO

INSERT INTO tb_modulo (mod_cd, mod_desc)
VALUES (2, 'Caixa')
GO

INSERT INTO tb_funcao (mod_cd, fun_cd, fun_desc)
VALUES (1, 1, 'Acesso')
GO

INSERT INTO tb_funcao (mod_cd, fun_cd, fun_desc)
VALUES (2, 1, 'Acesso')
GO

INSERT INTO tb_funcao (mod_cd, fun_cd, fun_desc)
VALUES (2, 2,  'Sangria')
GO

INSERT INTO tb_funcao (mod_cd, fun_cd, fun_desc)
VALUES (2, 3,  'Fechamento Caixa')
GO

INSERT INTO tb_funcao (mod_cd, fun_cd, fun_desc)
VALUES (2, 4,  'Libera Caixa')
GO

INSERT INTO tb_funcao (mod_cd, fun_cd, fun_desc)
VALUES (2, 5,  'Modo Talão')
GO

INSERT INTO tb_funcao (mod_cd, fun_cd, fun_desc)
VALUES (2, 6,  'Canc. Combo')
GO

INSERT INTO tb_funcao (mod_cd, fun_cd, fun_desc)
VALUES (2, 7,  'Canc. Ingresso')
GO

INSERT INTO tb_funcao (mod_cd, fun_cd, fun_desc)
VALUES (2, 8,  'Canc. Operação')
GO

INSERT INTO tb_funcao (mod_cd, fun_cd, fun_desc)
VALUES (2, 9,  'Posição Caixa')
GO

INSERT INTO tb_funcao (mod_cd, fun_cd, fun_desc)
VALUES (2, 10,  'Reimpressão')
GO

INSERT INTO tb_funcao (mod_cd, fun_cd, fun_desc)
VALUES (1, 11,  'Perfis de Acesso')
GO

INSERT INTO tb_funcao (mod_cd, fun_cd, fun_desc)
VALUES (2, 11,  'Perfis de Acesso')
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (1, 2, 1)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (9, 1, 1)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (9, 2, 1)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (9, 2, 2)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (9, 2, 3)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (9, 2, 4)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (9, 2, 5)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (9, 2, 6)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (9, 2, 7)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (9, 2, 8)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (9, 2, 9)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (9, 2, 10)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (9, 1, 11)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (8, 2, 1)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (8, 2, 2)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (8, 2, 3)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (8, 2, 4)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (8, 2, 5)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (8, 2, 6)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (8, 2, 7)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (8, 2, 8)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (8, 2, 9)
GO

INSERT INTO tb_perfil_funcao (per_cd, mod_cd, fun_cd)
VALUES (8, 2, 10)
GO

INSERT INTO tb_usuario ( usu_nm, usu_login, usu_senha, usu_dt_inc) VALUES ( 'ADMIN', 'ADMIN', '12345678', GETDATE())
GO

DECLARE @usu_cd int
select @usu_cd = @@IDENTITY

INSERT INTO TB_USUARIO_PERFIL ( usu_cd, per_cd ) VALUES ( @usu_cd, 9 )
GO

INSERT INTO tbUltimoAcesso
VALUES ('31121950000000')
GO

-- =============================================
-- CRIANDO PROCEDURES
-- =============================================

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upACESSO_CARREGA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upACESSO_CARREGA
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upACESSO_DELETA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upACESSO_DELETA
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upACESSO_INSERE') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upACESSO_INSERE
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upACESSO_VERIFICA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upACESSO_VERIFICA
GO

--****************************************************
CREATE PROCEDURE upACESSO_CARREGA
	(@Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   SELECT mod_cd, per_cd, fun_cd
   FROM TB_PERFIL_FUNCAO
   ORDER BY mod_cd, per_cd, fun_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         	
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upACESSO_DELETA
	(@Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   DELETE FROM TB_PERFIL_FUNCAO
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         	
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upACESSO_INSERE
	(@mod_cd        int,
	 @per_cd        int,
	 @fun_cd        int,
	 @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   INSERT INTO TB_PERFIL_FUNCAO
          (mod_cd, per_cd, fun_cd)
   VALUES (@mod_cd, @per_cd, @fun_cd)
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         	
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upACESSO_VERIFICA
	(@mod_cd        int,
	 @per_cd        int,
	 @fun_cd        int,
	 @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   SELECT mod_cd, per_cd, fun_cd
   FROM TB_PERFIL_FUNCAO
   WHERE mod_cd = @mod_cd
   AND   per_cd = @per_cd
   AND   fun_cd = @fun_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         	
         RETURN
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upExisteMovto') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upExisteMovto
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upExcluiMovto') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upExcluiMovto
GO

--****************************************************
CREATE PROCEDURE upExisteMovto
	(@Data		datetime,
	 @emp_cd 	int,
	 @cin_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
    DECLARE @qtdeBol int

    SELECT @Erro   = 0
    SELECT @MsgErr = ''

    SELECT  @qtdeBol = COUNT(*)
    FROM tb_boletim
    WHERE emp_cd     = @emp_cd
    AND   cin_cd     = @cin_cd
    AND   bol_dt_mov = @Data

    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          RETURN
       END
       
    IF @qtdeBol = 0
       BEGIN
          SELECT @Erro   = 1
          SELECT @MsgErr = 'Movimento não existe'
         
          RETURN
       END
    
GO

--****************************************************
CREATE PROCEDURE upExcluiMovto
	(@Data		datetime,
	 @emp_cd 	int,
	 @cin_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS

    DELETE FROM tb_bol_ingre
    WHERE bol_dt_mov = @Data
    AND   emp_cd     = @emp_cd
    AND   cin_cd     = @cin_cd

    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          RETURN
       END
  
    DELETE FROM tb_bol_talao
    WHERE bol_dt_mov = @Data
    AND   emp_cd     = @emp_cd
    AND   cin_cd     = @cin_cd

    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          RETURN
       END

    DELETE FROM tb_bol_sessao
    WHERE bol_dt_mov = @Data
    AND   emp_cd     = @emp_cd
    AND   cin_cd     = @cin_cd

    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          RETURN
       END

    DELETE FROM tb_boletim
    WHERE bol_dt_mov = @Data
    AND   emp_cd     = @emp_cd
    AND   cin_cd     = @cin_cd

    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          RETURN
       END

    DELETE FROM tb_bol_param
    WHERE bol_dt_mov = @Data
    AND   emp_cd     = @emp_cd
    AND   cin_cd     = @cin_cd

    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          RETURN
       END

    DELETE FROM tb_bol_filme
    WHERE bol_dt_mov = @Data
    AND   emp_cd     = @emp_cd
    AND   cin_cd     = @cin_cd

    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          RETURN
       END

    DELETE FROM tb_bol_distrib
    WHERE bol_dt_mov = @Data

    --SELECT @Erro = @@ERROR
    
    --IF @Erro <> 0
    --   BEGIN
    --      SELECT @MsgErr = description
    --      FROM master..sysmessages
    --      WHERE error = @Erro
    --     
    --      RETURN
    --   END

    DELETE FROM tb_bol_catraca_sala
    WHERE bol_dt_mov = @Data
    AND   emp_cd     = @emp_cd
    AND   cin_cd     = @cin_cd

    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          RETURN
       END

    DELETE FROM tb_bol_catraca
    WHERE bol_dt_mov = @Data
    AND   emp_cd     = @emp_cd
    AND   cin_cd     = @cin_cd

    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          RETURN
       END

    DELETE FROM tb_bol_sala
    WHERE bol_dt_mov = @Data
    AND   emp_cd     = @emp_cd
    AND   cin_cd     = @cin_cd

    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          RETURN
       END

    DELETE FROM tb_bol_cin
    WHERE bol_dt_mov = @Data
    AND   emp_cd     = @emp_cd
    AND   cin_cd     = @cin_cd

    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          RETURN
       END

    DELETE FROM tb_bol_empr
    WHERE bol_dt_mov = @Data
    AND   emp_cd     = @emp_cd

    --SELECT @Erro = @@ERROR
    
    --IF @Erro <> 0
    --   BEGIN
    --      SELECT @MsgErr = description
    --      FROM master..sysmessages
    --      WHERE error = @Erro
    --     
    --      RETURN
    --   END

    DELETE FROM tb_sis_log
    WHERE convert(datetime, convert(char(10), slg_data, 103), 103) = @Data
    AND   emp_cd  = @emp_cd
    AND   cin_cd  = @cin_cd  

    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          RETURN
       END

GO



SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_CAPA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_CAPA
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_VENDAS_DIA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_VENDAS_DIA
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_PRE_VENDA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_PRE_VENDA
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_CORTESIA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_CORTESIA
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_DEVOLUCAO') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_DEVOLUCAO
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_SESSOES_FILME') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_SESSOES_FILME
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_TOTAL_SESSAO') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_TOTAL_SESSAO
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_VENDA_ANTECIPADA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_VENDA_ANTECIPADA
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_FORMA_PAGTO') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_FORMA_PAGTO
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_NUMERACAO_TALAO') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_NUMERACAO_TALAO
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_CATRACA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_CATRACA
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_INGRESSO_S_USO') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_INGRESSO_S_USO
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_VENDAS_TOTAL') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_VENDAS_TOTAL
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_VENDA_COMBO') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_VENDA_COMBO
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_VENDA_COMBO_TOTAL') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_VENDA_COMBO_TOTAL
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_VENDA_INGRESSO') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_VENDA_INGRESSO
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_VENDA_INGRESSO_TOTAL') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_VENDA_INGRESSO_TOTAL
GO

--****************************************************
/*
declare @erro int, @msgerr varchar(255), @data datetime
select @data = CONVERT(datetime, '12/10/2005', 103 )
exec upBOLETIM_CAPA @data, 2, @erro, @msgerr
*/
--****************************************************
CREATE PROCEDURE upBOLETIM_CAPA
	(@Data		datetime,
	 @sal_cd 	int,
	 @fil_cd 	int,
	 @ses_excl      char(1),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
--	SET DATEFIRST 1
	DECLARE @diaSemana	smallint, 
		@dt_abertura 	datetime
		
	SELECT @diaSemana = 8
	  FROM tb_feriado
	 WHERE fer_data = CONVERT(datetime, @Data, 103)
    
	IF @diaSemana is null
	    SELECT @diaSemana = datepart(dw,CONVERT(datetime, @Data, 103))
	    
	SELECT @dt_abertura = MAX(cxp_dt_abertura) 
	  FROM TB_CAIXA_MOVTO
	 WHERE CONVERT(datetime, CONVERT(char(10),cxp_dt_abertura,102),102) = CONVERT(datetime, CONVERT(char(10),@Data,102), 102)

	IF @ses_excl = 'N'
	   SELECT DISTINCT f.emp_nm, e.cin_nm, d.sal_desc, 
	          d.sal_lugares, @dt_abertura 'dt_abertura', getdate() 'dt_atual',
	          c.fil_nm, g.dis_nm 'fil_distribuidora', 
	          a.prg_dt_ini, a.prg_dt_fim, b.fil_cd, b.ses_pre_estreia 'pre_estreia'
	     FROM TB_PROGRAMACAO a,
	          TB_SESSAO b,
	          TB_FILME c,
	          TB_SALA d,
	          TB_CINEMA e,
	          TB_EMPRESA f,
	          TB_DISTRIBUIDORA g
	    WHERE a.prg_cd = b.prg_cd
	      AND b.fil_cd = c.fil_cd
	      AND b.sal_cd = d.sal_cd
	      AND d.cin_cd = e.cin_cd
	      AND e.emp_cd = f.emp_cd
	      AND c.dis_cd = g.dis_cd
	      AND @Data between a.prg_dt_ini AND a.prg_dt_fim
	      AND a.prg_dt_des IS NULL
	      AND b.sal_cd = @sal_cd
	      AND b.fil_cd = @fil_cd
	      AND b.ses_dia_semana = @diaSemana
	ELSE
	   SELECT DISTINCT f.emp_nm, e.cin_nm, d.sal_desc, 
	          d.sal_lugares, @dt_abertura 'dt_abertura', getdate() 'dt_atual',
	          c.fil_nm, g.dis_nm 'fil_distribuidora', 
	          a.prg_dt_ini, a.prg_dt_fim, b.fil_cd, b.sre_pre_estreia 'pre_estreia'
	     from tb_programacao a,
	          tb_sessao_real b,
	          tb_filme c,
	          tb_sala d,
	          tb_cinema e,
	          tb_empresa f,
	          TB_DISTRIBUIDORA g
	    WHERE b.sre_data = @Data
	      AND b.fil_cd   = c.fil_cd
	      AND b.sal_cd   = d.sal_cd
	      AND d.cin_cd   = e.cin_cd
	      AND e.emp_cd   = f.emp_cd
	      AND c.dis_cd   = g.dis_cd
	      AND @Data between a.prg_dt_ini AND a.prg_dt_fim
	      AND a.prg_dt_des IS NULL
	      AND b.sal_cd = @sal_cd
	      AND b.fil_cd = @fil_cd
	

    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          Return
       End

GO

--****************************************************
/*
declare @erro int, @msgerr varchar(255), @data datetime
select @data = getdate() --CONVERT(datetime, '25/12/2005', 103 )
exec upBOLETIM_CORTESIA @data, 1, 7, @erro, @msgerr
*/
--****************************************************
CREATE PROCEDURE upBOLETIM_CORTESIA
	(@Data		datetime,
	 @sal_cd 	int,
	 @fil_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
	SELECT y.igt_cd, ISNULL(x.qtde,0) 'qtde', y.igt_desc , x.desc_dia
	FROM (
		SELECT b.igt_cd, count(1) 'qtde', 
	               ISNULL(SUM(b.ing_valor),0) 'ing_valor',
	               ' ' 'desc_dia'
		  FROM TB_OPERACAO a,
		       TB_VENDA_INGRESSO b,
		       TB_CAIXA_MOVTO d
		 WHERE b.ope_cd = a.ope_cd
		   AND d.cxp_cd = a.cxp_cd
		   AND CONVERT(datetime, CONVERT(char(10),a.ope_dt_operacao,103), 103) >= @Data
		   AND b.sre_data = @Data
		   AND b.sal_cd = @sal_cd
		   AND b.fil_cd = @fil_cd
		   AND a.opt_cd IN ( 1, 3 )		-- Operação Venda Normal + devolução
		   AND b.igt_cd = 9 			-- Cortesia
		   AND a.ope_dt_des IS NULL
		GROUP BY b.igt_cd
		UNION
		SELECT b.igt_cd, count(1) 'qtde', 
	               ISNULL(SUM(b.ing_valor),0) 'ing_valor',
	               ' PRE' 'desc_dia'
		  FROM TB_OPERACAO a,
		       TB_VENDA_INGRESSO b, 
		       TB_CAIXA_MOVTO d
		 WHERE b.ope_cd = a.ope_cd
		   AND d.cxp_cd = a.cxp_cd
		   AND CONVERT(datetime, CONVERT(char(10),a.ope_dt_operacao,103), 103) < @Data
		   AND b.sre_data = @Data               -- Pré-Venda
		   AND b.sal_cd = @sal_cd
		   AND b.fil_cd = @fil_cd
		   AND a.opt_cd IN ( 1, 3 )		-- Operação Venda Normal + devolução
		   AND b.igt_cd = 9 			-- Cortesia
		   AND a.ope_dt_des IS NULL
		GROUP BY b.igt_cd
		
		) x,
		TB_INGRESSO_TIPO y
	 WHERE x.igt_cd = y.igt_cd
	   AND y.igt_cd = 9 		-- Cortesia
	 ORDER BY 1
	 
    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          Return
       End

GO

--****************************************************
/*
declare @erro int, @msgerr varchar(255), @data datetime
select @data = getdate() --CONVERT(datetime, '18/09/2005', 103 )
exec upBOLETIM_DEVOLUCAO @data, 2, 7, @erro, @msgerr
*/
--****************************************************
CREATE PROCEDURE upBOLETIM_DEVOLUCAO
	(@Data		datetime,
	 @sal_cd 	int,
	 @fil_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
	  
	SELECT 0 'ordem', b.igt_cd, count(1) 'qtde', ISNULL(SUM(b.ing_valor),0) 'ing_valor',
	       c.igt_desc 'igt_desc', b.ppr_cd, b.ing_valor 'ing_valorA'
        FROM TB_OPERACAO a,
             TB_VENDA_INGRESSO b,
             TB_INGRESSO_TIPO c,
             TB_CAIXA_MOVTO d
        WHERE b.ope_cd   = a.ope_cd
        AND   d.cxp_cd   = a.cxp_cd
        AND   c.igt_cd   = b.igt_cd
        AND   b.sal_cd   = @sal_cd
        AND   b.fil_cd   = @fil_cd
        AND   b.sre_data = @Data
        AND   a.opt_cd = 1			-- Operação Venda Normal + devolução
        AND   b.igt_cd IN ( 1, 2 ) 		-- Inteira e Meia
        AND   d.cxp_talao = 0			-- não talão
        AND   b.ing_dt_canc > @Data
        GROUP BY b.igt_cd, c.igt_desc, b.ppr_cd, b.ing_valor
	UNION
	SELECT 2 'ordem', b.igt_cd, count(1) 'qtde', ISNULL(SUM(b.ing_valor),0) 'ing_valor',
	       c.igt_desc 'igt_desc', b.ppr_cd, b.ing_valor 'ing_valorA'
        FROM TB_OPERACAO a,
             TB_VENDA_INGRESSO b,
             TB_INGRESSO_TIPO c,
             TB_CAIXA_MOVTO d
        WHERE b.ope_cd   = a.ope_cd
        AND   d.cxp_cd   = a.cxp_cd
        AND   c.igt_cd   = b.igt_cd
        AND   b.sal_cd   = @sal_cd
        AND   b.fil_cd   = @fil_cd
        AND   b.sre_data = @Data
        AND   a.opt_cd = 1			-- Operação Venda Normal + devolução
        AND   b.igt_cd IN ( 3, 4 ) 		-- Inteira e Meia Promoção
        AND   d.cxp_talao = 0			-- não talão
        AND   b.ing_dt_canc > @Data
        GROUP BY b.igt_cd, c.igt_desc, b.ppr_cd, b.ing_valor
	UNION
	SELECT 4 'ordem', b.igt_cd, count(1) 'qtde', ISNULL(SUM(b.ing_valor),0) 'ing_valor',
	       c.igt_desc + ' TA' 'igt_desc', b.ppr_cd, b.ing_valor 'ing_valorA'
        FROM TB_OPERACAO a,
             TB_VENDA_INGRESSO b,
             TB_INGRESSO_TIPO c,
             TB_CAIXA_MOVTO d
        WHERE b.ope_cd   = a.ope_cd
        AND   d.cxp_cd   = a.cxp_cd
        AND   c.igt_cd   = b.igt_cd
        AND   b.sal_cd   = @sal_cd
        AND   b.fil_cd   = @fil_cd
        AND   b.sre_data = @Data
        AND   a.opt_cd = 1			-- Operação Venda Normal + devolução
        AND   b.igt_cd IN ( 1, 2 ) 		-- Inteira e Meia
        AND   d.cxp_talao = 1			-- Talão
        AND   b.ing_dt_canc > @Data
        GROUP BY b.igt_cd, c.igt_desc, b.ppr_cd, b.ing_valor
	UNION
	SELECT 6 'ordem', b.igt_cd, count(1) 'qtde', ISNULL(SUM(b.ing_valor),0) 'ing_valor',
	       c.igt_desc + ' TA' 'igt_desc', b.ppr_cd, b.ing_valor 'ing_valorA'
        FROM TB_OPERACAO a,
             TB_VENDA_INGRESSO b,
             TB_INGRESSO_TIPO c,
             TB_CAIXA_MOVTO d
        WHERE b.ope_cd   = a.ope_cd
        AND   d.cxp_cd   = a.cxp_cd
        AND   c.igt_cd   = b.igt_cd
        AND   b.sal_cd   = @sal_cd
        AND   b.fil_cd   = @fil_cd
        AND   b.sre_data = @Data
        AND   a.opt_cd = 1			-- Operação Venda Normal + devolução
        AND   b.igt_cd IN ( 3, 4 ) 		-- Inteira e Meia Promoção
        AND   d.cxp_talao = 1			-- Talão
        AND   b.ing_dt_canc > @Data
        GROUP BY b.igt_cd, c.igt_desc, b.ppr_cd, b.ing_valor
	--UNION
	--SELECT 8 'ordem', b.igt_cd, count(1) 'qtde', ISNULL(SUM(b.ing_valor),0) 'ing_valor',
	--       c.igt_desc 'igt_desc', b.ppr_cd, b.ing_valor 'ing_valorA'
        --FROM TB_OPERACAO a,
        --     TB_VENDA_INGRESSO b,
        --     TB_INGRESSO_TIPO c,
        --     TB_CAIXA_MOVTO d
        --WHERE b.ope_cd   = a.ope_cd
        --AND   d.cxp_cd   = a.cxp_cd
        --AND   c.igt_cd   = b.igt_cd
        --AND   b.sal_cd   = @sal_cd
        --AND   b.fil_cd   = @fil_cd
        --AND   b.sre_data = @Data
        --AND   a.opt_cd = 1			-- Operação Venda Normal + devolução
        --AND   b.igt_cd IN ( 9 ) 		-- Cortesia
        --AND   d.cxp_talao = 0			-- não talão
        --AND   b.ing_dt_canc > @Data
        --GROUP BY b.igt_cd, c.igt_desc, b.ppr_cd, b.ing_valor

        ORDER BY 1
    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          Return
       End

GO

--****************************************************
/*
declare @erro int, @msgerr varchar(255), @data datetime
select @data = CONVERT(datetime, '7/10/2005', 103 )
exec upBOLETIM_FORMA_PAGTO @data, 2, 7, @erro, @msgerr
*/
--****************************************************
CREATE PROCEDURE upBOLETIM_FORMA_PAGTO
	(@Data		datetime,
	 @sal_cd 	int,
	 @fil_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
	DECLARE @HoraMaxSes     datetime,
	        @DataIniPer     datetime,
	        @DataFimPer     datetime

	SELECT @HoraMaxSes = par_hora_max_ses
	  FROM tb_parametro
	  
        SELECT @DataIniPer = CONVERT(datetime, CONVERT(char(10),@Data,103) + ' ' + CONVERT(char(10),@HoraMaxSes,108), 103)
        SELECT @DataFimPer = CONVERT(datetime, CONVERT(char(10),DATEADD(Day, 1, @Data),103) + ' ' + CONVERT(char(10),@HoraMaxSes,108), 103)

	SELECT y.pgt_cd, y.pgt_desc, CASE WHEN x.valor IS NULL THEN 0
                                          ELSE x.valor
                                     END valor
	  FROM (
		SELECT c.pgt_cd, 
		       SUM(b.ing_valor) valor
		  FROM TB_OPERACAO a,
		       TB_VENDA_INGRESSO b,
		       TB_PAGAMENTO c
		 WHERE a.ope_cd = b.ope_cd
		   AND a.ope_cd = c.ope_cd
                   AND a.ope_dt_operacao between @DataIniPer AND @DataFimPer
		   AND b.sal_cd = @sal_cd
		   AND b.fil_cd = @fil_cd
		   AND b.ing_dt_canc IS NULL
	      GROUP BY c.pgt_cd
		) x,
		TB_PAGAMENTO_TIPO y
	 WHERE y.pgt_cd = x.pgt_cd
	 ORDER BY y.pgt_cd
    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          Return
       End

GO

--****************************************************
/*
declare @erro int, @msgerr varchar(255), @data datetime
select @data = CONVERT(datetime, '7/10/2005', 103 )
exec upBOLETIM_NUMERACAO_TALAO @data, 2, 7, @erro, @msgerr
*/
--****************************************************
CREATE PROCEDURE upBOLETIM_NUMERACAO_TALAO
	(@Data		datetime,
	 @sal_cd 	int,
	 @fil_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
	DECLARE @HoraMaxSes     datetime,
	        @DataIniPer     datetime,
	        @DataFimPer     datetime

	SELECT @HoraMaxSes = par_hora_max_ses
	  FROM tb_parametro
	  
        SELECT @DataIniPer = CONVERT(datetime, CONVERT(char(10),@Data,103) + ' ' + CONVERT(char(10),@HoraMaxSes,108), 103)
        SELECT @DataFimPer = CONVERT(datetime, CONVERT(char(10),DATEADD(Day, 1, @Data),103) + ' ' + CONVERT(char(10),@HoraMaxSes,108), 103)
	  
	SELECT a.ope_cd, c.igt_desc, 
	       MIN(b.ing_num_talao) 'numIni', MAX(b.ing_num_talao) 'numFim'
	FROM TB_OPERACAO a,
	     TB_VENDA_INGRESSO b,
	     TB_INGRESSO_TIPO c
	WHERE a.ope_cd = b.ope_cd
	AND   b.igt_cd = c.igt_cd
	AND   a.ope_dt_operacao between @DataIniPer AND @DataFimPer
	AND b.sal_cd = @sal_cd
	AND b.fil_cd = @fil_cd
        AND a.ope_dt_des IS NULL
        AND b.ing_dt_canc IS NULL
        AND b.ing_num_talao IS NOT NULL
        GROUP BY  a.ope_cd, c.igt_desc
	ORDER BY 3
    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          Return
       End

GO

--****************************************************
/*
declare @erro int, @msgerr varchar(255), @data datetime
select @data = getdate() --CONVERT(datetime, '10/12/2005', 103 )
exec upBOLETIM_VENDAS_TOTAL @data, 2, 7, @erro, @msgerr
*/
--****************************************************
CREATE PROCEDURE upBOLETIM_VENDAS_TOTAL
	(@Data		datetime,
	 @sal_cd 	int,
	 @fil_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
	DECLARE @HoraLimite 	char(5),
	        @HoraMaxSes     datetime,
	        @DataIniPer     datetime,
	        @DataFimPer     datetime
	        
	SELECT @HoraMaxSes = par_hora_max_ses
	  FROM tb_parametro
	  
	SELECT 0 'ordem', b.igt_cd, count(1) 'qtde', ISNULL(SUM(b.ing_valor),0) 'ing_valor',
	       c.igt_desc 'igt_desc', b.ppr_cd, b.ing_valor 'ing_valorA'
        FROM TB_OPERACAO a,
             TB_VENDA_INGRESSO b,
             TB_INGRESSO_TIPO c,
             TB_CAIXA_MOVTO d
        WHERE d.cxp_cd   = a.cxp_cd
        AND   a.ope_cd   = b.ope_cd
        AND   c.igt_cd   = b.igt_cd
        AND   b.sal_cd   = @sal_cd
        AND   b.fil_cd   = @fil_cd
        AND   b.sre_data = @Data
        AND a.opt_cd IN ( 1, 3 )		-- Operação Venda Normal + devolução
        AND b.igt_cd IN ( 1, 2 ) 		-- Inteira e Meia
        AND d.cxp_talao = 0			-- não talão
        GROUP BY b.igt_cd, c.igt_desc, b.ppr_cd, b.ing_valor
	UNION
	SELECT 2 'ordem', b.igt_cd, count(1) 'qtde', ISNULL(SUM(b.ing_valor),0) 'ing_valor',
	       c.igt_desc 'igt_desc', b.ppr_cd, b.ing_valor 'ing_valorA'
        FROM TB_OPERACAO a,
             TB_VENDA_INGRESSO b,
             TB_INGRESSO_TIPO c,
             TB_CAIXA_MOVTO d
        WHERE d.cxp_cd   = a.cxp_cd
        AND   a.ope_cd   = b.ope_cd
        AND   c.igt_cd   = b.igt_cd
        AND   b.sal_cd   = @sal_cd
        AND   b.fil_cd   = @fil_cd
        AND   b.sre_data = @Data
        AND a.opt_cd IN ( 1, 3 )		-- Operação Venda Normal + devolução
        AND b.igt_cd IN ( 3, 4 ) 		-- Inteira e Meia Promoção
        AND d.cxp_talao = 0			-- não talão
        GROUP BY b.igt_cd, c.igt_desc, b.ppr_cd, b.ing_valor
	UNION
	SELECT 4 'ordem', b.igt_cd, count(1) 'qtde', ISNULL(SUM(b.ing_valor),0) 'ing_valor',
	       c.igt_desc + ' TA' 'igt_desc', b.ppr_cd, b.ing_valor 'ing_valorA'
        FROM TB_OPERACAO a,
             TB_VENDA_INGRESSO b,
             TB_INGRESSO_TIPO c,
             TB_CAIXA_MOVTO d
        WHERE d.cxp_cd   = a.cxp_cd
        AND   a.ope_cd   = b.ope_cd
        AND   c.igt_cd   = b.igt_cd
        AND   b.sal_cd   = @sal_cd
        AND   b.fil_cd   = @fil_cd
        AND   b.sre_data = @Data
        AND a.opt_cd IN ( 1, 3 )		-- Operação Venda Normal + devolução
        AND b.igt_cd IN ( 1, 2 ) 		-- Inteira e Meia
        AND d.cxp_talao = 1			-- Talão
        GROUP BY b.igt_cd, c.igt_desc, b.ppr_cd, b.ing_valor
	UNION
	SELECT 6 'ordem', b.igt_cd, count(1) 'qtde', ISNULL(SUM(b.ing_valor),0) 'ing_valor',
	       c.igt_desc + ' TA' 'igt_desc', b.ppr_cd, b.ing_valor 'ing_valorA'
        FROM TB_OPERACAO a,
             TB_VENDA_INGRESSO b,
             TB_INGRESSO_TIPO c,
             TB_CAIXA_MOVTO d
        WHERE d.cxp_cd   = a.cxp_cd
        AND   a.ope_cd   = b.ope_cd
        AND   c.igt_cd   = b.igt_cd
        AND   b.sal_cd   = @sal_cd
        AND   b.fil_cd   = @fil_cd
        AND   b.sre_data = @Data
        AND a.opt_cd IN ( 1, 3 )		-- Operação Venda Normal + devolução
        AND b.igt_cd IN ( 3, 4 ) 		-- Inteira e Meia Promoção
        AND d.cxp_talao = 1			-- Talão
        GROUP BY b.igt_cd, c.igt_desc, b.ppr_cd, b.ing_valor
        ORDER BY 1
        
    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          Return
       End

GO

--****************************************************
/*
declare @erro int, @msgerr varchar(255), @data datetime
select @data = getdate() --CONVERT(datetime, '25/12/2005', 103 )
exec upBOLETIM_PRE_VENDA @data, 1, 7, @erro, @msgerr
*/
--****************************************************
CREATE PROCEDURE upBOLETIM_PRE_VENDA
	(@Data		datetime,
	 @sal_cd 	int,
	 @fil_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
	SELECT 0 'ordem', b.igt_cd, count(1) 'qtde', ISNULL(SUM(b.ing_valor),0) 'ing_valor',
	       c.igt_desc 'igt_desc', b.ppr_cd, b.ing_valor 'ing_valorA'
        FROM TB_OPERACAO a,
             TB_VENDA_INGRESSO b,
             TB_INGRESSO_TIPO c,
             TB_CAIXA_MOVTO d
        WHERE b.ope_cd = a.ope_cd
        AND   d.cxp_cd = a.cxp_cd
        AND   c.igt_cd = b.igt_cd
        AND   CONVERT(datetime, CONVERT(char(10),a.ope_dt_operacao,103),103) < @Data
        AND   b.sre_data = @Data                                              -- Pré-Venda
        AND   b.sal_cd = @sal_cd
        AND   b.fil_cd = @fil_cd
        AND   a.opt_cd IN ( 1, 3 )		-- Operação Venda Normal + devolução
        AND   b.igt_cd IN ( 1, 2 ) 		-- Inteira e Meia
        AND   d.cxp_talao = 0			-- não talão
        AND   (b.ing_dt_canc IS NULL
        OR     b.ing_dt_canc > @Data)
	GROUP BY b.igt_cd, c.igt_desc, b.ppr_cd, b.ing_valor
	UNION
	SELECT 2 'ordem', b.igt_cd, count(1) 'qtde', ISNULL(SUM(b.ing_valor),0) 'ing_valor',
	       c.igt_desc 'igt_desc', b.ppr_cd, b.ing_valor 'ing_valorA'
        FROM TB_OPERACAO a,
             TB_VENDA_INGRESSO b,
             TB_INGRESSO_TIPO c,
             TB_CAIXA_MOVTO d
        WHERE b.ope_cd = a.ope_cd
        AND   d.cxp_cd = a.cxp_cd
        AND   c.igt_cd = b.igt_cd
        AND   CONVERT(datetime, CONVERT(char(10),a.ope_dt_operacao,103),103) < @Data
        AND   b.sre_data = @Data                                              -- Pré-Venda
        AND   b.sal_cd = @sal_cd
        AND   b.fil_cd = @fil_cd
        AND   a.opt_cd IN ( 1, 3 )		-- Operação Venda Normal + devolução
        AND   b.igt_cd IN ( 3, 4 ) 		-- Inteira e Meia Promoção
        AND   d.cxp_talao = 0			-- não talão
        AND   (b.ing_dt_canc IS NULL
        OR     b.ing_dt_canc > @Data)
	GROUP BY b.igt_cd, c.igt_desc, b.ppr_cd, b.ing_valor
	UNION
	SELECT 4 'ordem', b.igt_cd, count(1) 'qtde', ISNULL(SUM(b.ing_valor),0) 'ing_valor',
	       c.igt_desc + ' TA' 'igt_desc', b.ppr_cd, b.ing_valor 'ing_valorA'
        FROM TB_OPERACAO a,
             TB_VENDA_INGRESSO b,
             TB_INGRESSO_TIPO c,
             TB_CAIXA_MOVTO d
        WHERE b.ope_cd = a.ope_cd
        AND   d.cxp_cd = a.cxp_cd
        AND   c.igt_cd = b.igt_cd
        AND   CONVERT(datetime, CONVERT(char(10),a.ope_dt_operacao,103),103) < @Data
        AND   b.sre_data = @Data                                              -- Pré-Venda
        AND   b.sal_cd = @sal_cd
        AND   b.fil_cd = @fil_cd
        AND   a.opt_cd IN ( 1, 3 )		-- Operação Venda Normal + devolução
        AND   b.igt_cd IN ( 1, 2 ) 		-- Inteira e Meia
        AND   d.cxp_talao = 1			-- Talão
        AND   (b.ing_dt_canc IS NULL
        OR     b.ing_dt_canc > @Data)
	GROUP BY b.igt_cd, c.igt_desc, b.ppr_cd, b.ing_valor
	UNION
	SELECT 6 'ordem', b.igt_cd, count(1) 'qtde', ISNULL(SUM(b.ing_valor),0) 'ing_valor',
	       c.igt_desc + ' TA' 'igt_desc', b.ppr_cd, b.ing_valor 'ing_valorA'
        FROM TB_OPERACAO a,
             TB_VENDA_INGRESSO b,
             TB_INGRESSO_TIPO c,
             TB_CAIXA_MOVTO d
        WHERE b.ope_cd = a.ope_cd
        AND   d.cxp_cd = a.cxp_cd
        AND   c.igt_cd = b.igt_cd
        AND   CONVERT(datetime, CONVERT(char(10),a.ope_dt_operacao,103),103) < @Data
        AND   b.sre_data = @Data                                              -- Pré-Venda
        AND   b.sal_cd = @sal_cd
        AND   b.fil_cd = @fil_cd
        AND   a.opt_cd IN ( 1, 3 )		-- Operação Venda Normal + devolução
        AND   b.igt_cd IN ( 3, 4 ) 		-- Inteira e Meia Promoção
        AND   d.cxp_talao = 1			-- Talão
        AND   (b.ing_dt_canc IS NULL
        OR     b.ing_dt_canc > @Data)
	GROUP BY b.igt_cd, c.igt_desc, b.ppr_cd, b.ing_valor
	ORDER BY 1
    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          Return
       End

GO

--****************************************************
/*
declare @erro int, @msgerr varchar(255), @data datetime
select @data = getdate() --CONVERT(datetime, '12/10/2005', 103 )
exec upBOLETIM_SESSOES_FILME @data, 2, 7, @erro, @msgerr
*/
--****************************************************
CREATE PROCEDURE upBOLETIM_SESSOES_FILME
	(@Data		datetime,
	 @sal_cd 	int,
	 @fil_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
--    SET DATEFIRST 1
    
    DECLARE @diaSemana	smallint
    
    SELECT @diaSemana = 8
      FROM tb_feriado
     WHERE fer_data = CONVERT(datetime, @Data, 103)
     
    IF @diaSemana is null
        SELECT @diaSemana = datepart(dw,CONVERT(datetime, @Data, 103))
        
    SELECT distinct a.ses_horario
      FROM TB_SESSAO a,
           TB_PROGRAMACAO b,
           TB_SALA c,
           TB_FILME d
     WHERE a.prg_cd = b.prg_cd
       AND a.sal_cd = c.sal_cd
       AND a.fil_cd = d.fil_cd
       AND a.sal_cd = @sal_cd
       AND a.fil_cd = @fil_cd
       AND CONVERT(datetime, CONVERT(char(10),@Data,103), 103) between b.prg_dt_ini AND b.prg_dt_fim
       AND b.prg_dt_des IS NULL
       AND a.ses_dia_semana = @diaSemana
     ORDER BY a.ses_horario
    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          Return
       End

GO

--****************************************************
/*
declare @erro int, @msgerr varchar(255), @data datetime
select @data = CONVERT(datetime, '08/10/2005', 103 )
exec upBOLETIM_TOTAL_SESSAO @data, 1, 7, @erro, @msgerr
*/
--****************************************************
CREATE PROCEDURE upBOLETIM_TOTAL_SESSAO
	(@Data		datetime,
	 @sal_cd 	int,
	 @fil_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
--	SET DATEFIRST 1
	
	DECLARE @diaSemana	smallint, 
		@dt_abertura 	datetime
	SELECT @diaSemana = 8
	  FROM tb_feriado
	 WHERE fer_data = CONVERT(datetime, @Data, 103)
    
	IF @diaSemana is null
	    SELECT @diaSemana = datepart(dw,CONVERT(datetime, @Data, 103))
	    
	SELECT z.ses_horario, 
	       z.hora_excl,
	       ISNULL(w.qtde ,0) qtde_int,
	       ISNULL(x.qtde ,0) qtde_meia,
	       ISNULL(y.qtde ,0) qtde_cort
	  FROM (((SELECT DISTINCT 
	                 a.ses_horario,
		         'N' hora_excl
		  FROM tb_sessao a INNER JOIN tb_programacao b
		       ON a.prg_cd = b.prg_cd
		  WHERE b.prg_dt_des IS NULL
		  AND   a.sal_cd         = @sal_cd
		  AND   a.fil_cd         = @fil_cd
		  AND   a.ses_dia_semana = @diaSemana
		  AND   @Data between b.prg_dt_ini AND b.prg_dt_fim
		  UNION
                  SELECT DISTINCT 
                         tb_sessao_real.sre_horario,
                         'S' hora_excl
                  FROM tb_sessao_real
                  WHERE tb_sessao_real.sre_data = @Data
   	          AND   tb_sessao_real.sal_cd   = @sal_cd
	          AND   tb_sessao_real.fil_cd   = @fil_cd
	          AND   tb_sessao_real.sre_horario NOT IN (SELECT DISTINCT a.ses_horario
		  	                                  FROM tb_sessao a INNER JOIN tb_programacao b
		  	                                       ON a.prg_cd = b.prg_cd
		  	                                  WHERE b.prg_dt_des IS NULL
		  	                                  AND   a.sal_cd         = @sal_cd
		  	                                  AND   a.fil_cd         = @fil_cd
		  	                                  AND   a.ses_dia_semana = @diaSemana
		  	                                  AND   @Data between b.prg_dt_ini AND b.prg_dt_fim)
                 ) z LEFT JOIN
	         (SELECT a.sre_horario, 
	                 COUNT(1) AS qtde
		  FROM TB_VENDA_INGRESSO a INNER JOIN TB_OPERACAO b
		       ON a.ope_cd = b.ope_cd
		  WHERE a.sal_cd   = @sal_cd
		  AND   a.fil_cd   = @fil_cd
		  AND   a.sre_data = @Data
		  AND   b.opt_cd IN (1, 3)
		  AND   a.igt_cd IN (1, 3)
		  AND   a.ing_dt_canc IS NULL
		  GROUP BY a.sre_horario
	         ) w ON z.ses_horario = w.sre_horario) LEFT JOIN 
	        (SELECT a.sre_horario, 
	                COUNT(1) AS qtde
		 FROM TB_VENDA_INGRESSO a INNER JOIN TB_OPERACAO b
		      ON a.ope_cd = b.ope_cd
		 WHERE a.sal_cd   = @sal_cd
		 AND   a.fil_cd   = @fil_cd
		 AND   a.sre_data = @Data
		 AND   b.opt_cd IN (1, 3)
		 AND   a.igt_cd IN (2, 4)
		 AND   a.ing_dt_canc IS NULL
		 GROUP BY a.sre_horario
	        ) x ON z.ses_horario = x.sre_horario) LEFT JOIN 
	       (SELECT a.sre_horario, 
	               COUNT(1) AS qtde
		FROM TB_VENDA_INGRESSO a INNER JOIN TB_OPERACAO b
		     ON a.ope_cd = b.ope_cd
		WHERE a.sal_cd   = @sal_cd
		AND   a.fil_cd   = @fil_cd
		AND   a.sre_data = @Data
		AND   b.opt_cd IN (1, 3)
		AND   a.igt_cd = 9
		AND   a.ing_dt_canc IS NULL
		GROUP BY a.sre_horario
	       ) y ON z.ses_horario = y.sre_horario 
 	ORDER BY z.ses_horario
    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          Return
       End

GO

--****************************************************
/*
declare @erro int, @msgerr varchar(255), @data datetime
select @data = CONVERT(datetime, '7/10/2005', 103 )
exec upBOLETIM_VENDA_ANTECIPADA @data, 2, 7, @erro, @msgerr
*/
--****************************************************
CREATE PROCEDURE upBOLETIM_VENDA_ANTECIPADA
	(@Data		datetime,
	 @sal_cd 	int,
	 @fil_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
	DECLARE @HoraLimite 	char(5),
	        @HoraMaxSes     datetime,
	        @DataIniPer     datetime,
	        @DataFimPer     datetime

	SELECT @HoraMaxSes = par_hora_max_ses
	  FROM tb_parametro
	  
        SELECT @DataIniPer = CONVERT(datetime, CONVERT(char(10),@Data,103) + ' ' + CONVERT(char(10),@HoraMaxSes,108), 103)
        SELECT @DataFimPer = CONVERT(datetime, CONVERT(char(10),DATEADD(Day, 1, @Data),103) + ' ' + CONVERT(char(10),@HoraMaxSes,108), 103)
	
	SELECT b.sre_data, 0 'ordem', b.igt_cd, count(1) 'qtde', 
               ISNULL(SUM(b.ing_valor),0) 'ing_valor', c.igt_desc 'igt_desc',
               b.ppr_cd, b.ing_valor 'ingvalor'
	  FROM TB_OPERACAO a,
	       TB_VENDA_INGRESSO b,
	       TB_INGRESSO_TIPO c,
	       TB_CAIXA_MOVTO d
	 WHERE b.ope_cd = a.ope_cd
	   AND d.cxp_cd = a.cxp_cd
	   AND c.igt_cd = b.igt_cd
           AND a.ope_dt_operacao between @DataIniPer AND @DataFimPer
	   AND b.sal_cd = @sal_cd
	   AND b.fil_cd = @fil_cd
	   AND b.sre_data > @Data               -- Venda Antecipada
	   AND a.opt_cd = 1			-- Operação Venda Normal
	   AND b.igt_cd IN ( 1, 2, 9 ) 		-- Inteira, Meia e Cortesia
	   AND d.cxp_talao = 0			-- não talão
	   AND a.ope_dt_des IS NULL
	   AND b.ing_dt_canc IS NULL
	GROUP BY b.sre_data, b.igt_cd, c.igt_desc, b.ppr_cd, b.ing_valor
	UNION
	SELECT b.sre_data, 2 'ordem', b.igt_cd, count(1) 'qtde', 
               ISNULL(SUM(b.ing_valor),0) 'ing_valor', c.igt_desc 'igt_desc',
               b.ppr_cd, b.ing_valor 'ingvalor'
	  FROM TB_OPERACAO a,
	       TB_VENDA_INGRESSO b,
	       TB_INGRESSO_TIPO c,
	       TB_CAIXA_MOVTO d
	 WHERE b.ope_cd = a.ope_cd
	   AND d.cxp_cd = a.cxp_cd
	   AND c.igt_cd = b.igt_cd
           AND a.ope_dt_operacao between @DataIniPer AND @DataFimPer
	   AND b.sal_cd = @sal_cd
	   AND b.fil_cd = @fil_cd
	   AND b.sre_data > @Data               -- Venda Antecipada
	   AND a.opt_cd = 1			-- Operação Venda Normal
	   AND b.igt_cd IN ( 3, 4 ) 		-- Inteira e Meia Promoção
	   AND d.cxp_talao = 0			-- não talão Promoção
	   AND a.ope_dt_des IS NULL
	   AND b.ing_dt_canc IS NULL
	GROUP BY b.sre_data, b.igt_cd, c.igt_desc, b.ppr_cd, b.ing_valor
	UNION
	SELECT b.sre_data, 4 'ordem', b.igt_cd, count(1) 'qtde', 
               ISNULL(SUM(b.ing_valor),0) 'ing_valor', c.igt_desc 'igt_desc',
               b.ppr_cd, b.ing_valor 'ingvalor'
	  FROM TB_OPERACAO a,
	       TB_VENDA_INGRESSO b,
	       TB_INGRESSO_TIPO c,
	       TB_CAIXA_MOVTO d
	 WHERE b.ope_cd = a.ope_cd
	   AND d.cxp_cd = a.cxp_cd
	   AND c.igt_cd = b.igt_cd
           AND a.ope_dt_operacao between @DataIniPer AND @DataFimPer
	   AND b.sal_cd = @sal_cd
	   AND b.fil_cd = @fil_cd
	   AND b.sre_data > @Data               -- Venda Antecipada
	   AND a.opt_cd = 1			-- Operação Venda Normal
	   AND b.igt_cd IN ( 1, 2, 9 ) 		-- Inteira, Meia e Cortesia
	   AND d.cxp_talao = 1			-- talão
	   AND a.ope_dt_des IS NULL
	   AND b.ing_dt_canc IS NULL
	GROUP BY b.sre_data, b.igt_cd, c.igt_desc, b.ppr_cd, b.ing_valor
	UNION
	SELECT b.sre_data, 6 'ordem', b.igt_cd, count(1) 'qtde', 
               ISNULL(SUM(b.ing_valor),0) 'ing_valor', c.igt_desc 'igt_desc',
               b.ppr_cd, b.ing_valor 'ingvalor'
	  FROM TB_OPERACAO a,
	       TB_VENDA_INGRESSO b,
	       TB_INGRESSO_TIPO c,
	       TB_CAIXA_MOVTO d
	 WHERE b.ope_cd = a.ope_cd
	   AND d.cxp_cd = a.cxp_cd
	   AND c.igt_cd = b.igt_cd
           AND a.ope_dt_operacao between @DataIniPer AND @DataFimPer
	   AND b.sal_cd = @sal_cd
	   AND b.fil_cd = @fil_cd
	   AND b.sre_data > @Data               -- Venda Antecipada
	   AND a.opt_cd = 1			-- Operação Venda Normal
	   AND b.igt_cd IN ( 3, 4 ) 		-- Inteira e Meia Promoção
	   AND d.cxp_talao = 1			-- talão Promoção
	   AND a.ope_dt_des IS NULL
	   AND b.ing_dt_canc IS NULL
	GROUP BY b.sre_data, b.igt_cd, c.igt_desc, b.ppr_cd, b.ing_valor
	ORDER BY 1, 2
    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          Return
       End

GO

--****************************************************
/*
declare @erro int, @msgerr varchar(255), @data datetime
select @data = getdate() --CONVERT(datetime, '10/12/2005', 103 )
exec upBOLETIM_VENDAS_DIA @data, 2, 7, @erro, @msgerr
*/
--****************************************************
CREATE PROCEDURE upBOLETIM_VENDAS_DIA
	(@Data		datetime,
	 @sal_cd 	int,
	 @fil_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
	DECLARE @HoraLimite 	char(5),
	        @HoraMaxSes     datetime,
	        @DataIniPer     datetime,
	        @DataFimPer     datetime
	        
	SELECT @HoraMaxSes = par_hora_max_ses
	  FROM tb_parametro
	  
	SELECT @DataIniPer = CONVERT(datetime, CONVERT(char(10),@Data,103) + ' ' + CONVERT(char(10),@HoraMaxSes,108), 103)
	SELECT @DataFimPer = CONVERT(datetime, CONVERT(char(10),DATEADD(Day, 1, @Data),103) + ' ' + CONVERT(char(10),@HoraMaxSes,108), 103)
	  
	SELECT 0 'ordem', b.igt_cd, count(1) 'qtde', ISNULL(SUM(b.ing_valor),0) 'ing_valor',
	       c.igt_desc 'igt_desc', b.ppr_cd, b.ing_valor 'ing_valorA'
        FROM TB_OPERACAO a,
             TB_VENDA_INGRESSO b,
             TB_INGRESSO_TIPO c,
             TB_CAIXA_MOVTO d
        WHERE d.cxp_cd   = a.cxp_cd
        AND   a.ope_cd   = b.ope_cd
        AND   c.igt_cd   = b.igt_cd
        AND   a.ope_dt_operacao between @DataIniPer AND @DataFimPer
        AND   b.sal_cd   = @sal_cd
        AND   b.fil_cd   = @fil_cd
        AND   b.sre_data = @Data
        AND a.opt_cd IN ( 1, 3 )		-- Operação Venda Normal + devolução
        AND b.igt_cd IN ( 1, 2 ) 		-- Inteira e Meia
        AND d.cxp_talao = 0			-- não talão
        GROUP BY b.igt_cd, c.igt_desc, b.ppr_cd, b.ing_valor
	UNION
	SELECT 2 'ordem', b.igt_cd, count(1) 'qtde', ISNULL(SUM(b.ing_valor),0) 'ing_valor',
	       c.igt_desc 'igt_desc', b.ppr_cd, b.ing_valor 'ing_valorA'
        FROM TB_OPERACAO a,
             TB_VENDA_INGRESSO b,
             TB_INGRESSO_TIPO c,
             TB_CAIXA_MOVTO d
        WHERE d.cxp_cd   = a.cxp_cd
        AND   a.ope_cd   = b.ope_cd
        AND   c.igt_cd   = b.igt_cd
        AND   a.ope_dt_operacao between @DataIniPer AND @DataFimPer
        AND   b.sal_cd   = @sal_cd
        AND   b.fil_cd   = @fil_cd
        AND   b.sre_data = @Data
        AND a.opt_cd IN ( 1, 3 )		-- Operação Venda Normal + devolução
        AND b.igt_cd IN ( 3, 4 ) 		-- Inteira e Meia Promoção
        AND d.cxp_talao = 0			-- não talão
        GROUP BY b.igt_cd, c.igt_desc, b.ppr_cd, b.ing_valor
	UNION
	SELECT 4 'ordem', b.igt_cd, count(1) 'qtde', ISNULL(SUM(b.ing_valor),0) 'ing_valor',
	       c.igt_desc + ' TA' 'igt_desc', b.ppr_cd, b.ing_valor 'ing_valorA'
        FROM TB_OPERACAO a,
             TB_VENDA_INGRESSO b,
             TB_INGRESSO_TIPO c,
             TB_CAIXA_MOVTO d
        WHERE d.cxp_cd   = a.cxp_cd
        AND   a.ope_cd   = b.ope_cd
        AND   c.igt_cd   = b.igt_cd
        AND   a.ope_dt_operacao between @DataIniPer AND @DataFimPer
        AND   b.sal_cd   = @sal_cd
        AND   b.fil_cd   = @fil_cd
        AND   b.sre_data = @Data
        AND a.opt_cd IN ( 1, 3 )		-- Operação Venda Normal + devolução
        AND b.igt_cd IN ( 1, 2 ) 		-- Inteira e Meia
        AND d.cxp_talao = 1			-- Talão
        GROUP BY b.igt_cd, c.igt_desc, b.ppr_cd, b.ing_valor
	UNION
	SELECT 6 'ordem', b.igt_cd, count(1) 'qtde', ISNULL(SUM(b.ing_valor),0) 'ing_valor',
	       c.igt_desc + ' TA' 'igt_desc', b.ppr_cd, b.ing_valor 'ing_valorA'
        FROM TB_OPERACAO a,
             TB_VENDA_INGRESSO b,
             TB_INGRESSO_TIPO c,
             TB_CAIXA_MOVTO d
        WHERE d.cxp_cd   = a.cxp_cd
        AND   a.ope_cd   = b.ope_cd
        AND   c.igt_cd   = b.igt_cd
        AND   a.ope_dt_operacao between @DataIniPer AND @DataFimPer
        AND   b.sal_cd   = @sal_cd
        AND   b.fil_cd   = @fil_cd
        AND   b.sre_data = @Data
        AND a.opt_cd IN ( 1, 3 )		-- Operação Venda Normal + devolução
        AND b.igt_cd IN ( 3, 4 ) 		-- Inteira e Meia Promoção
        AND d.cxp_talao = 1			-- Talão
        GROUP BY b.igt_cd, c.igt_desc, b.ppr_cd, b.ing_valor
        ORDER BY 1
        
    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          Return
       End

GO


--****************************************************
CREATE PROCEDURE upBOLETIM_CATRACA
	(@Data		datetime,
	 @sal_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS

    SELECT tb_catraca.cat_nm,
           tb_catraca_cont.ctc_ini_cont,
           tb_catraca_cont.ctc_fim_cont
    FROM tb_catraca_cont,
         tb_catraca_sala,
         tb_catraca
    WHERE tb_catraca_cont.cat_cd = tb_catraca_sala.cat_cd
    AND   tb_catraca_cont.cat_cd = tb_catraca.cat_cd
    AND   tb_catraca_cont.ctc_dt = @Data
    AND   tb_catraca_sala.sal_cd = @sal_cd

    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          RETURN
       END

GO

--****************************************************
CREATE PROCEDURE upBOLETIM_INGRESSO_S_USO
	(@Data		datetime,
	 @sal_cd 	int,
	 @fil_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS

    SELECT COUNT(*) AS qtdeSUso
    FROM tb_venda_ingresso
    WHERE sal_cd     = @sal_cd
    AND   fil_cd     = @fil_cd
    AND   sre_data   = @Data
    AND   ing_status <> 1
    AND   ing_dt_canc IS NULL


    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          RETURN
       END

GO

--****************************************************
CREATE PROCEDURE upBOLETIM_VENDA_COMBO
   (@dt_ini datetime,
    @dt_fim datetime,
    @Erro   int OUTPUT,
    @MsgErr varchar(255) OUTPUT)
AS
   DECLARE @dtFimAux DATETIME
   
   SELECT @dtFimAux = DATEADD(d, 1, @dt_fim)

   SELECT CONVERT(DATETIME, CONVERT(CHAR(10), TB_OPERACAO.ope_dt_operacao, 103), 103) AS ope_dt_operacao,
          TB_COMBO.cbo_cd, 
          TB_COMBO.cbo_nm,
          ISNULL(SUM(TB_VENDA_COMBO.vcb_qtde),0)  'qtde',
          ISNULL(SUM(TB_VENDA_COMBO.vcb_valor * TB_VENDA_COMBO.vcb_qtde),0) 'valor'
        FROM TB_OPERACAO,
             TB_VENDA_COMBO,
             TB_COMBO
        WHERE  TB_OPERACAO.ope_cd        = TB_VENDA_COMBO.ope_cd
        AND   TB_VENDA_COMBO.cbo_cd      = TB_COMBO.cbo_cd 
        AND   TB_OPERACAO.ope_dt_des     IS NULL
        AND   TB_VENDA_COMBO.vcb_dt_canc IS NULL
        AND   TB_OPERACAO.ope_dt_operacao BETWEEN @dt_ini AND @dtFimAux
        GROUP BY CONVERT(DATETIME, CONVERT(CHAR(10), TB_OPERACAO.ope_dt_operacao, 103), 103),
                 TB_COMBO.cbo_cd, 
                 TB_COMBO.cbo_nm
        ORDER BY 1,
                 TB_COMBO.cbo_cd, 
                 TB_COMBO.cbo_nm
   
     SELECT @Erro = @@ERROR
   
     IF @Erro <> 0
        BEGIN
           SELECT @MsgErr = description
           FROM master..sysmessages
           WHERE error = @Erro
         
           RETURN
        END

GO

--****************************************************
CREATE PROCEDURE upBOLETIM_VENDA_COMBO_TOTAL
   (@dt_ini datetime,
    @dt_fim datetime,
    @Erro   int OUTPUT,
    @MsgErr varchar(255) OUTPUT)
AS
   DECLARE @dtFimAux DATETIME
   
   SELECT @dtFimAux = DATEADD(d, 1, @dt_fim)

   SELECT TB_COMBO.cbo_cd, 
          TB_COMBO.cbo_nm,
          ISNULL(SUM(TB_VENDA_COMBO.vcb_qtde),0)  'qtde',
          ISNULL(SUM(TB_VENDA_COMBO.vcb_valor * TB_VENDA_COMBO.vcb_qtde),0) 'valor'
        FROM TB_OPERACAO,
             TB_VENDA_COMBO,
             TB_COMBO
        WHERE  TB_OPERACAO.ope_cd        = TB_VENDA_COMBO.ope_cd
        AND   TB_VENDA_COMBO.cbo_cd      = TB_COMBO.cbo_cd 
        AND   TB_OPERACAO.ope_dt_des     IS NULL
        AND   TB_VENDA_COMBO.vcb_dt_canc IS NULL
        AND   TB_OPERACAO.ope_dt_operacao BETWEEN @dt_ini AND @dtFimAux
        GROUP BY TB_COMBO.cbo_cd, 
                 TB_COMBO.cbo_nm
        ORDER BY TB_COMBO.cbo_cd, 
                 TB_COMBO.cbo_nm
   
     SELECT @Erro = @@ERROR
   
     IF @Erro <> 0
        BEGIN
           SELECT @MsgErr = description
           FROM master..sysmessages
           WHERE error = @Erro
         
           RETURN
        END

GO


--****************************************************
CREATE PROCEDURE upBOLETIM_VENDA_INGRESSO
   (@dt_ini datetime,
    @dt_fim datetime,
    @Erro   int OUTPUT,
    @MsgErr varchar(255) OUTPUT)
AS
   DECLARE @dtFimAux DATETIME
   
   SELECT @dtFimAux = DATEADD(d, 1, @dt_fim)

   SELECT CONVERT(DATETIME, CONVERT(CHAR(10), TB_OPERACAO.ope_dt_operacao, 103), 103) AS ope_dt_operacao,
          TB_INGRESSO_TIPO.igt_cd, 
          TB_INGRESSO_TIPO.igt_desc,
          SUM(1) AS 'qtde',
          ISNULL(SUM(TB_VENDA_INGRESSO.ing_valor),0) AS 'valor'
        FROM TB_OPERACAO,
             TB_VENDA_INGRESSO,
             TB_INGRESSO_TIPO
        WHERE TB_OPERACAO.ope_cd            = TB_VENDA_INGRESSO.ope_cd
        AND   TB_VENDA_INGRESSO.igt_cd      = TB_INGRESSO_TIPO.igt_cd 
        AND   TB_OPERACAO.ope_dt_des        IS NULL
        AND   TB_VENDA_INGRESSO.ing_dt_canc IS NULL
        AND   TB_OPERACAO.ope_dt_operacao BETWEEN @dt_ini AND @dtFimAux
        GROUP BY CONVERT(DATETIME, CONVERT(CHAR(10), TB_OPERACAO.ope_dt_operacao, 103), 103),
                 TB_INGRESSO_TIPO.igt_cd, 
                 TB_INGRESSO_TIPO.igt_desc
        ORDER BY 1,
                 TB_INGRESSO_TIPO.igt_cd, 
                 TB_INGRESSO_TIPO.igt_desc
   
     SELECT @Erro = @@ERROR
   
     IF @Erro <> 0
        BEGIN
           SELECT @MsgErr = description
           FROM master..sysmessages
           WHERE error = @Erro
         
           RETURN
        END

GO

--****************************************************
CREATE PROCEDURE upBOLETIM_VENDA_INGRESSO_TOTAL
   (@dt_ini datetime,
    @dt_fim datetime,
    @Erro   int OUTPUT,
    @MsgErr varchar(255) OUTPUT)
AS
   DECLARE @dtFimAux DATETIME
   
   SELECT @dtFimAux = DATEADD(d, 1, @dt_fim)

   SELECT TB_INGRESSO_TIPO.igt_cd, 
          TB_INGRESSO_TIPO.igt_desc,
          SUM(1) AS 'qtde',
          ISNULL(SUM(TB_VENDA_INGRESSO.ing_valor),0) AS 'valor'
        FROM TB_OPERACAO,
             TB_VENDA_INGRESSO,
             TB_INGRESSO_TIPO
        WHERE TB_OPERACAO.ope_cd            = TB_VENDA_INGRESSO.ope_cd
        AND   TB_VENDA_INGRESSO.igt_cd      = TB_INGRESSO_TIPO.igt_cd 
        AND   TB_OPERACAO.ope_dt_des        IS NULL
        AND   TB_VENDA_INGRESSO.ing_dt_canc IS NULL
        AND   TB_OPERACAO.ope_dt_operacao BETWEEN @dt_ini AND @dtFimAux
        GROUP BY TB_INGRESSO_TIPO.igt_cd, 
                 TB_INGRESSO_TIPO.igt_desc
        ORDER BY TB_INGRESSO_TIPO.igt_cd, 
                 TB_INGRESSO_TIPO.igt_desc
   
     SELECT @Erro = @@ERROR
   
     IF @Erro <> 0
        BEGIN
           SELECT @MsgErr = description
           FROM master..sysmessages
           WHERE error = @Erro
         
           RETURN
        END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM2_GERA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM2_GERA
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upVER_TIPO_INGRESSO') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upVER_TIPO_INGRESSO
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upVER_TIPO_PAGAMENTO') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upVER_TIPO_PAGAMENTO
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upVER_TIPO_OPERACAO') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upVER_TIPO_OPERACAO
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upINCLUI_PARAM') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upINCLUI_PARAM
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upINCLUI_EMPRESA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upINCLUI_EMPRESA
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upINCLUI_CINEMA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upINCLUI_CINEMA
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upINCLUI_CATRACA_SALA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upINCLUI_CATRACA_SALA
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upINCLUI_CATRACA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upINCLUI_CATRACA
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upINCLUI_SALA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upINCLUI_SALA
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upINCLUI_DISTRIBUIDORA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upINCLUI_DISTRIBUIDORA
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upINCLUI_FILME') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upINCLUI_FILME
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upINCLUI_BOLETIM') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upINCLUI_BOLETIM
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upINCLUI_VENDA_INGR') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upINCLUI_VENDA_INGR
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upINCLUI_BOL_TALAO') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upINCLUI_BOL_TALAO
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upINCLUI_BOL_SESSAO') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upINCLUI_BOL_SESSAO
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_CAPA2') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_CAPA2
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_VENDAS_DIA2') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_VENDAS_DIA2
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_PRE_VENDA2') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_PRE_VENDA2
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_CORTESIA2') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_CORTESIA2
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_DEVOLUCAO2') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_DEVOLUCAO2
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_SESSOES_FILME2') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_SESSOES_FILME2
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_TOTAL_SESSAO2') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_TOTAL_SESSAO2
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_VENDA_ANTECIPADA2') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_VENDA_ANTECIPADA2
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_FORMA_PAGTO2') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_FORMA_PAGTO2
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_NUMERACAO_TALAO2') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_NUMERACAO_TALAO2
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_BOL_FILME_CARTAZ') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_BOL_FILME_CARTAZ
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_BOLETIM_VERIFICA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_BOLETIM_VERIFICA
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_CATRACA2') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_CATRACA2
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_INGRESSO_S_USO2') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_INGRESSO_S_USO2
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upBOLETIM_VENDAS_TOTAL2') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upBOLETIM_VENDAS_TOTAL2
GO


--****************************************************
CREATE PROCEDURE upVER_TIPO_INGRESSO
   (@Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS
   INSERT INTO tb_bol_tp_ingr (igt_cd, igt_desc)
   SELECT igt_cd,
          igt_desc
   FROM tb_ingresso_tipo
   WHERE igt_cd NOT IN (SELECT igt_cd
                        FROM tb_bol_tp_ingr)

   SELECT @Erro = @@ERROR
    
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

   UPDATE tb_bol_tp_ingr
   SET  tb_bol_tp_ingr.igt_desc = tb_ingresso_tipo.igt_desc
   FROM tb_bol_tp_ingr,
        tb_ingresso_tipo
   WHERE tb_ingresso_tipo.igt_cd   = tb_bol_tp_ingr.igt_cd
   AND   tb_ingresso_tipo.igt_desc <> tb_bol_tp_ingr.igt_desc

   SELECT @Erro = @@ERROR
    
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END
GO

--****************************************************
CREATE PROCEDURE upVER_TIPO_PAGAMENTO
   (@Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS
   INSERT INTO tb_bol_pag_tp (pgt_cd, pgt_desc)
   SELECT pgt_cd,
          pgt_desc
   FROM tb_pagamento_tipo
   WHERE pgt_cd NOT IN (SELECT pgt_cd
                        FROM tb_bol_pag_tp)

   SELECT @Erro = @@ERROR
    
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

   UPDATE tb_bol_pag_tp
   SET  tb_bol_pag_tp.pgt_desc = tb_pagamento_tipo.pgt_desc
   FROM tb_bol_tp_ingr,
        tb_pagamento_tipo
   WHERE tb_pagamento_tipo.pgt_cd   = tb_bol_pag_tp.pgt_cd
   AND   tb_pagamento_tipo.pgt_desc <> tb_bol_pag_tp.pgt_desc

   SELECT @Erro = @@ERROR
    
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END
GO

--****************************************************
CREATE PROCEDURE upVER_TIPO_OPERACAO
   (@Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS
   INSERT INTO tb_bol_ope_tp (opt_cd, opt_desc, opt_sinal)
   SELECT opt_cd,
          opt_desc,
          opt_sinal
   FROM tb_operacao_tipo
   WHERE opt_cd NOT IN (SELECT opt_cd
                        FROM tb_bol_ope_tp)

   SELECT @Erro = @@ERROR
    
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

   UPDATE tb_bol_ope_tp
   SET  tb_bol_ope_tp.opt_desc  = tb_operacao_tipo.opt_desc,
        tb_bol_ope_tp.opt_sinal = tb_operacao_tipo.opt_sinal
   FROM tb_bol_ope_tp,
        tb_operacao_tipo
   WHERE tb_operacao_tipo.opt_cd     = tb_bol_ope_tp.opt_cd
   AND   (tb_operacao_tipo.opt_desc  <> tb_bol_ope_tp.opt_desc
   OR     tb_operacao_tipo.opt_sinal <> tb_bol_ope_tp.opt_sinal)

   SELECT @Erro = @@ERROR
    
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END
GO


--****************************************************
CREATE PROCEDURE upINCLUI_PARAM
   (@DataMov datetime,
    @Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS
   DECLARE @emp_cd      int,
           @cin_cd      int
           
   SELECT @emp_cd = emp_cd
   FROM tb_empresa

   SELECT @Erro = @@ERROR   

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END
   
   SELECT @cin_cd = cin_cd
   FROM tb_cinema

   SELECT @Erro = @@ERROR   

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

   DELETE FROM tb_bol_param 
   WHERE bol_dt_mov = @DataMov

   SELECT @Erro = @@ERROR   

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

   INSERT INTO tb_bol_param 
          (bol_dt_mov,
           emp_cd,
           cin_cd,
           par_hora_max_ses, 
           par_hora_limite, 
           par_hora_limite12,
           par_hora_limite23,
           par_hora_limite34,
           par_hora_limite45,
           par_hora_limite56,
           par_custo_ingresso,
           par_imposto_mun, 
           par_direitos_aut, 
           par_outros,
           par_perc_meias,
           par_perc_cortesias)
   SELECT @DataMov,
          @emp_cd,
          @cin_cd,
          par_hora_max_ses, 
          par_hora_limite, 
          par_hora_limite12,
          par_hora_limite23,
          par_hora_limite34,
          par_hora_limite45,
          par_hora_limite56,
          par_custo_ingresso,
          par_imposto_mun, 
          par_direitos_aut, 
          par_outros,
          par_perc_meias,
          par_perc_cortesias
   FROM tb_parametro
   
   SELECT @Erro = @@ERROR   

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

GO

--****************************************************
CREATE PROCEDURE upINCLUI_EMPRESA
   (@DataMov datetime,
    @Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS
   DELETE FROM tb_bol_empr
   WHERE bol_dt_mov = @DataMov

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

   INSERT INTO tb_bol_empr
         (bol_dt_mov,
          emp_cd,
          emp_nm)
   SELECT @DataMov,
          emp_cd,
          emp_nm
   FROM tb_empresa
   
   SELECT @Erro = @@ERROR

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

GO

--****************************************************
CREATE PROCEDURE upINCLUI_CINEMA
   (@DataMov datetime,
    @Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS
   DECLARE @emp_cd      int
           
   SELECT @emp_cd = emp_cd
   FROM tb_empresa

   SELECT @Erro = @@ERROR   

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

   DELETE FROM tb_bol_cin
   WHERE bol_dt_mov = @DataMov
   
   SELECT @Erro = @@ERROR

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

   INSERT INTO tb_bol_cin
         (bol_dt_mov,
          emp_cd,
          cin_cd,
          cin_nm,
          cin_cnpj,
          cin_end,
          cin_num_end,
          cin_cmp_end,
          cin_brr_end,
          cin_cid_end,
          cin_uf_end,
          cin_cep_end)
   SELECT @DataMov,
          @emp_cd,
          cin_cd,
          cin_nm,
          cin_cnpj,
          cin_end,
          cin_num_end,
          cin_cmp_end,
          cin_brr_end,
          cin_cid_end,
          cin_uf_end,
          cin_cep_end
   FROM tb_cinema
   
   SELECT @Erro = @@ERROR

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

GO

--****************************************************
CREATE PROCEDURE upINCLUI_CATRACA
   (@DataMov datetime,
    @Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS
   DECLARE @emp_cd      int,
           @cin_cd      int
           
   SELECT @emp_cd = emp_cd
   FROM tb_empresa

   SELECT @Erro = @@ERROR   

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END
   
   SELECT @cin_cd = cin_cd
   FROM tb_cinema

   SELECT @Erro = @@ERROR   

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

   DELETE FROM tb_bol_catraca_sala
   WHERE bol_dt_mov = @DataMov

   SELECT @Erro = @@ERROR

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

   DELETE FROM tb_bol_catraca
   WHERE bol_dt_mov = @DataMov

   SELECT @Erro = @@ERROR

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

   INSERT INTO tb_bol_catraca
         (bol_dt_mov,
          emp_cd,
          cin_cd,
          cat_cd,
          cat_nm,
          ctc_ini_cont,
          ctc_fim_cont)
   SELECT @DataMov,
          @emp_cd,
          @cin_cd,
          tb_catraca_cont.cat_cd,
          tb_catraca.cat_nm,
          tb_catraca_cont.ctc_ini_cont,
          tb_catraca_cont.ctc_fim_cont
   FROM tb_catraca_cont,
        tb_catraca
   WHERE tb_catraca_cont.cat_cd = tb_catraca.cat_cd
   AND   tb_catraca_cont.ctc_dt = @DataMov
   
   SELECT @Erro = @@ERROR

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

GO

--****************************************************
CREATE PROCEDURE upINCLUI_SALA
   (@DataMov datetime,
    @Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS
   DECLARE @emp_cd      int,
           @cin_cd      int
           
   SELECT @emp_cd = emp_cd
   FROM tb_empresa

   SELECT @Erro = @@ERROR   

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END
   
   SELECT @cin_cd = cin_cd
   FROM tb_cinema

   SELECT @Erro = @@ERROR   

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

   DELETE FROM tb_bol_catraca_sala
   WHERE bol_dt_mov = @DataMov

   SELECT @Erro = @@ERROR

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

   DELETE FROM tb_bol_sala
   WHERE bol_dt_mov = @DataMov
   
   SELECT @Erro = @@ERROR

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

   INSERT INTO tb_bol_sala
         (bol_dt_mov,
          emp_cd,
          cin_cd,
          sal_cd,
          sal_desc,
          sal_lugares)
   SELECT @DataMov,
          @emp_cd,
          @cin_cd,
          tb_sala.sal_cd,
          tb_sala.sal_desc,
          tb_sala.sal_lugares
   FROM tb_sala
        
   
   SELECT @Erro = @@ERROR

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

GO

--****************************************************
CREATE PROCEDURE upINCLUI_CATRACA_SALA
   (@DataMov datetime,
    @Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS
   DECLARE @emp_cd      int,
           @cin_cd      int
           
   SELECT @emp_cd = emp_cd
   FROM tb_empresa

   SELECT @Erro = @@ERROR   

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END
   
   SELECT @cin_cd = cin_cd
   FROM tb_cinema

   SELECT @Erro = @@ERROR   

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

   DELETE FROM tb_bol_catraca_sala
   WHERE bol_dt_mov = @DataMov

   SELECT @Erro = @@ERROR

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

   INSERT INTO tb_bol_catraca_sala
         (bol_dt_mov,
          emp_cd,
          cin_cd,
          cat_cd,
          sal_cd)
   SELECT @DataMov,
          @emp_cd,
          @cin_cd,
          tb_catraca_sala.cat_cd,
          tb_catraca_sala.sal_cd
   FROM tb_catraca_sala,
        tb_catraca_cont
   WHERE tb_catraca_sala.cat_cd = tb_catraca_cont.cat_cd
   AND   tb_catraca_cont.ctc_dt = @DataMov
        
   
   SELECT @Erro = @@ERROR

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

GO


--****************************************************
CREATE PROCEDURE upINCLUI_DISTRIBUIDORA
   (@DataMov datetime,
    @Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS

   DELETE FROM tb_bol_distrib
   WHERE tb_bol_distrib.bol_dt_mov = @DataMov
   
   SELECT @Erro = @@ERROR

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

   INSERT INTO tb_bol_distrib
         (bol_dt_mov,
          dis_cd,
          dis_nm)
   SELECT DISTINCT @DataMov,
          dis_cd,
          dis_nm
   FROM tb_distribuidora
   
   SELECT @Erro = @@ERROR

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

GO



--****************************************************
CREATE PROCEDURE upINCLUI_FILME
   (@DataMov datetime,
    @Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS
   DECLARE @diaSemana smallint,
           @emp_cd    int,
           @cin_cd    int
           

--   SET DATEFIRST 1

   SELECT @diaSemana = dbo.ufDiaSemana(@DataMov)
   
   SELECT @emp_cd = emp_cd
   FROM tb_empresa

   SELECT @Erro = @@ERROR   

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END
   
   SELECT @cin_cd = cin_cd
   FROM tb_cinema

   SELECT @Erro = @@ERROR   

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

   DELETE FROM tb_bol_filme
   WHERE tb_bol_filme.bol_dt_mov = @DataMov
   
   SELECT @Erro = @@ERROR

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

   INSERT INTO tb_bol_filme
         (bol_dt_mov,
          emp_cd,
          cin_cd,
          fil_cd,
          fil_nm,
          fil_dt_ini,
          fil_dt_fim,
          dis_cd,
          fil_censura,
          fil_id_nacio)
   SELECT DISTINCT @DataMov,
          @emp_cd,
          @cin_cd,
          tb_filme.fil_cd,
          tb_filme.fil_nm,
          fil_dt_ini,
          fil_dt_fim,
          tb_filme.dis_cd,
          tb_filme.fil_censura,
          tb_filme.fil_id_nacio
   FROM tb_programacao,
        tb_sessao,
        tb_filme
   WHERE tb_programacao.prg_cd = tb_sessao.prg_cd
   AND   tb_sessao.fil_cd      = tb_filme.fil_cd
   AND   tb_programacao.prg_dt_des IS NULL
   AND   @DataMov BETWEEN tb_programacao.prg_dt_ini AND tb_programacao.prg_dt_fim
   AND   tb_sessao.ses_dt_des IS NULL
   AND   tb_sessao.ses_dia_semana = @diaSemana
   AND   @DataMov between tb_filme.fil_dt_ini AND tb_filme.fil_dt_fim
   UNION
   SELECT DISTINCT @DataMov,
          @emp_cd,
          @cin_cd,
          tb_filme.fil_cd,
          tb_filme.fil_nm,
          tb_filme.fil_dt_ini,
          tb_filme.fil_dt_fim,
          tb_filme.dis_cd,
          tb_filme.fil_censura,
          tb_filme.fil_id_nacio
   FROM tb_venda_ingresso,
        tb_filme
   WHERE tb_venda_ingresso.fil_cd        = tb_filme.fil_cd
   AND   (tb_venda_ingresso.sre_data     = @DataMov
   OR     tb_venda_ingresso.ing_dt_venda > @DataMov)
   
   SELECT @Erro = @@ERROR

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

GO

--****************************************************
CREATE PROCEDURE upINCLUI_BOLETIM
   (@DataMov datetime,
    @Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS
   DECLARE @dt_abertura datetime,
           @dt_emissao  datetime,
           @diaSemana   smallint,
           @HoraMaxSes  datetime,
           @DataIniPer  datetime,
           @DataFimPer  datetime

   SELECT @diaSemana = dbo.ufDiaSemana(@DataMov)

   SELECT @HoraMaxSes = par_hora_max_ses
   FROM tb_parametro
	  
   SELECT @DataIniPer = CONVERT(datetime, CONVERT(char(10),@DataMov,103) + ' ' + CONVERT(char(10),@HoraMaxSes,108), 103)
   SELECT @DataFimPer = CONVERT(datetime, CONVERT(char(10),DATEADD(Day, 1, @DataMov),103) + ' ' + CONVERT(char(10),@HoraMaxSes,108), 103)

   SELECT @dt_abertura = MIN(cxp_dt_abertura) 
   FROM tb_caixa_movto
   WHERE CONVERT(char(10),cxp_dt_abertura,102) = CONVERT(char(10),@DataMov,102)
   
   SELECT @Erro = @@ERROR

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END
   
   SELECT @dt_emissao = getdate()
   
   INSERT INTO tb_boletim
                   (bol_dt_mov,
                    emp_cd,
                    cin_cd,
                    sal_cd,
                    fil_cd,
                    bol_dt_abertura,
                    bol_dt_emissao,
                    bol_dt_ini_per,
                    bol_dt_fim_per,
                    bol_status)
   SELECT DISTINCT @DataMov,
                   tb_cinema.emp_cd,
                   tb_sala.cin_cd,
                   tb_sessao.sal_cd,
                   tb_sessao.fil_cd,
                   @dt_abertura,
                   @dt_emissao,
                   tb_programacao.prg_dt_ini, 
                   tb_programacao.prg_dt_fim,
                   'N'
   FROM tb_sessao,
        tb_programacao,
        tb_filme,
        tb_sala,
        tb_cinema
   WHERE tb_sessao.prg_cd = tb_programacao.prg_cd
   AND   tb_sessao.fil_cd = tb_filme.fil_cd
   AND   tb_sessao.sal_cd = tb_sala.sal_cd
   AND   tb_sala.cin_cd   = tb_cinema.cin_cd
   AND   tb_sessao.ses_dt_des      IS NULL
   AND   tb_programacao.prg_dt_des IS NULL
   AND   tb_filme.fil_dt_des       IS NULL
   AND   @DataMov between tb_programacao.prg_dt_ini and tb_programacao.prg_dt_fim
   AND   tb_sessao.ses_dia_semana = dbo.ufDiaSemana(@DataMov)  
   UNION	
   SELECT DISTINCT @DataMov,
                   tb_cinema.emp_cd,
                   tb_sala.cin_cd,
                   tb_venda_ingresso.sal_cd,
                   tb_venda_ingresso.fil_cd,
                   @dt_abertura,
                   @dt_emissao,
                   NULL prg_dt_ini, 
                   NULL prg_dt_fim,
                   'S'
   FROM tb_venda_ingresso,
        tb_filme,
        tb_sala,
        tb_cinema
   WHERE tb_venda_ingresso.fil_cd       = tb_filme.fil_cd
   AND   tb_venda_ingresso.sal_cd       = tb_sala.sal_cd
   AND   tb_sala.cin_cd                 = tb_cinema.cin_cd
   AND   tb_venda_ingresso.sre_data     = @DataMov
   --AND   tb_venda_ingresso.ing_dt_venda BETWEEN @DataIniPer AND @DataFimPer   
   AND   tb_filme.fil_dt_des       IS NULL
   AND   convert(varchar(10),tb_venda_ingresso.fil_cd) + 
         convert(varchar(3),tb_venda_ingresso.sal_cd) NOT IN
               (SELECT DISTINCT convert(varchar(10),tb_sessao.fil_cd) + 
                                convert(varchar(3),tb_sessao.sal_cd)
                FROM tb_sessao,
                     tb_programacao
                WHERE tb_sessao.prg_cd = tb_programacao.prg_cd
                AND   tb_sessao.ses_dt_des      IS NULL
                AND   tb_programacao.prg_dt_des IS NULL
                AND   @DataMov between tb_programacao.prg_dt_ini and tb_programacao.prg_dt_fim
                AND   tb_sessao.ses_dia_semana = dbo.ufDiaSemana(@DataMov))  
   UNION
   SELECT DISTINCT @DataMov,
                   tb_cinema.emp_cd,
                   tb_sala.cin_cd,
                   tb_venda_ingresso.sal_cd,
                   tb_venda_ingresso.fil_cd,
                   @dt_abertura,
                   @dt_emissao,
                   tb_programacao.prg_dt_ini, 
                   tb_programacao.prg_dt_fim,
                   'P'
   FROM tb_venda_ingresso,
        tb_sessao,
        tb_programacao,
        tb_filme,
        tb_sala,
        tb_cinema
   WHERE tb_venda_ingresso.fil_cd                    = tb_sessao.fil_cd
   AND   tb_venda_ingresso.sal_cd                    = tb_sessao.sal_cd
   AND   tb_venda_ingresso.sre_horario               = tb_sessao.ses_horario
   AND   dbo.ufDiaSemana(tb_venda_ingresso.sre_data) = tb_sessao.ses_dia_semana
   AND   tb_sessao.prg_cd                            = tb_programacao.prg_cd
   AND   tb_venda_ingresso.fil_cd   = tb_filme.fil_cd
   AND   tb_venda_ingresso.sal_cd   = tb_sala.sal_cd
   AND   tb_sala.cin_cd             = tb_cinema.cin_cd
   AND   tb_venda_ingresso.sre_data <> @DataMov
   AND   tb_venda_ingresso.ing_dt_canc IS NULL
   AND   tb_venda_ingresso.ing_dt_venda BETWEEN @DataIniPer AND @DataFimPer
   AND convert(varchar(10),tb_venda_ingresso.fil_cd) + 
       convert(varchar(3),tb_venda_ingresso.sal_cd) NOT IN
               (SELECT DISTINCT convert(varchar(10),tb_sessao.fil_cd) + 
                                convert(varchar(3),tb_sessao.sal_cd)
                FROM tb_sessao,
                     tb_programacao
                WHERE tb_sessao.prg_cd = tb_programacao.prg_cd
                AND tb_sessao.ses_dt_des      IS NULL
                AND tb_programacao.prg_dt_des IS NULL
                AND @DataMov BETWEEN tb_programacao.prg_dt_ini and tb_programacao.prg_dt_fim
                AND tb_sessao.ses_dia_semana = dbo.ufDiaSemana(@DataMov))  
   UNION
   SELECT DISTINCT @DataMov,
                   tb_cinema.emp_cd,
                   tb_sala.cin_cd,
                   tb_venda_ingresso.sal_cd,
                   tb_venda_ingresso.fil_cd,
                   @dt_abertura,
                   @dt_emissao,
                   NULL prg_dt_ini, 
                   NULL prg_dt_fim,
                   'Q'
   FROM tb_venda_ingresso,
        tb_filme,
        tb_sala,
        tb_cinema
   WHERE tb_venda_ingresso.fil_cd   = tb_filme.fil_cd
   AND   tb_venda_ingresso.sal_cd   = tb_sala.sal_cd
   AND   tb_sala.cin_cd             = tb_cinema.cin_cd
   AND   tb_venda_ingresso.sre_data <> @DataMov
   AND   tb_venda_ingresso.ing_dt_canc IS NULL
   AND   tb_venda_ingresso.ing_dt_venda BETWEEN @DataIniPer AND @DataFimPer
   AND convert(varchar(10),tb_venda_ingresso.fil_cd) + 
       convert(varchar(3),tb_venda_ingresso.sal_cd) NOT IN
               (SELECT DISTINCT convert(varchar(10),tb_sessao.fil_cd) + 
                                convert(varchar(3),tb_sessao.sal_cd)
                FROM tb_sessao,
                     tb_programacao
                WHERE tb_sessao.prg_cd = tb_programacao.prg_cd
                AND tb_sessao.ses_dt_des      IS NULL
                AND tb_programacao.prg_dt_des IS NULL
                AND @DataMov BETWEEN tb_programacao.prg_dt_ini and tb_programacao.prg_dt_fim
                AND tb_sessao.ses_dia_semana = dbo.ufDiaSemana(@DataMov))  
   AND convert(varchar(10),tb_venda_ingresso.fil_cd) + 
       convert(varchar(3),tb_venda_ingresso.sal_cd) NOT IN
               (SELECT DISTINCT convert(varchar(10),tb_sessao.fil_cd) + 
                                convert(varchar(3),tb_sessao.sal_cd)
                FROM tb_sessao,
                     tb_programacao
                WHERE tb_sessao.prg_cd = tb_programacao.prg_cd
                AND tb_sessao.ses_dt_des      IS NULL
                AND tb_programacao.prg_dt_des IS NULL
                AND tb_venda_ingresso.sre_data BETWEEN tb_programacao.prg_dt_ini and tb_programacao.prg_dt_fim
                AND tb_sessao.ses_dia_semana = dbo.ufDiaSemana(tb_venda_ingresso.sre_data))  
   AND convert(varchar(10),tb_venda_ingresso.fil_cd) + 
       convert(varchar(3),tb_venda_ingresso.sal_cd) NOT IN
               (SELECT DISTINCT convert(varchar(10),tb_venda_ingresso.fil_cd) + 
                                convert(varchar(3),tb_venda_ingresso.sal_cd)
                FROM tb_venda_ingresso,
                     tb_filme,
                     tb_sala,
                     tb_cinema
                WHERE tb_venda_ingresso.fil_cd   = tb_filme.fil_cd
                AND   tb_venda_ingresso.sal_cd   = tb_sala.sal_cd
                AND   tb_sala.cin_cd             = tb_cinema.cin_cd
                AND   tb_venda_ingresso.sre_data = @DataMov
                AND   tb_filme.fil_dt_des       IS NULL
                AND   convert(varchar(10),tb_venda_ingresso.fil_cd) + 
                      convert(varchar(3),tb_venda_ingresso.sal_cd) NOT IN
                      (SELECT DISTINCT convert(varchar(10),tb_sessao.fil_cd) + 
                                       convert(varchar(3),tb_sessao.sal_cd)
                       FROM tb_sessao,
                            tb_programacao
                       WHERE tb_sessao.prg_cd = tb_programacao.prg_cd
                       AND   tb_sessao.ses_dt_des      IS NULL
                       AND   tb_programacao.prg_dt_des IS NULL
                       AND   @DataMov between tb_programacao.prg_dt_ini and tb_programacao.prg_dt_fim
                       AND   tb_sessao.ses_dia_semana = dbo.ufDiaSemana(@DataMov)))
                    
   SELECT @Erro = @@ERROR

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

GO

--****************************************************
CREATE PROCEDURE upINCLUI_VENDA_INGR
   (@DataMov datetime,
    @Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS

   DECLARE @DataIniPer     datetime,
           @DataFimPer     datetime,
           @HoraMaxSes     datetime
           
   SELECT @HoraMaxSes = par_hora_max_ses
   FROM tb_parametro

   SELECT @DataIniPer = convert(datetime, convert(char(10),@DataMov,103) + ' ' + convert(char(10),@HoraMaxSes,108), 103)
   SELECT @DataFimPer = convert(datetime, convert(char(10),DATEADD(Day, 1, @DataMov),103) + ' ' + convert(char(10),@HoraMaxSes,108), 103)

   INSERT INTO tb_bol_ingre
       (bol_dt_mov,
        emp_cd,
        cin_cd,
        sal_cd,
        fil_cd,
        sre_data,
        sre_horario,
        igt_cd,
        ing_status,
        ing_dt_venda,
        opt_cd,
        pgt_cd,
        bin_dev,
        bin_qtde,
        ing_valor,
        cxp_talao,
        ppr_cd)
   SELECT @DataMov,
          tb_cinema.emp_cd,
          tb_cinema.cin_cd,
          tb_venda_ingresso.sal_cd,
          tb_venda_ingresso.fil_cd,
          tb_venda_ingresso.sre_data,
          tb_venda_ingresso.sre_horario,
          tb_venda_ingresso.igt_cd,
          tb_venda_ingresso.ing_status,
          tb_venda_ingresso.ing_dt_venda,
          tb_operacao.opt_cd,
          tb_pagamento.pgt_cd,
          CASE
             WHEN tb_venda_ingresso.ing_dt_canc IS NOT NULL THEN 'S'
             ELSE 'N'
          END                              AS Devolvido,
          COUNT(*)                         AS bin_qtde,
          tb_venda_ingresso.ing_valor      AS ing_valor,
          tb_caixa_movto.cxp_talao,
          ISNULL(tb_venda_ingresso.ppr_cd, 0)
   FROM tb_venda_ingresso,
        tb_operacao,
        tb_pagamento,
        tb_caixa,
        tb_cinema,
        tb_caixa_movto
   WHERE tb_venda_ingresso.ope_cd    = tb_operacao.ope_cd
   AND   tb_operacao.ope_cd          = tb_pagamento.ope_cd
   AND   tb_operacao.cxa_cd          = tb_caixa.cxa_cd
   AND   tb_caixa.cin_cd             = tb_cinema.cin_cd
   AND   tb_operacao.cxp_cd          = tb_caixa_movto.cxp_cd
   AND   (tb_venda_ingresso.sre_data = @DataMov
   OR     tb_venda_ingresso.ing_dt_venda between @DataIniPer AND @DataFimPer)
   AND   (tb_venda_ingresso.ing_dt_canc > @DataMov
   OR     tb_venda_ingresso.ing_dt_canc IS NULL)
   GROUP BY tb_cinema.emp_cd,
            tb_cinema.cin_cd,
            tb_venda_ingresso.sal_cd,
            tb_venda_ingresso.fil_cd,
            tb_venda_ingresso.sre_data,
            tb_venda_ingresso.sre_horario,
            tb_venda_ingresso.igt_cd,
            tb_venda_ingresso.ing_status,
            tb_venda_ingresso.ing_dt_venda,
            tb_operacao.opt_cd,
            tb_pagamento.pgt_cd,
            CASE
               WHEN tb_venda_ingresso.ing_dt_canc IS NOT NULL THEN 'S'
               ELSE 'N'
            END,
            tb_caixa_movto.cxp_talao,
            tb_venda_ingresso.ppr_cd,
            tb_venda_ingresso.ing_valor
            
   SELECT @Erro = @@ERROR

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

GO

--****************************************************
CREATE PROCEDURE upINCLUI_BOL_TALAO
   (@DataMov   datetime,
    @Erro      int OUTPUT,
    @MsgErr    varchar(255) OUTPUT)
AS

   DECLARE @DataIniPer     datetime,
           @DataFimPer     datetime,
           @HoraMaxSes     datetime
           
   SELECT @HoraMaxSes = par_hora_max_ses
   FROM tb_parametro

   SELECT @DataIniPer = convert(datetime, convert(char(10),@DataMov,103) + ' ' + convert(char(10),@HoraMaxSes,108), 103)
   SELECT @DataFimPer = convert(datetime, convert(char(10),DATEADD(Day, 1, @DataMov),103) + ' ' + convert(char(10),@HoraMaxSes,108), 103)

   INSERT INTO tb_bol_talao
         (bol_dt_mov,
          emp_cd,
          cin_cd,
          sal_cd,
          fil_cd,
          igt_cd,
          num_talao_ini,
          num_talao_fim)
   SELECT tb_venda_ingresso.sre_data,
          tb_cinema.emp_cd,
          tb_cinema.cin_cd,
          tb_venda_ingresso.sal_cd,
          tb_venda_ingresso.fil_cd,
          tb_venda_ingresso.igt_cd,
          MIN(tb_venda_ingresso.ing_num_talao), 
          MAX(tb_venda_ingresso.ing_num_talao)
   FROM tb_operacao,
        tb_venda_ingresso,
        tb_ingresso_tipo,
        tb_caixa,
        tb_cinema
   WHERE tb_operacao.ope_cd       = tb_venda_ingresso.ope_cd
   AND   tb_venda_ingresso.igt_cd = tb_ingresso_tipo.igt_cd
   AND   tb_operacao.cxa_cd       = tb_caixa.cxa_cd
   AND   tb_caixa.cin_cd          = tb_cinema.cin_cd
   AND   tb_operacao.ope_dt_operacao between @DataIniPer AND @DataFimPer
   AND   tb_operacao.ope_dt_des IS NULL
   AND   tb_venda_ingresso.ing_dt_canc IS NULL
   AND   tb_venda_ingresso.ing_num_talao IS NOT NULL
   GROUP BY tb_venda_ingresso.sre_data,
            tb_cinema.emp_cd,
            tb_cinema.cin_cd,
            tb_venda_ingresso.sal_cd,
            tb_venda_ingresso.fil_cd,
            tb_venda_ingresso.igt_cd

   SELECT @Erro = @@ERROR
    
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      End

GO

--****************************************************
CREATE PROCEDURE upINCLUI_BOL_SESSAO
   (@DataMov   datetime,
    @Erro      int OUTPUT,
    @MsgErr    varchar(255) OUTPUT)
AS

--   SET DATEFIRST 1
   
   DECLARE @diaSemana   smallint, 
           @dt_abertura datetime,
           @DataIniPer  datetime,
           @DataFimPer  datetime,
           @HoraMaxSes  datetime

   SELECT @diaSemana = dbo.ufDiaSemana(@DataMov)

   SELECT @HoraMaxSes = par_hora_max_ses
   FROM tb_parametro

   SELECT @DataIniPer = convert(datetime, convert(char(10),@DataMov,103) + ' ' + convert(char(10),@HoraMaxSes,108), 103)
   SELECT @DataFimPer = convert(datetime, convert(char(10),DATEADD(Day, 1, @DataMov),103) + ' ' + convert(char(10),@HoraMaxSes,108), 103)

   INSERT INTO tb_bol_sessao
         (bol_dt_mov,
          emp_cd,
          cin_cd,
          sal_cd,
          fil_cd,
          sre_data,
          ses_horario,
          ses_pre_estreia)
   SELECT distinct 
          @DataMov,
          tb_cinema.emp_cd,
          tb_cinema.cin_cd,
          tb_sessao.sal_cd,
          tb_sessao.fil_cd,
          @DataMov,
          tb_sessao.ses_horario,
          tb_sessao.ses_pre_estreia
      FROM tb_sessao,
           tb_programacao,
           tb_sala,
           tb_cinema,
           tb_filme
   WHERE tb_sessao.prg_cd = tb_programacao.prg_cd
   AND   tb_sessao.sal_cd = tb_sala.sal_cd
   AND   tb_sala.cin_cd   = tb_cinema.cin_cd
   AND   tb_sessao.fil_cd = tb_filme.fil_cd
   AND   @DataMov between tb_programacao.prg_dt_ini AND tb_programacao.prg_dt_fim
   AND   tb_programacao.prg_dt_des IS NULL
   AND   @DataMov between tb_filme.fil_dt_ini AND tb_filme.fil_dt_fim
   AND   tb_sessao.ses_dia_semana = @diaSemana
   UNION
   SELECT distinct 
          @DataMov,
          tb_cinema.emp_cd,
          tb_cinema.cin_cd,
          tb_sessao.sal_cd,
          tb_sessao.fil_cd,
          tb_venda_ingresso.sre_data,
          tb_sessao.ses_horario,
          tb_sessao.ses_pre_estreia
      FROM tb_sessao,
           tb_programacao,
           tb_sala,
           tb_cinema,
           tb_venda_ingresso
   WHERE tb_sessao.prg_cd = tb_programacao.prg_cd
   AND   tb_sessao.sal_cd = tb_sala.sal_cd
   AND   tb_sala.cin_cd   = tb_cinema.cin_cd
   AND   tb_sessao.sal_cd = tb_venda_ingresso.sal_cd
   AND   tb_sessao.fil_cd = tb_venda_ingresso.fil_cd
   AND   tb_venda_ingresso.sre_data between tb_programacao.prg_dt_ini AND tb_programacao.prg_dt_fim
   AND   tb_programacao.prg_dt_des IS NULL
   AND   tb_sessao.ses_dia_semana       = dbo.ufDiaSemana(tb_venda_ingresso.sre_data)
   AND   tb_venda_ingresso.ing_dt_venda between @DataIniPer AND @DataFimPer

   SELECT @Erro = @@ERROR
    
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      End

GO

--****************************************************
CREATE PROCEDURE upBOLETIM2_GERA
   (@DataMov datetime,
    @Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS
   DECLARE @totIngre int

   DELETE FROM tb_bol_talao
   WHERE bol_dt_mov = @DataMov

   DELETE FROM tb_bol_sessao
   WHERE bol_dt_mov = @DataMov

   DELETE FROM tb_bol_ingre
   WHERE bol_dt_mov = @DataMov

   DELETE FROM tb_ctrl_boletim 
   WHERE bol_dt_mov = @DataMov

   DELETE FROM tb_boletim 
   WHERE bol_dt_mov = @DataMov
   
   DELETE FROM tb_bol_filme
   WHERE tb_bol_filme.bol_dt_mov = @DataMov
   

   -- Inclui parâmetros
   execute upINCLUI_PARAM @DataMov, @Erro, @MsgErr

   IF @Erro <> 0
      Return

   -- verifica tipo ingresso 
   execute upVER_TIPO_INGRESSO @Erro, @MsgErr

   IF @Erro <> 0
      Return

   -- verifica tipo pagamento
   execute upVER_TIPO_PAGAMENTO @Erro, @MsgErr

   IF @Erro <> 0
      Return

   -- verifica tipo operação
   execute upVER_TIPO_OPERACAO @Erro, @MsgErr

   IF @Erro <> 0
      Return

   -- Inclui Empresa
   execute upINCLUI_EMPRESA @DataMov, @Erro, @MsgErr

   IF @Erro <> 0
      Return

   -- Inclui Cinema
   execute upINCLUI_CINEMA @DataMov, @Erro, @MsgErr

   IF @Erro <> 0
      Return
   
   -- Inclui Catraca
   execute upINCLUI_CATRACA @DataMov, @Erro, @MsgErr

   IF @Erro <> 0
      Return

   -- Inclui Sala
   execute upINCLUI_SALA @DataMov, @Erro, @MsgErr

   IF @Erro <> 0
      Return

   -- Inclui catraca_sala
   execute upINCLUI_CATRACA_SALA @DataMov, @Erro, @MsgErr

   IF @Erro <> 0
      Return
      

   -- Inclui Distribuidora
   execute upINCLUI_DISTRIBUIDORA @DataMov, @Erro, @MsgErr

   IF @Erro <> 0
      Return

   -- Inclui Filme
   execute upINCLUI_FILME @DataMov, @Erro, @MsgErr

   IF @Erro <> 0
      Return
   
   -- Inclui Boletim
   execute upINCLUI_BOLETIM @DataMov, @Erro, @MsgErr

   IF @Erro <> 0
      Return

   -- Inclui Venda Ingresso
   execute upINCLUI_VENDA_INGR @DataMov, @Erro, @MsgErr
   
   IF @Erro <> 0
      Return
   
   execute upINCLUI_BOL_TALAO @DataMov, @Erro, @MsgErr

   IF @Erro <> 0
      Return
   
   execute upINCLUI_BOL_SESSAO @DataMov, @Erro, @MsgErr

   IF @Erro <> 0
      Return

GO


--****************************************************
--****************************************************
--****************************************************


--****************************************************
CREATE PROCEDURE upBOLETIM_CAPA2
   (@Data      datetime,
    @sal_cd    int,
    @fil_cd    int,
    @Erro      int OUTPUT,
    @MsgErr    varchar(255) OUTPUT)
AS
--   SET DATEFIRST 1
   DECLARE @ses_pre_estreia VARCHAR(1)

   --SELECT @ses_pre_estreia = ses_pre_estreia
   SELECT DISTINCT @ses_pre_estreia = tb_bol_sessao.ses_pre_estreia
   FROM tb_bol_sessao
   WHERE convert(char(10),tb_bol_sessao.bol_dt_mov,102) = convert(char(10),@Data,102)
   AND   tb_bol_sessao.fil_cd     = @fil_cd
   AND   tb_bol_sessao.sal_cd     = @sal_cd

   SELECT DISTINCT tb_bol_empr.emp_nm, 
                   tb_bol_cin.cin_nm, 
                   tb_bol_sala.sal_desc, 
                   tb_bol_sala.sal_lugares, 
                   tb_boletim.bol_dt_abertura 'dt_abertura', 
                   getdate() 'dt_atual',
                   tb_bol_filme.fil_nm, 
                   tb_bol_distrib.dis_nm 'fil_distribuidora', 
                   tb_boletim.bol_dt_ini_per  'prg_dt_ini', 
                   tb_boletim.bol_dt_fim_per 'prg_dt_fim', 
                   tb_bol_filme.fil_cd,
                   @ses_pre_estreia 'pre_estreia'
   FROM tb_boletim,
        tb_bol_filme,
        tb_bol_sala,
        tb_bol_cin,
        tb_bol_empr,
        tb_bol_distrib
   WHERE tb_boletim.fil_cd       = tb_bol_filme.fil_cd
   AND   tb_boletim.bol_dt_mov   = tb_bol_filme.bol_dt_mov
   AND   tb_bol_filme.dis_cd     = tb_bol_distrib.dis_cd
   AND   tb_bol_filme.bol_dt_mov = tb_bol_distrib.bol_dt_mov
   AND   tb_boletim.sal_cd       = tb_bol_sala.sal_cd
   AND   tb_boletim.bol_dt_mov   = tb_bol_sala.bol_dt_mov
   AND   tb_boletim.cin_cd       = tb_bol_cin.cin_cd
   AND   tb_boletim.bol_dt_mov   = tb_bol_cin.bol_dt_mov
   AND   tb_boletim.emp_cd       = tb_bol_empr.emp_cd
   AND   tb_boletim.bol_dt_mov   = tb_bol_empr.bol_dt_mov
   AND   tb_boletim.sal_cd       = @sal_cd
   AND   tb_boletim.fil_cd       = @fil_cd
   AND   convert(char(10),tb_boletim.bol_dt_mov,102) = convert(char(10),@Data,102)
          
   SELECT @Erro = @@ERROR
    
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      End

GO

--****************************************************
CREATE PROCEDURE upBOLETIM_SESSOES_FILME2
   (@Data      datetime,
    @sal_cd    int,
    @fil_cd    int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
--   SET DATEFIRST 1
    
   SELECT distinct tb_bol_ingre.sre_horario
   FROM tb_bol_ingre
   WHERE tb_bol_ingre.sal_cd = @sal_cd
   AND   tb_bol_ingre.fil_cd = @fil_cd
   AND   convert(char(10),tb_bol_ingre.bol_dt_mov,102) = convert(char(10),@Data,102)   
   ORDER BY tb_bol_ingre.sre_horario
    
   SELECT @Erro = @@ERROR
    
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      End

GO

--****************************************************
CREATE PROCEDURE upBOLETIM_VENDAS_DIA2
   (@Data      datetime,
    @sal_cd    int,
    @fil_cd    int,
    @Erro      int OUTPUT,
    @MsgErr    varchar(255) OUTPUT)
AS
    
   DECLARE @HoraLimite  char(5),
           @HoraMaxSes  datetime,
           @DataIniPer  datetime,
           @DataFimPer  datetime


   SELECT @HoraMaxSes = par_hora_max_ses
   FROM tb_bol_param
   WHERE tb_bol_param. bol_dt_mov = @Data

   SELECT @DataIniPer = convert(datetime, convert(char(10),@Data,103) + ' ' + convert(char(10),@HoraMaxSes,108), 103)
   SELECT @DataFimPer = convert(datetime, convert(char(10),DATEADD(Day, 1, @Data),103) + ' ' + convert(char(10),@HoraMaxSes,108), 103)

   SELECT 0                               'ordem', 
          tb_bol_ingre.igt_cd,  
          SUM(tb_bol_ingre.bin_qtde)      'qtde', 
          tb_bol_ingre.ing_valor          'ing_valor',
          tb_bol_tp_ingr.igt_desc         'igt_desc', 
          tb_bol_ingre.ppr_cd
   FROM tb_bol_ingre,
        tb_bol_tp_ingr
   WHERE tb_bol_ingre.igt_cd = tb_bol_tp_ingr.igt_cd
   AND   tb_bol_ingre.sal_cd = @sal_cd
   AND   tb_bol_ingre.fil_cd = @fil_cd
   AND   tb_bol_ingre.bol_dt_mov = @Data
   AND   tb_bol_ingre.sre_data   = @Data
   AND   tb_bol_ingre.ing_dt_venda between @DataIniPer AND @DataFimPer
   AND   tb_bol_ingre.opt_cd IN ( 1, 3 )     -- Operação Venda Normal + devolução
   AND   tb_bol_ingre.igt_cd IN ( 1, 2 )     -- Inteira e Meia
   AND   tb_bol_ingre.cxp_talao = 0          -- não talão
   GROUP BY tb_bol_ingre.igt_cd, 
            tb_bol_tp_ingr.igt_desc, 
            tb_bol_ingre.ppr_cd,
            tb_bol_ingre.ing_valor

   UNION

   SELECT 2                               'ordem', 
          tb_bol_ingre.igt_cd, 
          SUM(tb_bol_ingre.bin_qtde)      'qtde', 
          tb_bol_ingre.ing_valor          'ing_valor',
          tb_bol_tp_ingr.igt_desc         'igt_desc', 
          tb_bol_ingre.ppr_cd
   FROM tb_bol_ingre,
        tb_bol_tp_ingr
   WHERE tb_bol_ingre.igt_cd = tb_bol_tp_ingr.igt_cd
   AND   tb_bol_ingre.sal_cd = @sal_cd
   AND   tb_bol_ingre.fil_cd = @fil_cd
   AND   tb_bol_ingre.bol_dt_mov = @Data
   AND   tb_bol_ingre.sre_data   = @Data
   AND   tb_bol_ingre.ing_dt_venda between @DataIniPer AND @DataFimPer
   AND   tb_bol_ingre.opt_cd IN ( 1, 3 )  -- Operação Venda Normal + devolução
   AND   tb_bol_ingre.igt_cd IN ( 3, 4 )  -- Inteira e Meia Promoção
   AND   tb_bol_ingre.cxp_talao = 0       -- não talão
   GROUP BY tb_bol_ingre.igt_cd, 
            tb_bol_tp_ingr.igt_desc, 
            tb_bol_ingre.ppr_cd,
            tb_bol_ingre.ing_valor

   UNION

   SELECT 4                                  'ordem', 
          tb_bol_ingre.igt_cd, 
          SUM(tb_bol_ingre.bin_qtde)         'qtde', 
          tb_bol_ingre.ing_valor             'ing_valor',
          tb_bol_tp_ingr.igt_desc + ' TA'    'igt_desc', 
          tb_bol_ingre.ppr_cd
   FROM tb_bol_ingre,
        tb_bol_tp_ingr
   WHERE tb_bol_ingre.igt_cd = tb_bol_tp_ingr.igt_cd
   AND   tb_bol_ingre.sal_cd = @sal_cd
   AND   tb_bol_ingre.fil_cd = @fil_cd
   AND   tb_bol_ingre.bol_dt_mov = @Data
   AND   tb_bol_ingre.sre_data   = @Data
   AND   tb_bol_ingre.ing_dt_venda between @DataIniPer AND @DataFimPer
   AND   tb_bol_ingre.opt_cd IN ( 1, 3 ) -- Operação Venda Normal + devolução
   AND   tb_bol_ingre.igt_cd IN ( 1, 2 ) -- Inteira e Meia
   AND   tb_bol_ingre.cxp_talao = 1      -- Talão
   GROUP BY tb_bol_ingre.igt_cd, 
            tb_bol_tp_ingr.igt_desc, 
            tb_bol_ingre.ppr_cd,
            tb_bol_ingre.ing_valor

   UNION

   SELECT 6                                  'ordem', 
          tb_bol_ingre.igt_cd, 
          SUM(tb_bol_ingre.bin_qtde)         'qtde', 
          tb_bol_ingre.ing_valor             'ing_valor',
          tb_bol_tp_ingr.igt_desc + ' TA'    'igt_desc', 
          tb_bol_ingre.ppr_cd
   FROM tb_bol_ingre,
        tb_bol_tp_ingr
   WHERE tb_bol_ingre.igt_cd = tb_bol_tp_ingr.igt_cd
   AND   tb_bol_ingre.sal_cd = @sal_cd
   AND   tb_bol_ingre.fil_cd = @fil_cd
   AND   tb_bol_ingre.bol_dt_mov = @Data
   AND   tb_bol_ingre.sre_data   = @Data
   AND   tb_bol_ingre.ing_dt_venda between @DataIniPer AND @DataFimPer
   AND   tb_bol_ingre.opt_cd IN ( 1, 3 ) -- Operação Venda Normal + devolução
   AND   tb_bol_ingre.igt_cd IN ( 3, 4 ) -- Inteira e Meia Promoção
   AND   tb_bol_ingre.cxp_talao = 1      -- Talão
   GROUP BY tb_bol_ingre.igt_cd, 
            tb_bol_tp_ingr.igt_desc, 
            tb_bol_ingre.ppr_cd,
            tb_bol_ingre.ing_valor

   ORDER BY 1

   SELECT @Erro = @@ERROR
    
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      End

GO

--****************************************************
CREATE PROCEDURE upBOLETIM_PRE_VENDA2
   (@Data      datetime,
    @sal_cd    int,
    @fil_cd    int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS

   DECLARE @HoraLimite  char(5),
           @HoraMaxSes     datetime,
           @DataIniPer     datetime,
           @DataFimPer     datetime

   SELECT @HoraMaxSes = par_hora_max_ses
   FROM tb_bol_param
   WHERE tb_bol_param. bol_dt_mov = @Data
     
   SELECT @DataIniPer = convert(datetime, convert(char(10),@Data,103) + ' ' + convert(char(10),@HoraMaxSes,108), 103)
   SELECT @DataFimPer = convert(datetime, convert(char(10),DATEADD(Day, 1, @Data),103) + ' ' + convert(char(10),@HoraMaxSes,108), 103)

   SELECT 0                               'ordem', 
          tb_bol_ingre.igt_cd,  
          SUM(tb_bol_ingre.bin_qtde)      'qtde', 
          tb_bol_ingre.ing_valor          'ing_valor',
          tb_bol_tp_ingr.igt_desc         'igt_desc', 
          tb_bol_ingre.ppr_cd
   FROM tb_bol_ingre,
        tb_bol_tp_ingr
   WHERE tb_bol_ingre.igt_cd = tb_bol_tp_ingr.igt_cd
   AND   tb_bol_ingre.sal_cd = @sal_cd
   AND   tb_bol_ingre.fil_cd = @fil_cd
   AND   tb_bol_ingre.bol_dt_mov = @Data
   AND   tb_bol_ingre.sre_data   = @Data
   AND   tb_bol_ingre.ing_dt_venda < @Data
   AND   tb_bol_ingre.opt_cd IN ( 1, 3 ) -- Operação Venda Normal + devolução
   AND   tb_bol_ingre.igt_cd IN ( 1, 2 ) -- Inteira e Meia
   AND   tb_bol_ingre.cxp_talao = 0      -- não talão
   GROUP BY tb_bol_ingre.igt_cd, 
            tb_bol_tp_ingr.igt_desc, 
            tb_bol_ingre.ppr_cd,
            tb_bol_ingre.ing_valor

   UNION

   SELECT 2                               'ordem', 
          tb_bol_ingre.igt_cd, 
          SUM(tb_bol_ingre.bin_qtde)      'qtde', 
          tb_bol_ingre.ing_valor          'ing_valor',
          tb_bol_tp_ingr.igt_desc         'igt_desc', 
          tb_bol_ingre.ppr_cd
   FROM tb_bol_ingre,
        tb_bol_tp_ingr
   WHERE tb_bol_ingre.igt_cd = tb_bol_tp_ingr.igt_cd
   AND   tb_bol_ingre.sal_cd = @sal_cd
   AND   tb_bol_ingre.fil_cd = @fil_cd
   AND   tb_bol_ingre.bol_dt_mov = @Data
   AND   tb_bol_ingre.sre_data   = @Data
   AND   tb_bol_ingre.ing_dt_venda < @Data
   AND   tb_bol_ingre.opt_cd IN ( 1, 3 ) -- Operação Venda Normal + devolução
   AND   tb_bol_ingre.igt_cd IN ( 3, 4 ) -- Inteira e Meia Promoção
   AND   tb_bol_ingre.cxp_talao = 0      -- não talão
   GROUP BY tb_bol_ingre.igt_cd, 
            tb_bol_tp_ingr.igt_desc, 
            tb_bol_ingre.ppr_cd,
            tb_bol_ingre.ing_valor

   UNION

   SELECT 4                                  'ordem', 
          tb_bol_ingre.igt_cd, 
          SUM(tb_bol_ingre.bin_qtde)         'qtde', 
          tb_bol_ingre.ing_valor             'ing_valor',
          tb_bol_tp_ingr.igt_desc + ' TA'    'igt_desc', 
          tb_bol_ingre.ppr_cd
   FROM tb_bol_ingre,
        tb_bol_tp_ingr
   WHERE tb_bol_ingre.igt_cd = tb_bol_tp_ingr.igt_cd
   AND   tb_bol_ingre.sal_cd = @sal_cd
   AND   tb_bol_ingre.fil_cd = @fil_cd
   AND   tb_bol_ingre.bol_dt_mov = @Data
   AND   tb_bol_ingre.sre_data   = @Data
   AND   tb_bol_ingre.ing_dt_venda < @Data
   AND   tb_bol_ingre.opt_cd IN ( 1, 3 ) -- Operação Venda Normal + devolução
   AND   tb_bol_ingre.igt_cd IN ( 1, 2 ) -- Inteira e Meia
   AND   tb_bol_ingre.cxp_talao = 1      -- Talão
   GROUP BY tb_bol_ingre.igt_cd, 
            tb_bol_tp_ingr.igt_desc, 
            tb_bol_ingre.ppr_cd,
            tb_bol_ingre.ing_valor

   UNION

   SELECT 6                                  'ordem', 
          tb_bol_ingre.igt_cd, 
          SUM(tb_bol_ingre.bin_qtde)         'qtde', 
          tb_bol_ingre.ing_valor             'ing_valor',
          tb_bol_tp_ingr.igt_desc + ' TA'    'igt_desc', 
          tb_bol_ingre.ppr_cd
   FROM tb_bol_ingre,
        tb_bol_tp_ingr
   WHERE tb_bol_ingre.igt_cd = tb_bol_tp_ingr.igt_cd
   AND   tb_bol_ingre.sal_cd = @sal_cd
   AND   tb_bol_ingre.fil_cd = @fil_cd
   AND   tb_bol_ingre.bol_dt_mov = @Data
   AND   tb_bol_ingre.sre_data   = @Data
   AND   tb_bol_ingre.ing_dt_venda < @Data
   AND   tb_bol_ingre.opt_cd IN ( 1, 3 ) -- Operação Venda Normal + devolução
   AND   tb_bol_ingre.igt_cd IN ( 3, 4 ) -- Inteira e Meia Promoção
   AND   tb_bol_ingre.cxp_talao = 1      -- Talão
   GROUP BY tb_bol_ingre.igt_cd, 
            tb_bol_tp_ingr.igt_desc, 
            tb_bol_ingre.ppr_cd,
            tb_bol_ingre.ing_valor

   ORDER BY 1

    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          Return
       End

GO

--****************************************************
CREATE PROCEDURE upBOLETIM_VENDAS_TOTAL2
   (@Data      datetime,
    @sal_cd    int,
    @fil_cd    int,
    @Erro      int OUTPUT,
    @MsgErr    varchar(255) OUTPUT)
AS
    
   DECLARE @HoraLimite  char(5),
           @HoraMaxSes  datetime,
           @DataIniPer  datetime,
           @DataFimPer  datetime


   SELECT @HoraMaxSes = par_hora_max_ses
   FROM tb_bol_param
   WHERE tb_bol_param. bol_dt_mov = @Data

   SELECT 0                               'ordem', 
          tb_bol_ingre.igt_cd,  
          SUM(tb_bol_ingre.bin_qtde)      'qtde', 
          tb_bol_ingre.ing_valor          'ing_valor',
          tb_bol_tp_ingr.igt_desc         'igt_desc', 
          tb_bol_ingre.ppr_cd
   FROM tb_bol_ingre,
        tb_bol_tp_ingr
   WHERE tb_bol_ingre.igt_cd = tb_bol_tp_ingr.igt_cd
   AND   tb_bol_ingre.sal_cd = @sal_cd
   AND   tb_bol_ingre.fil_cd = @fil_cd
   AND   tb_bol_ingre.bol_dt_mov = @Data
   AND   tb_bol_ingre.sre_data   = @Data
   AND   tb_bol_ingre.opt_cd IN ( 1, 3 )     -- Operação Venda Normal + devolução
   AND   tb_bol_ingre.igt_cd IN ( 1, 2 )     -- Inteira e Meia
   AND   tb_bol_ingre.cxp_talao = 0          -- não talão
   GROUP BY tb_bol_ingre.igt_cd, 
            tb_bol_tp_ingr.igt_desc, 
            tb_bol_ingre.ppr_cd,
            tb_bol_ingre.ing_valor

   UNION

   SELECT 2                               'ordem', 
          tb_bol_ingre.igt_cd, 
          SUM(tb_bol_ingre.bin_qtde)      'qtde', 
          tb_bol_ingre.ing_valor          'ing_valor',
          tb_bol_tp_ingr.igt_desc         'igt_desc', 
          tb_bol_ingre.ppr_cd
   FROM tb_bol_ingre,
        tb_bol_tp_ingr
   WHERE tb_bol_ingre.igt_cd = tb_bol_tp_ingr.igt_cd
   AND   tb_bol_ingre.sal_cd = @sal_cd
   AND   tb_bol_ingre.fil_cd = @fil_cd
   AND   tb_bol_ingre.bol_dt_mov = @Data
   AND   tb_bol_ingre.sre_data   = @Data
   AND   tb_bol_ingre.opt_cd IN ( 1, 3 )  -- Operação Venda Normal + devolução
   AND   tb_bol_ingre.igt_cd IN ( 3, 4 )  -- Inteira e Meia Promoção
   AND   tb_bol_ingre.cxp_talao = 0       -- não talão
   GROUP BY tb_bol_ingre.igt_cd, 
            tb_bol_tp_ingr.igt_desc, 
            tb_bol_ingre.ppr_cd,
            tb_bol_ingre.ing_valor

   UNION

   SELECT 4                                  'ordem', 
          tb_bol_ingre.igt_cd, 
          SUM(tb_bol_ingre.bin_qtde)         'qtde', 
          tb_bol_ingre.ing_valor             'ing_valor',
          tb_bol_tp_ingr.igt_desc + ' TA'    'igt_desc', 
          tb_bol_ingre.ppr_cd
   FROM tb_bol_ingre,
        tb_bol_tp_ingr
   WHERE tb_bol_ingre.igt_cd = tb_bol_tp_ingr.igt_cd
   AND   tb_bol_ingre.sal_cd = @sal_cd
   AND   tb_bol_ingre.fil_cd = @fil_cd
   AND   tb_bol_ingre.bol_dt_mov = @Data
   AND   tb_bol_ingre.sre_data   = @Data
   AND   tb_bol_ingre.opt_cd IN ( 1, 3 ) -- Operação Venda Normal + devolução
   AND   tb_bol_ingre.igt_cd IN ( 1, 2 ) -- Inteira e Meia
   AND   tb_bol_ingre.cxp_talao = 1      -- Talão
   GROUP BY tb_bol_ingre.igt_cd, 
            tb_bol_tp_ingr.igt_desc, 
            tb_bol_ingre.ppr_cd,
            tb_bol_ingre.ing_valor

   UNION

   SELECT 6                                  'ordem', 
          tb_bol_ingre.igt_cd, 
          SUM(tb_bol_ingre.bin_qtde)         'qtde', 
          tb_bol_ingre.ing_valor             'ing_valor',
          tb_bol_tp_ingr.igt_desc + ' TA'    'igt_desc', 
          tb_bol_ingre.ppr_cd
   FROM tb_bol_ingre,
        tb_bol_tp_ingr
   WHERE tb_bol_ingre.igt_cd = tb_bol_tp_ingr.igt_cd
   AND   tb_bol_ingre.sal_cd = @sal_cd
   AND   tb_bol_ingre.fil_cd = @fil_cd
   AND   tb_bol_ingre.bol_dt_mov = @Data
   AND   tb_bol_ingre.sre_data   = @Data
   AND   tb_bol_ingre.opt_cd IN ( 1, 3 ) -- Operação Venda Normal + devolução
   AND   tb_bol_ingre.igt_cd IN ( 3, 4 ) -- Inteira e Meia Promoção
   AND   tb_bol_ingre.cxp_talao = 1      -- Talão
   GROUP BY tb_bol_ingre.igt_cd, 
            tb_bol_tp_ingr.igt_desc, 
            tb_bol_ingre.ppr_cd,
            tb_bol_ingre.ing_valor

   ORDER BY 1

   SELECT @Erro = @@ERROR
    
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      End

GO

--****************************************************
CREATE PROCEDURE upBOLETIM_CORTESIA2
   (@Data      datetime,
    @sal_cd    int,
    @fil_cd    int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   DECLARE @HoraLimite  char(5),
           @HoraMaxSes     datetime,
           @DataIniPer     datetime,
           @DataFimPer     datetime

   SELECT @HoraLimite = convert(char(5),par_hora_limite,108),
          @HoraMaxSes = par_hora_max_ses
   FROM tb_bol_param
   WHERE tb_bol_param. bol_dt_mov = @Data
     
   SELECT @DataIniPer = convert(datetime, convert(char(10),@Data,103) + ' ' + convert(char(10),@HoraMaxSes,108), 103)
   SELECT @DataFimPer = convert(datetime, convert(char(10),DATEADD(Day, 1, @Data),103) + ' ' + convert(char(10),@HoraMaxSes,108), 103)

   SELECT y.igt_cd, 
          ISNULL(x.qtde,0) 'qtde', 
          y.igt_desc , 
          x.desc_dia
   FROM (
          SELECT tb_bol_ingre.igt_cd, 
                 SUM(tb_bol_ingre.bin_qtde) 'qtde', 
                 ' '                        'desc_dia'
          FROM tb_bol_ingre
          WHERE tb_bol_ingre.sal_cd = @sal_cd
          AND   tb_bol_ingre.fil_cd = @fil_cd
          AND   tb_bol_ingre.bol_dt_mov = @Data
          AND   tb_bol_ingre.sre_data   = @Data
          AND   tb_bol_ingre.ing_dt_venda between @DataIniPer AND @DataFimPer
          AND   tb_bol_ingre.opt_cd  IN ( 1, 3 )  -- Operação Venda Normal + devolução
          AND   tb_bol_ingre.igt_cd  = 9          -- Cortesia
          AND   tb_bol_ingre.bin_dev = 'N'
          GROUP BY tb_bol_ingre.igt_cd
          
          UNION
          
          SELECT tb_bol_ingre.igt_cd, 
                 SUM(tb_bol_ingre.bin_qtde) 'qtde', 
                 ' PRE'                     'desc_dia'
          FROM tb_bol_ingre
          WHERE tb_bol_ingre.sal_cd = @sal_cd
          AND   tb_bol_ingre.fil_cd = @fil_cd
          AND   tb_bol_ingre.bol_dt_mov   = @Data
          AND   tb_bol_ingre.sre_data     = @Data
          AND   tb_bol_ingre.ing_dt_venda < @Data
          AND   tb_bol_ingre.opt_cd  IN ( 1, 3 )  -- Operação Venda Normal + devolução
          AND   tb_bol_ingre.igt_cd  = 9          -- Cortesia
          AND   tb_bol_ingre.bin_dev = 'N'
          GROUP BY tb_bol_ingre.igt_cd
        ) x,
   tb_bol_tp_ingr y
   WHERE x.igt_cd = y.igt_cd
   AND   y.igt_cd = 9      -- Cortesia
    ORDER BY 1

   SELECT @Erro = @@ERROR
    
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      End

GO

--****************************************************
CREATE PROCEDURE upBOLETIM_DEVOLUCAO2
   (@Data      datetime,
    @sal_cd    int,
    @fil_cd    int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS

   DECLARE @HoraLimite  char(5),
           @HoraMaxSes     datetime,
           @DataIniPer     datetime,
           @DataFimPer     datetime

   SELECT @HoraMaxSes = par_hora_max_ses
   FROM tb_bol_param
   WHERE tb_bol_param. bol_dt_mov = @Data
     
   SELECT 0                               'ordem', 
          tb_bol_ingre.igt_cd,  
          SUM(tb_bol_ingre.bin_qtde)      'qtde', 
          tb_bol_ingre.ing_valor          'ing_valor',
          tb_bol_tp_ingr.igt_desc         'igt_desc', 
          tb_bol_ingre.ppr_cd
   FROM tb_bol_ingre,
        tb_bol_tp_ingr
   WHERE tb_bol_ingre.igt_cd = tb_bol_tp_ingr.igt_cd
   AND   tb_bol_ingre.sal_cd = @sal_cd
   AND   tb_bol_ingre.fil_cd = @fil_cd
   AND   tb_bol_ingre.bol_dt_mov = @Data
   AND   tb_bol_ingre.sre_data   = @Data
   AND   tb_bol_ingre.opt_cd IN ( 1, 3 )     -- Operação Venda Normal + devolução
   AND   tb_bol_ingre.igt_cd IN ( 1, 2 )     -- Inteira e Meia
   AND   tb_bol_ingre.cxp_talao = 0          -- não talão
   AND   tb_bol_ingre.bin_dev = 'S'
   GROUP BY tb_bol_ingre.igt_cd, 
            tb_bol_tp_ingr.igt_desc, 
            tb_bol_ingre.ppr_cd,
            tb_bol_ingre.ing_valor

   UNION

   SELECT 2                               'ordem', 
          tb_bol_ingre.igt_cd, 
          SUM(tb_bol_ingre.bin_qtde)      'qtde', 
          tb_bol_ingre.ing_valor          'ing_valor',
          tb_bol_tp_ingr.igt_desc         'igt_desc', 
          tb_bol_ingre.ppr_cd
   FROM tb_bol_ingre,
        tb_bol_tp_ingr
   WHERE tb_bol_ingre.igt_cd = tb_bol_tp_ingr.igt_cd
   AND   tb_bol_ingre.sal_cd = @sal_cd
   AND   tb_bol_ingre.fil_cd = @fil_cd
   AND   tb_bol_ingre.bol_dt_mov = @Data
   AND   tb_bol_ingre.sre_data   = @Data
   AND   tb_bol_ingre.opt_cd IN ( 1, 3 )     -- Operação Venda Normal + devolução
   AND   tb_bol_ingre.igt_cd IN ( 3, 4 )     -- Inteira e Meia Promoção
   AND   tb_bol_ingre.cxp_talao = 0          -- não talão
   AND   tb_bol_ingre.bin_dev = 'S'
   GROUP BY tb_bol_ingre.igt_cd, 
            tb_bol_tp_ingr.igt_desc, 
            tb_bol_ingre.ppr_cd,
            tb_bol_ingre.ing_valor

   UNION

   SELECT 4                                  'ordem', 
          tb_bol_ingre.igt_cd, 
          SUM(tb_bol_ingre.bin_qtde)         'qtde', 
          tb_bol_ingre.ing_valor             'ing_valor',
          tb_bol_tp_ingr.igt_desc + ' TA'    'igt_desc', 
          tb_bol_ingre.ppr_cd
   FROM tb_bol_ingre,
        tb_bol_tp_ingr
   WHERE tb_bol_ingre.igt_cd = tb_bol_tp_ingr.igt_cd
   AND   tb_bol_ingre.sal_cd = @sal_cd
   AND   tb_bol_ingre.fil_cd = @fil_cd
   AND   tb_bol_ingre.bol_dt_mov = @Data
   AND   tb_bol_ingre.sre_data   = @Data
   AND   tb_bol_ingre.opt_cd IN ( 1, 3 )     -- Operação Venda Normal + devolução
   AND   tb_bol_ingre.igt_cd IN ( 1, 2 )     -- Inteira e Meia
   AND   tb_bol_ingre.cxp_talao = 1            -- Talão
   AND   tb_bol_ingre.bin_dev = 'S'
   GROUP BY tb_bol_ingre.igt_cd, 
            tb_bol_tp_ingr.igt_desc, 
            tb_bol_ingre.ppr_cd,
            tb_bol_ingre.ing_valor

   UNION

   SELECT 6                                  'ordem', 
          tb_bol_ingre.igt_cd, 
          SUM(tb_bol_ingre.bin_qtde)         'qtde', 
          tb_bol_ingre.ing_valor             'ing_valor',
          tb_bol_tp_ingr.igt_desc + ' TA'    'igt_desc', 
          tb_bol_ingre.ppr_cd
   FROM tb_bol_ingre,
        tb_bol_tp_ingr
   WHERE tb_bol_ingre.igt_cd = tb_bol_tp_ingr.igt_cd
   AND   tb_bol_ingre.sal_cd = @sal_cd
   AND   tb_bol_ingre.fil_cd = @fil_cd
   AND   tb_bol_ingre.bol_dt_mov = @Data
   AND   tb_bol_ingre.sre_data   = @Data
   AND   tb_bol_ingre.opt_cd IN ( 1, 3 )     -- Operação Venda Normal + devolução
   AND   tb_bol_ingre.igt_cd IN ( 3, 4 )     -- Inteira e Meia Promoção
   AND   tb_bol_ingre.cxp_talao = 1          -- Talão
   AND   tb_bol_ingre.bin_dev = 'S'
   GROUP BY tb_bol_ingre.igt_cd, 
            tb_bol_tp_ingr.igt_desc, 
            tb_bol_ingre.ppr_cd,
            tb_bol_ingre.ing_valor

   --UNION

   --SELECT 8                               'ordem', 
   --       tb_bol_ingre.igt_cd, 
   --       SUM(tb_bol_ingre.bin_qtde)      'qtde', 
   --       tb_bol_ingre.ing_valor          'ing_valor',
   --       tb_bol_tp_ingr.igt_desc         'igt_desc', 
   --       tb_bol_ingre.ppr_cd
   --FROM tb_bol_ingre,
   --     tb_bol_tp_ingr
   --WHERE tb_bol_ingre.igt_cd = tb_bol_tp_ingr.igt_cd
   --AND   tb_bol_ingre.sal_cd = @sal_cd
   --AND   tb_bol_ingre.fil_cd = @fil_cd
   --AND   tb_bol_ingre.bol_dt_mov = @Data
   --AND   tb_bol_ingre.sre_data   = @Data
   --AND   tb_bol_ingre.opt_cd IN ( 1, 3 )     -- Operação Venda Normal + devolução
   --AND   tb_bol_ingre.igt_cd IN ( 9 )        -- Cortesia
   --AND   tb_bol_ingre.cxp_talao = 0          -- não talão
   --AND   tb_bol_ingre.bin_dev = 'S'
   --GROUP BY tb_bol_ingre.igt_cd, 
   --         tb_bol_tp_ingr.igt_desc, 
   --         tb_bol_ingre.ppr_cd,
   --         tb_bol_ingre.ing_valor

   ORDER BY 1

    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          Return
       End

GO

--****************************************************
CREATE PROCEDURE upBOLETIM_TOTAL_SESSAO2
   (@Data      datetime,
    @sal_cd    int,
    @fil_cd    int,
    @Erro      int OUTPUT,
    @MsgErr    varchar(255) OUTPUT)
AS

   
   SELECT z.ses_horario, 
          z.hora_excl,
          ISNULL(w.qtde ,0) qtde_int,
          ISNULL(x.qtde ,0) qtde_meia,
          ISNULL(y.qtde ,0) qtde_cort
   FROM (((SELECT tb_bol_sessao.ses_horario,
                  'N' hora_excl
           FROM tb_bol_sessao
           WHERE tb_bol_sessao.bol_dt_mov = @Data
           AND   tb_bol_sessao.sre_data   = @Data
           AND   tb_bol_sessao.sal_cd     = @sal_cd
           AND   tb_bol_sessao.fil_cd     = @fil_cd
           UNION
           SELECT DISTINCT 
                  tb_bol_ingre.sre_horario,
                  'S' hora_excl
           FROM tb_bol_ingre
           WHERE tb_bol_ingre.sre_data = @Data
           AND   tb_bol_ingre.sal_cd   = @sal_cd
           AND   tb_bol_ingre.fil_cd   = @fil_cd
           AND   tb_bol_ingre.sre_horario NOT IN (SELECT tb_bol_sessao.ses_horario
                                                  FROM tb_bol_sessao
                                                  WHERE tb_bol_sessao.bol_dt_mov = @Data
                                                  AND   tb_bol_sessao.sre_data   = @Data
                                                  AND   tb_bol_sessao.sal_cd     = @sal_cd
                                                  AND   tb_bol_sessao.fil_cd     = @fil_cd)
          ) z LEFT JOIN 
          (SELECT tb_bol_ingre.sre_horario, 
                  SUM(tb_bol_ingre.bin_qtde) AS 'qtde'
           FROM tb_bol_ingre
           WHERE tb_bol_ingre.sal_cd     = @sal_cd
           AND   tb_bol_ingre.fil_cd     = @fil_cd
           AND   tb_bol_ingre.bol_dt_mov = @Data
           AND   tb_bol_ingre.sre_data   = @Data
           AND   tb_bol_ingre.opt_cd IN (1, 3)
           AND   tb_bol_ingre.igt_cd IN (1, 3)
           AND   tb_bol_ingre.bin_dev    = 'N'
           GROUP BY tb_bol_ingre.sre_horario
          ) w ON z.ses_horario = w.sre_horario) LEFT JOIN
         (SELECT tb_bol_ingre.sre_horario, 
                 SUM(tb_bol_ingre.bin_qtde) AS 'qtde'
          FROM tb_bol_ingre
          WHERE tb_bol_ingre.sal_cd     = @sal_cd
          AND   tb_bol_ingre.fil_cd     = @fil_cd
          AND   tb_bol_ingre.bol_dt_mov = @Data
          AND   tb_bol_ingre.sre_data   = @Data
          AND   tb_bol_ingre.opt_cd IN (1, 3)
          AND   tb_bol_ingre.igt_cd IN (2, 4)
          AND   tb_bol_ingre.bin_dev    = 'N'
          GROUP BY tb_bol_ingre.sre_horario
         ) x ON z.ses_horario = x.sre_horario) LEFT JOIN
        (SELECT tb_bol_ingre.sre_horario, 
                SUM(tb_bol_ingre.bin_qtde) AS 'qtde'
         FROM tb_bol_ingre
         WHERE tb_bol_ingre.sal_cd     = @sal_cd
         AND   tb_bol_ingre.fil_cd     = @fil_cd
         AND   tb_bol_ingre.bol_dt_mov = @Data
         AND   tb_bol_ingre.sre_data   = @Data
         AND   tb_bol_ingre.opt_cd IN (1, 3)
         AND   tb_bol_ingre.igt_cd     = 9
         AND   tb_bol_ingre.bin_dev    = 'N'
         GROUP BY tb_bol_ingre.sre_horario
        ) y ON z.ses_horario = y.sre_horario
   ORDER BY z.ses_horario

   SELECT @Erro = @@ERROR
    
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      End

GO

--****************************************************
CREATE PROCEDURE upBOLETIM_VENDA_ANTECIPADA2
   (@Data      datetime,
    @sal_cd    int,
    @fil_cd    int,
    @Erro      int OUTPUT,
    @MsgErr    varchar(255) OUTPUT)
AS

   DECLARE @HoraLimite  char(5),
           @HoraMaxSes  datetime,
           @DataIniPer  datetime,
           @DataFimPer  datetime


   SELECT @HoraMaxSes = par_hora_max_ses
   FROM tb_bol_param
   WHERE tb_bol_param. bol_dt_mov = @Data
     
   SELECT @DataIniPer = convert(datetime, convert(char(10),@Data,103) + ' ' + convert(char(10),@HoraMaxSes,108), 103)
   SELECT @DataFimPer = convert(datetime, convert(char(10),DATEADD(Day, 1, @Data),103) + ' ' + convert(char(10),@HoraMaxSes,108), 103)
     
   SELECT tb_bol_ingre.sre_data,
          0                               'ordem', 
          tb_bol_ingre.igt_cd,  
          SUM(tb_bol_ingre.bin_qtde)      'qtde', 
          tb_bol_ingre.ing_valor          'ing_valor',
          tb_bol_tp_ingr.igt_desc         'igt_desc', 
          tb_bol_ingre.ppr_cd
   FROM tb_bol_ingre,
        tb_bol_tp_ingr
   WHERE tb_bol_ingre.igt_cd = tb_bol_tp_ingr.igt_cd
   AND   tb_bol_ingre.sal_cd = @sal_cd
   AND   tb_bol_ingre.fil_cd = @fil_cd
   AND   tb_bol_ingre.bol_dt_mov   = @Data
   AND   tb_bol_ingre.sre_data     > @Data
   AND   tb_bol_ingre.ing_dt_venda between @DataIniPer AND @DataFimPer
   AND   tb_bol_ingre.opt_cd = 1            -- Operação Venda Normal
   AND   tb_bol_ingre.igt_cd IN ( 1, 2, 9 ) -- Inteira e Meia
   AND   tb_bol_ingre.cxp_talao = 0         -- não talão
   AND   tb_bol_ingre.bin_dev = 'N'
   GROUP BY tb_bol_ingre.sre_data,
            tb_bol_ingre.igt_cd, 
            tb_bol_tp_ingr.igt_desc, 
            tb_bol_ingre.ppr_cd,
            tb_bol_ingre.ing_valor

   UNION

   SELECT tb_bol_ingre.sre_data,
          2                               'ordem', 
          tb_bol_ingre.igt_cd, 
          SUM(tb_bol_ingre.bin_qtde)      'qtde', 
          tb_bol_ingre.ing_valor          'ing_valor',
          tb_bol_tp_ingr.igt_desc         'igt_desc', 
          tb_bol_ingre.ppr_cd
   FROM tb_bol_ingre,
        tb_bol_tp_ingr
   WHERE tb_bol_ingre.igt_cd = tb_bol_tp_ingr.igt_cd
   AND   tb_bol_ingre.sal_cd = @sal_cd
   AND   tb_bol_ingre.fil_cd = @fil_cd
   AND   tb_bol_ingre.bol_dt_mov   = @Data
   AND   tb_bol_ingre.sre_data     > @Data
   AND   tb_bol_ingre.ing_dt_venda between @DataIniPer AND @DataFimPer
   AND   tb_bol_ingre.opt_cd = 1         -- Operação Venda Normal + devolução
   AND   tb_bol_ingre.igt_cd IN ( 3, 4 ) -- Inteira e Meia Promoção
   AND   tb_bol_ingre.cxp_talao = 0      -- não talão
   AND   tb_bol_ingre.bin_dev = 'N'
   GROUP BY tb_bol_ingre.sre_data,
            tb_bol_ingre.igt_cd, 
            tb_bol_tp_ingr.igt_desc, 
            tb_bol_ingre.ppr_cd,
            tb_bol_ingre.ing_valor

   UNION

   SELECT tb_bol_ingre.sre_data,
          4                                  'ordem', 
          tb_bol_ingre.igt_cd, 
          SUM(tb_bol_ingre.bin_qtde)         'qtde', 
          tb_bol_ingre.ing_valor             'ing_valor',
          tb_bol_tp_ingr.igt_desc + ' TA'    'igt_desc', 
          tb_bol_ingre.ppr_cd
   FROM tb_bol_ingre,
        tb_bol_tp_ingr
   WHERE tb_bol_ingre.igt_cd = tb_bol_tp_ingr.igt_cd
   AND   tb_bol_ingre.sal_cd = @sal_cd
   AND   tb_bol_ingre.fil_cd = @fil_cd
   AND   tb_bol_ingre.bol_dt_mov   = @Data
   AND   tb_bol_ingre.sre_data     > @Data
   AND   tb_bol_ingre.ing_dt_venda between @DataIniPer AND @DataFimPer
   AND   tb_bol_ingre.opt_cd = 1            -- Operação Venda Normal + devolução
   AND   tb_bol_ingre.igt_cd IN ( 1, 2, 9 ) -- Inteira e Meia
   AND   tb_bol_ingre.cxp_talao = 1         -- Talão
   AND   tb_bol_ingre.bin_dev = 'N'
   GROUP BY tb_bol_ingre.sre_data,
            tb_bol_ingre.igt_cd, 
            tb_bol_tp_ingr.igt_desc, 
            tb_bol_ingre.ppr_cd,
            tb_bol_ingre.ing_valor

   UNION

   SELECT tb_bol_ingre.sre_data,
          6                                  'ordem', 
          tb_bol_ingre.igt_cd, 
          SUM(tb_bol_ingre.bin_qtde)         'qtde', 
          tb_bol_ingre.ing_valor             'ing_valor',
          tb_bol_tp_ingr.igt_desc + ' TA'    'igt_desc', 
          tb_bol_ingre.ppr_cd
   FROM tb_bol_ingre,
        tb_bol_tp_ingr
   WHERE tb_bol_ingre.igt_cd = tb_bol_tp_ingr.igt_cd
   AND   tb_bol_ingre.sal_cd = @sal_cd
   AND   tb_bol_ingre.fil_cd = @fil_cd
   AND   tb_bol_ingre.bol_dt_mov   = @Data
   AND   tb_bol_ingre.sre_data     > @Data
   AND   tb_bol_ingre.ing_dt_venda between @DataIniPer AND @DataFimPer
   AND   tb_bol_ingre.opt_cd = 1         -- Operação Venda Normal + devolução
   AND   tb_bol_ingre.igt_cd IN ( 3, 4 ) -- Inteira e Meia Promoção
   AND   tb_bol_ingre.cxp_talao = 1      -- Talão
   AND   tb_bol_ingre.bin_dev = 'N'
   GROUP BY tb_bol_ingre.sre_data,
            tb_bol_ingre.igt_cd, 
            tb_bol_tp_ingr.igt_desc, 
            tb_bol_ingre.ppr_cd,
            tb_bol_ingre.ing_valor

   ORDER BY 1, 2, 3

    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          Return
       End

GO

--****************************************************
CREATE PROCEDURE upBOLETIM_FORMA_PAGTO2
   (@Data      datetime,
    @sal_cd    int,
    @fil_cd    int,
    @Erro      int OUTPUT,
    @MsgErr    varchar(255) OUTPUT)
AS
   DECLARE @HoraMaxSes     datetime,
           @DataIniPer     datetime,
           @DataFimPer     datetime

   SELECT @HoraMaxSes = par_hora_max_ses
   FROM tb_parametro
	  
   SELECT @DataIniPer = CONVERT(datetime, CONVERT(char(10),@Data,103) + ' ' + CONVERT(char(10),@HoraMaxSes,108), 103)
   SELECT @DataFimPer = CONVERT(datetime, CONVERT(char(10),DATEADD(Day, 1, @Data),103) + ' ' + CONVERT(char(10),@HoraMaxSes,108), 103)

   SELECT y.pgt_cd, 
          y.pgt_desc, 
          CASE WHEN x.valor IS NULL THEN 0
               ELSE x.valor
          END valor
   FROM (
         SELECT tb_bol_ingre.pgt_cd, 
                SUM(tb_bol_ingre.bin_qtde * tb_bol_ingre.ing_valor) valor
         FROM tb_bol_ingre
         WHERE tb_bol_ingre.sal_cd = @sal_cd
         AND   tb_bol_ingre.fil_cd = @fil_cd
         AND   tb_bol_ingre.ing_dt_venda between @DataIniPer AND @DataFimPer
         AND   tb_bol_ingre.bol_dt_mov = @Data
         AND   tb_bol_ingre.bin_dev = 'N'
         GROUP BY tb_bol_ingre.pgt_cd
       ) x,
       tb_bol_pag_tp y
   WHERE y.pgt_cd = x.pgt_cd
   ORDER BY y.pgt_cd

   SELECT @Erro = @@ERROR
    
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      End

GO

--****************************************************
CREATE PROCEDURE upBOLETIM_NUMERACAO_TALAO2
   (@Data      datetime,
    @sal_cd    int,
    @fil_cd    int,
    @Erro      int OUTPUT,
    @MsgErr    varchar(255) OUTPUT)
AS

   DECLARE @HoraMaxSes     datetime,
           @DataIniPer     datetime,
           @DataFimPer     datetime

   SELECT @HoraMaxSes = par_hora_max_ses
   FROM tb_bol_param
   WHERE tb_bol_param. bol_dt_mov = @Data
     
   SELECT @DataIniPer = convert(datetime, convert(char(10),@Data,103) + ' ' + convert(char(10),@HoraMaxSes,108), 103)
   SELECT @DataFimPer = convert(datetime, convert(char(10),DATEADD(Day, 1, @Data),103) + ' ' + convert(char(10),@HoraMaxSes,108), 103)

   SELECT tb_bol_tp_ingr.igt_desc, 
          tb_bol_talao.num_talao_ini 'numIni', 
          tb_bol_talao.num_talao_fim 'numFim'
   FROM tb_bol_talao,
        tb_bol_tp_ingr
   WHERE tb_bol_talao.igt_cd = tb_bol_tp_ingr.igt_cd
   AND   tb_bol_talao.sal_cd = @sal_cd
   AND   tb_bol_talao.fil_cd = @fil_cd
   AND   tb_bol_talao.bol_dt_mov = @Data
   ORDER BY 1

   SELECT @Erro = @@ERROR
    
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      End
GO

--*******************************************************
CREATE PROCEDURE upTB_BOL_FILME_CARTAZ
   (@DataMov datetime,
    @Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS

    SELECT DISTINCT tb_boletim.fil_cd, 
                    tb_boletim.sal_cd, 
                    tb_bol_sala.sal_desc + ' - ' + tb_bol_filme.fil_nm  'Sala - Filme',
                    tb_boletim.bol_status
    FROM tb_boletim,
         tb_bol_filme,
         tb_bol_sala
    WHERE tb_boletim.sal_cd     = tb_bol_sala.sal_cd
    AND   tb_boletim.bol_dt_mov = tb_bol_sala.bol_dt_mov
    AND   tb_boletim.fil_cd     = tb_bol_filme.fil_cd
    AND   tb_boletim.bol_dt_mov = tb_bol_filme.bol_dt_mov
    AND   tb_boletim.bol_dt_mov = @DataMov
    ORDER BY tb_bol_sala.sal_desc + ' - ' + tb_bol_filme.fil_nm

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END
GO

--*******************************************************
CREATE PROCEDURE upTB_BOLETIM_VERIFICA
   (@DataMov datetime,
    @Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS

   DECLARE @dtMov  datetime

   SELECT @dtMov = tb_boletim.bol_dt_mov
   FROM tb_boletim
   WHERE tb_boletim.bol_dt_mov = @dtMov

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

   SELECT @MsgErr = ''
   SELECT @Erro   = 0
      
   IF @dtMov IS NULL
      BEGIN
         SELECT @MsgErr = 'Boletim não existe'
         SELECT @Erro   = 1
      END
   
GO

--****************************************************
CREATE PROCEDURE upBOLETIM_CATRACA2
	(@Data		datetime,
	 @sal_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS

    SELECT tb_bol_catraca.cat_nm,
           tb_bol_catraca.ctc_ini_cont,
           tb_bol_catraca.ctc_fim_cont
    FROM tb_bol_catraca,
         tb_bol_catraca_sala
    WHERE tb_bol_catraca.cat_cd      = tb_bol_catraca_sala.cat_cd
    AND   tb_bol_catraca.bol_dt_mov  = tb_bol_catraca_sala.bol_dt_mov
    AND   tb_bol_catraca.bol_dt_mov  = @Data
    AND   tb_bol_catraca_sala.sal_cd = @sal_cd

    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          RETURN
       END

GO

--****************************************************
CREATE PROCEDURE upBOLETIM_INGRESSO_S_USO2
	(@Data		datetime,
	 @sal_cd 	int,
	 @fil_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS

    SELECT SUM(bin_qtde) AS qtdeSUso
    FROM tb_bol_ingre
    WHERE sal_cd     = @sal_cd
    AND   fil_cd     = @fil_cd
    AND   sre_data   = @Data
    AND   bol_dt_mov = @Data
    AND   ing_status <> 1
    AND   bin_dev    = 'N'


    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          RETURN
       END

GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufSESSOES]')  and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufSESSOES]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufNacionalidade]')  and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufNacionalidade]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufSEMANA]')  and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufSEMANA]
GO

--****************************************************
CREATE FUNCTION ufSESSOES (@DataMov    datetime,
                           @emp_cd     int,
                           @cin_cd     int,
                           @sal_cd     int,
                           @fil_cd     int)
RETURNS VARCHAR(255)
AS
   BEGIN	
      DECLARE @ret         VARCHAR(255),
              @ses_horario datetime
      
      DECLARE curSessoes CURSOR
      FOR
         SELECT ses_horario
         FROM tb_bol_sessao
         WHERE bol_dt_mov = @DataMov
         AND   emp_cd     = @emp_cd
         AND   cin_cd     = @cin_cd
         AND   sal_cd     = @sal_cd
         AND   fil_cd     = @fil_cd
         AND   sre_data   = @DataMov
         

      OPEN curSessoes
      
      FETCH NEXT FROM curSessoes INTO @ses_horario

      IF (@@FETCH_STATUS <> -1)
         BEGIN
            IF (@@FETCH_STATUS <> -2)
               SELECT @ret =RTRIM(convert(char(5), @ses_horario, 108))
       
            FETCH NEXT FROM curSessoes INTO @ses_horario
            WHILE (@@FETCH_STATUS <> -1)
               BEGIN
                  IF (@@FETCH_STATUS <> -2)
                     SELECT @ret = @ret + '/' + RTRIM(convert(char(5), @ses_horario, 108))
           
              FETCH NEXT FROM curSessoes INTO @ses_horario
            END
         END
      
      CLOSE curSessoes
      DEALLOCATE curSessoes

      RETURN @ret
   END	
GO


--****************************************************
CREATE FUNCTION ufNacionalidade (@fil_id_nacio  VARCHAR(1))
RETURNS VARCHAR(255)
AS
   BEGIN
      DECLARE @ret VARCHAR(255)

      IF @fil_id_nacio = 'N'
          SELECT @ret =  'Nacional'
      ELSE
         IF @fil_id_nacio = 'E'
            SELECT @ret =   'Estrangeiro'
         ELSE
            SELECT @ret =   ''
      
      RETURN @ret
   END
GO

--****************************************************
CREATE FUNCTION ufSEMANA (@DataMov    datetime,
                          @fil_cd     int)
RETURNS INT
AS
   BEGIN	
      DECLARE @ret            INT,
              @dias           INT,
              @fil_dt_ini     datetime,
              @fil_dt_fim     datetime

              
      SELECT @fil_dt_ini = fil_dt_ini,
             @fil_dt_fim = fil_dt_fim
      FROM tb_bol_filme
      WHERE bol_dt_mov = @DataMov
      AND   fil_cd     = @fil_cd
      
      SELECT @dias = ABS(DATEDIFF(day, @fil_dt_ini, @DataMov)) + 1
      
      SELECT @ret = CEILING(@dias/7)
      
      IF @ret = 0
          SELECT @ret = + 1

      RETURN @ret
   END	
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upDIA_SEMANA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upDIA_SEMANA
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upSESSOES_DIA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upSESSOES_DIA
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upSESSOES_DIA_PROMOCAO') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upSESSOES_DIA_PROMOCAO
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upINGRESSO_OPERACAO') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upINGRESSO_OPERACAO
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upCOMBO_OPERACAO') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upCOMBO_OPERACAO
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upDATA_SISTEMA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upDATA_SISTEMA
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upVALOR_EM_CAIXA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upVALOR_EM_CAIXA
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upPOSICAO_CAIXAS') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upPOSICAO_CAIXAS
GO

--****************************************************
CREATE PROCEDURE upCOMBO_OPERACAO
    (@ope_cd       int,
     @vcb_cd       numeric(12),
     @Erro         int OUTPUT,
     @MsgErr       varchar(255) OUTPUT)
AS
    IF @ope_cd IS NOT NULL
      
    SELECT e.cin_nm, 
           e.cin_cid_end, 
           e.cin_uf_end, 
           c.cbo_nm,
           c.cbo_desc,
           a.ope_dt_operacao, 
           a.cxa_cd,         
           b.vcb_qtde, 
           b.vcb_valor, 
           b.vcb_cd
          FROM TB_OPERACAO a,
           TB_VENDA_COMBO b,
           TB_COMBO c,
           TB_CAIXA d,
           TB_CINEMA e
         WHERE b.ope_cd = a.ope_cd
       AND c.cbo_cd = b.cbo_cd
       AND d.cxa_cd = a.cxa_cd
       AND e.cin_cd = d.cin_cd
       AND a.ope_dt_des IS NULL
       AND b.vcb_dt_canc IS NULL
       AND a.ope_cd = @ope_cd
    ELSE
    SELECT e.cin_nm, 
           e.cin_cid_end, 
           e.cin_uf_end, 
           c.cbo_nm,
           c.cbo_desc,
           a.ope_dt_operacao, 
           a.cxa_cd,         
           b.vcb_qtde, 
           b.vcb_valor, 
           b.vcb_cd
          FROM TB_OPERACAO a,
           TB_VENDA_COMBO b,
           TB_COMBO c,
           TB_CAIXA d,
           TB_CINEMA e
         WHERE b.ope_cd = a.ope_cd
       AND c.cbo_cd = b.cbo_cd
       AND d.cxa_cd = a.cxa_cd
       AND e.cin_cd = d.cin_cd
       AND a.ope_dt_des IS NULL
       AND b.vcb_dt_canc IS NULL
       AND b.vcb_cd = @vcb_cd
       
     SELECT @Erro = @@ERROR
   
     IF @Erro <> 0
        BEGIN
           SELECT @MsgErr = description
           FROM master..sysmessages
           WHERE error = @Erro
         
           RETURN
        END

GO

--****************************************************
CREATE PROCEDURE upDATA_SISTEMA
    (@Erro   int          OUTPUT,
     @MsgErr varchar(255) OUTPUT)
AS
    SELECT convert(datetime,convert(char(11),GETDATE())) + convert(datetime,convert(char(8),GETDATE(),108))
    
    SELECT @Erro = @@ERROR
   
    IF @Erro <> 0
        BEGIN
           SELECT @MsgErr = description
           FROM master..sysmessages
           WHERE error = @Erro
         
           RETURN
        END

GO

--****************************************************
/*
declare @erro int, @msgerr varchar(255), @data datetime
select @data = '25/12/2005' --convert(datetime, '25/12/2005', 103 )
exec upDIA_SEMANA @data, @erro, @msgerr
*/
--****************************************************
CREATE PROCEDURE upDIA_SEMANA
    (@DataExibicao  datetime,
     @Erro          int OUTPUT,
     @MsgErr        varchar(255) OUTPUT)
AS
    declare @diaSemana smallint
    
--    set DATEFIRST 1
    
    select @diaSemana = 8
      From tb_feriado
     where fer_data = convert(datetime, @DataExibicao, 103)
    
    if @diaSemana is null
        select @diaSemana = datepart(dw,convert(datetime, @DataExibicao, 103))
    
    select @diaSemana diaSemana
    
    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          Return
       End

GO

--****************************************************
CREATE PROCEDURE upINGRESSO_OPERACAO
    (@ope_cd       int,
     @ing_cd       numeric(12),
     @Erro         int OUTPUT,
     @MsgErr       varchar(255) OUTPUT)
AS
    DECLARE @par_imp_MFIM BIT
    
    select @par_imp_MFIM = par_imp_MFIM
      from tb_parametro

   IF @ope_cd IS NOT NULL
      SELECT e.cin_nm, 
             e.cin_cid_end, 
             e.cin_uf_end, 
             d.sal_desc,
             c.fil_nm,
             b.sre_data, 
             convert(char(5),b.sre_horario,108) 'sre_horario',
             a.ope_dt_operacao, 
             a.cxa_cd,
             b.igt_cd,
             CASE @par_imp_MFIM
                WHEN 1 THEN CASE b.igt_cd
                               WHEN 1 THEN 'HOMEM'
                               WHEN 2 THEN 'MULHER'
                               WHEN 3 THEN RTRIM(SUBSTRING(f.ppr_desc,1,6)) + '-HOM.'
                               WHEN 4 THEN RTRIM(SUBSTRING(f.ppr_desc,1,6)) + '-MUL.'
                               WHEN 9 THEN 'CORTESIA'
                            END
                ELSE CASE b.igt_cd
                        WHEN 1 THEN 'INTEIRA'
                        WHEN 2 THEN 'MEIA'
                        WHEN 3 THEN RTRIM(SUBSTRING(f.ppr_desc,1,6)) + '-INT.'
                        WHEN 4 THEN RTRIM(SUBSTRING(f.ppr_desc,1,6)) + '-MEIA'
                        WHEN 9 THEN 'CORTESIA'
                     END
             END 'igt_tipo', 
             b.ing_valor, 
             b.ing_cd, 
             d.sal_lugares, 
             b.ppr_cd, 
             e.cin_end,
             e.cin_num_end,
             e.cin_cmp_end,
             e.cin_brr_end,
             e.cin_cnpj,
             e.cin_inscricao,
             f.ppr_desc, 
             f.ppr_patrocinador,
             b.ing_num_ing
      FROM ((((TB_OPERACAO a INNER JOIN TB_VENDA_INGRESSO b ON a.ope_cd = b.ope_cd) 
              INNER JOIN TB_FILME c ON b.fil_cd = c.fil_cd)
             INNER JOIN TB_SALA d ON b.sal_cd = d.sal_cd)
            INNER JOIN TB_CINEMA e ON d.cin_cd = e.cin_cd)
           LEFT JOIN TB_PROG_PRECO f ON b.ppr_cd = f.ppr_cd
      WHERE a.ope_dt_des IS NULL
      AND   b.ing_dt_canc IS NULL
      AND   a.ope_cd = @ope_cd
      ORDER BY 5, 6, 11
   ELSE
      SELECT e.cin_nm, 
             e.cin_cid_end, 
             e.cin_uf_end, 
             d.sal_desc,
             c.fil_nm,
             b.sre_data, 
             convert(char(5),b.sre_horario,108) 'sre_horario',
             a.ope_dt_operacao, 
             a.cxa_cd,         
             CASE @par_imp_MFIM
                WHEN 1 THEN CASE b.igt_cd
                               WHEN 1 THEN 'HOMEM'
                               WHEN 2 THEN 'MULHER'
                               WHEN 3 THEN RTRIM(SUBSTRING(f.ppr_desc,1,6)) + '-HOM.'
                               WHEN 4 THEN RTRIM(SUBSTRING(f.ppr_desc,1,6)) + '-MUL.'
                               WHEN 9 THEN 'CORTESIA'
                            END
                ELSE CASE b.igt_cd
                        WHEN 1 THEN 'INTEIRA'
                        WHEN 2 THEN 'MEIA'
                        WHEN 3 THEN RTRIM(SUBSTRING(f.ppr_desc,1,6)) + '-INT.'
                        WHEN 4 THEN RTRIM(SUBSTRING(f.ppr_desc,1,6)) + '-MEIA'
                        WHEN 9 THEN 'CORTESIA'
                     END
             END 'igt_tipo', 
             b.ing_valor, 
             b.ing_cd, 
             d.sal_lugares, 
             b.ppr_cd, 
             e.cin_end,
             e.cin_num_end,
             e.cin_cmp_end,
             e.cin_brr_end,
             e.cin_cnpj,
             e.cin_inscricao,
             f.ppr_desc, 
             f.ppr_patrocinador,
             b.ing_num_ing
      FROM ((((TB_OPERACAO a INNER JOIN TB_VENDA_INGRESSO b ON a.ope_cd = b.ope_cd)
              INNER JOIN TB_FILME c ON c.fil_cd = b.fil_cd)
             INNER JOIN TB_SALA d ON b.sal_cd = d.sal_cd)
            INNER JOIN TB_CINEMA e ON d.cin_cd = e.cin_cd)
           LEFT JOIN TB_PROG_PRECO f ON b.ppr_cd = f.ppr_cd
     WHERE a.ope_dt_des IS NULL
     AND   b.ing_dt_canc IS NULL
     AND   b.ing_cd = @ing_cd
     ORDER BY 5, 6, 11
     
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
       
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upPOSICAO_CAIXAS
   (@Erro   int OUTPUT,
         @MsgErr varchar(255) OUTPUT)
AS
     SELECT b.cxa_desc 'Caixa',
            convert(char(11), a.cxp_dt_abertura, 103) + convert(char(8), a.cxp_dt_abertura, 108) 'Data Abertura',
            c.usu_nm 'Usuário Abertura',
            SUM(d.ope_valor*e.opt_sinal) 'Total Caixa'
     FROM TB_CAIXA_MOVTO a,
          TB_CAIXA b,
          TB_USUARIO c,
          TB_OPERACAO d,
          TB_OPERACAO_TIPO e
     WHERE a.cxp_status   IN (0,1)
     AND   a.cxa_cd       = b.cxa_cd
     AND   a.usu_abertura = c.usu_cd
     AND   a.cxp_cd       = d.cxp_cd
     AND   d.opt_cd       = e.opt_cd
     AND   d.ope_dt_des IS NULL
     GROUP BY b.cxa_desc,
              convert(char(11), a.cxp_dt_abertura, 103) + convert(char(8), a.cxp_dt_abertura, 108),
              c.usu_nm
              
     SELECT @Erro = @@ERROR
   
     IF @Erro <> 0
        BEGIN
           SELECT @MsgErr = description
           FROM master..sysmessages
           WHERE error = @Erro
         
           RETURN
        END

GO

--****************************************************
/*
declare @erro int, @msgerr varchar(255), @data datetime
select @data = convert(datetime, '6/8/2005', 103 )
exec upSESSOES_DIA @data, 8, 1, @erro, @msgerr
*/
--****************************************************
CREATE PROCEDURE upSESSOES_DIA
    (@DataExibicao datetime,
     @fil_cd       int,
     @sal_cd       int,
     @Erro         int OUTPUT,
     @MsgErr        varchar(255) OUTPUT)
AS
    declare @par_hora_limite12   datetime,
            @par_hora_limite23   datetime,
            @par_hora_limite34   datetime,
            @par_hora_limite45   datetime,
            @par_hora_limite56   datetime,
            @diaSemana           smallint,
            @pre_vl_inteira_ate  money,
            @pre_vl_inteira_apos money,
            @pre_vl_meia_ate     money,
            @pre_vl_meia_apos    money,
            @pre_vl_inteira3     money,
            @pre_vl_inteira4     money,
            @pre_vl_inteira5     money,
            @pre_vl_inteira6     money,
            @pre_vl_meia3        money,
            @pre_vl_meia4        money,
            @pre_vl_meia5        money,
            @pre_vl_meia6        money
            
    select @par_hora_limite12 = CONVERT(datetime, '1900-01-01 ' + convert(char(8),par_hora_limite12 ,108)),
           @par_hora_limite23 = CONVERT(datetime, '1900-01-01 ' + convert(char(8),par_hora_limite23 ,108)),
           @par_hora_limite34 = CONVERT(datetime, '1900-01-01 ' + convert(char(8),par_hora_limite34 ,108)),
           @par_hora_limite45 = CONVERT(datetime, '1900-01-01 ' + convert(char(8),par_hora_limite45 ,108)),
           @par_hora_limite56 = CONVERT(datetime, '1900-01-01 ' + convert(char(8),par_hora_limite56 ,108))
      from tb_parametro
    
    --SET DATEFIRST 1
    
    ------------------------
    -- Pega dia da Semana --
    ------------------------
    
    select @diaSemana = 8
      From tb_feriado
     where fer_data = convert(datetime, @DataExibicao, 103)
    
    if @diaSemana is null
        select @diaSemana = datepart(dw,convert(datetime, @DataExibicao, 103))
    ---------------------------------------------
    -- Pega Preço do Filme na Data de Exibição --
    ---------------------------------------------
    
    select @pre_vl_inteira_ate  = a.pre_vl_inteira_ate,
           @pre_vl_inteira_apos = a.pre_vl_inteira_apos,
           @pre_vl_inteira3     = a.pre_vl_inteira3,
           @pre_vl_inteira4     = a.pre_vl_inteira4,
           @pre_vl_inteira5     = a.pre_vl_inteira5,
           @pre_vl_inteira6     = a.pre_vl_inteira6,
           @pre_vl_meia_ate     = a.pre_vl_meia_ate,
           @pre_vl_meia_apos    = a.pre_vl_meia_apos,
           @pre_vl_meia3        = a.pre_vl_meia3,
           @pre_vl_meia4        = a.pre_vl_meia4,
           @pre_vl_meia5        = a.pre_vl_meia5,
           @pre_vl_meia6        = a.pre_vl_meia6
      from tb_preco a,
           tb_prog_preco b
     Where a.ppr_cd = b.ppr_cd
       and b.ppr_flg_promocao = 0
       and convert(datetime, @DataExibicao, 103) between b.ppr_dt_ini and b.ppr_dt_fim
       and a.fil_cd = @fil_cd
       and a.pre_dia_semana = @diaSemana
       and a.pre_dt_des is null
       and b.ppr_dt_des is null
    
    ---------------------------------------------
    -- Caso não tenha preço para o filme       --
    -- Pega Preço Genérico na Data de Exibição --
    ---------------------------------------------
    
    if @pre_vl_inteira_ate is null
          select @pre_vl_inteira_ate  = a.pre_vl_inteira_ate,
                 @pre_vl_inteira_apos = a.pre_vl_inteira_apos,
                 @pre_vl_inteira3     = a.pre_vl_inteira3,
                 @pre_vl_inteira4     = a.pre_vl_inteira4,
                 @pre_vl_inteira5     = a.pre_vl_inteira5,
                 @pre_vl_inteira6     = a.pre_vl_inteira6,
                 @pre_vl_meia_ate     = a.pre_vl_meia_ate,
                 @pre_vl_meia_apos    = a.pre_vl_meia_apos,
                 @pre_vl_meia3        = a.pre_vl_meia3,
                 @pre_vl_meia4        = a.pre_vl_meia4,
                 @pre_vl_meia5        = a.pre_vl_meia5,
                 @pre_vl_meia6        = a.pre_vl_meia6
          from tb_preco a,
               tb_prog_preco b
         Where a.ppr_cd = b.ppr_cd
           and b.ppr_flg_promocao = 0
           and convert(datetime, @DataExibicao, 103) between b.ppr_dt_ini and b.ppr_dt_fim
           and a.fil_cd = 0
           and a.pre_dia_semana = @diaSemana
           and a.pre_dt_des is null
           and b.ppr_dt_des is null
    
    SELECT x.fil_cd, 
           x.sal_cd, 
           x.ses_horario, 
           convert(char(5),x.ses_horario,108)               AS 'SESSÃO',
           x.sal_lugares - isnull(y.sre_lugares_vendidos,0) AS 'LOTAÇÃO',  
           x.pre_vl_inteira                                 AS 'INTEIRA', 
           x.pre_vl_meia                                    AS 'MEIA',
           isnull(y.sre_meias,0) / y.sre_lugares            AS precMeias,
           isnull(y.sre_cortesias,0) / y.sre_lugares        AS precCortesias
      FROM (SELECT b.fil_cd, 
                   b.sal_cd, 
                   b.ses_horario, 
                   d.sal_desc,
                   CASE 
                      WHEN e.sal_lugares IS NOT NULL THEN e.sal_lugares
                      ELSE d.sal_lugares
                   END AS sal_lugares,
                  CASE 
                     WHEN b.ses_horario <  @par_hora_limite12 THEN @pre_vl_inteira_ate
                     WHEN b.ses_horario >= @par_hora_limite12 AND 
                          b.ses_horario <  @par_hora_limite23 THEN @pre_vl_inteira_apos
                     WHEN b.ses_horario >= @par_hora_limite23 AND 
                          b.ses_horario <  @par_hora_limite34 THEN @pre_vl_inteira3
                     WHEN b.ses_horario >= @par_hora_limite34 AND 
                          b.ses_horario <  @par_hora_limite45 THEN @pre_vl_inteira4
                     WHEN b.ses_horario >= @par_hora_limite45 AND 
                          b.ses_horario <  @par_hora_limite56 THEN @pre_vl_inteira5
                     WHEN b.ses_horario >= @par_hora_limite56 THEN @pre_vl_inteira6
                  END AS pre_vl_inteira,
                  CASE               
                     WHEN b.ses_horario <  @par_hora_limite12 THEN @pre_vl_meia_ate
                     WHEN b.ses_horario >= @par_hora_limite12 AND 
                          b.ses_horario <  @par_hora_limite23 THEN @pre_vl_meia_apos
                     WHEN b.ses_horario >= @par_hora_limite23 AND 
                          b.ses_horario <  @par_hora_limite34 THEN @pre_vl_meia3
                     WHEN b.ses_horario >= @par_hora_limite34 AND 
                          b.ses_horario <  @par_hora_limite45 THEN @pre_vl_meia4
                     WHEN b.ses_horario >= @par_hora_limite45 AND 
                          b.ses_horario <  @par_hora_limite56 THEN @pre_vl_meia5
                     WHEN b.ses_horario >= @par_hora_limite56 THEN @pre_vl_meia6
                  END AS pre_vl_meia 
            FROM (((tb_programacao a INNER JOIN tb_sessao b ON a.prg_cd =  b.prg_cd)
                   INNER JOIN tb_filme c ON b.fil_cd =  c.fil_cd)
                   INNER JOIN tb_sala d ON b.sal_cd =  d.sal_cd)
                   LEFT JOIN (SELECT sal_cd, 
                                     sal_lugares
                              FROM tb_sala_lugar
                              WHERE CONVERT(DATETIME, @DataExibicao, 103) BETWEEN sal_dt_ini AND sal_dt_fim) e
                   ON b.sal_cd = e.sal_cd
            WHERE b.ses_dt_des IS NULL
            AND   a.prg_dt_des IS NULL
            AND   c.fil_dt_des IS NULL
            AND   d.sal_dt_des IS NULL
            AND   CONVERT(datetime, @DataExibicao, 103) BETWEEN a.prg_dt_ini AND a.prg_dt_fim
            AND   b.ses_dia_semana = @diaSemana
            AND   b.fil_cd         = @fil_cd
            AND   d.sal_cd         = @sal_cd
           ) x LEFT JOIN tb_sessao_real y
               ON (x.fil_cd      = y.fil_cd      AND
                   x.sal_cd      = y.sal_cd      AND
                   x.ses_horario = y.sre_horario AND  
                   CONVERT(datetime, @DataExibicao, 103) = y.sre_data)
      WHERE x.pre_vl_inteira IS NOT NULL
      ORDER BY x.sal_cd, 
               x.ses_horario
    
     SELECT @Erro = @@ERROR
   
     IF @Erro <> 0
        BEGIN
           SELECT @MsgErr = description
           FROM master..sysmessages
           WHERE error = @Erro
         
           RETURN
        END

GO

--****************************************************
/*
declare @erro int, @msgerr varchar(255), @data datetime, @horario datetime
select @data = convert(datetime, '1/7/2005', 103 )
select @horario = convert(datetime, '1/1/1900 13:30:00', 103 )
exec upSESSOES_DIA_PROMOCAO @data, 7, 1, @horario, @erro, @msgerr
*/
--****************************************************
CREATE PROCEDURE upSESSOES_DIA_PROMOCAO
    (@DataExibicao datetime,
     @fil_cd       int,
     @sal_cd       int,
     @ses_horario  datetime,
     @Erro         int OUTPUT,
     @MsgErr        varchar(255) OUTPUT)
AS

    declare @par_hora_limite12   datetime,
            @par_hora_limite23   datetime,
            @par_hora_limite34   datetime,
            @par_hora_limite45   datetime,
            @par_hora_limite56   datetime,
            @diaSemana           smallint,
            @pre_vl_inteira_ate  money,
       @pre_vl_inteira_apos money,
       @pre_vl_meia_ate     money,
       @pre_vl_meia_apos    money,
       @pre_vl_inteira3     money,
       @pre_vl_inteira4     money,
       @pre_vl_inteira5     money,
       @pre_vl_inteira6     money,
       @pre_vl_meia3        money,
       @pre_vl_meia4        money,
       @pre_vl_meia5        money,
            @pre_vl_meia6        money,
            @ppr_cd              int,
            @ppr_desc            varchar(50),
            @ppr_patroinador     varchar(50)

    select @par_hora_limite12 = CONVERT(datetime, '1900-01-01 ' + convert(char(8),par_hora_limite12 ,108)),
           @par_hora_limite23 = CONVERT(datetime, '1900-01-01 ' + convert(char(8),par_hora_limite23 ,108)),
           @par_hora_limite34 = CONVERT(datetime, '1900-01-01 ' + convert(char(8),par_hora_limite34 ,108)),
           @par_hora_limite45 = CONVERT(datetime, '1900-01-01 ' + convert(char(8),par_hora_limite45 ,108)),
           @par_hora_limite56 = CONVERT(datetime, '1900-01-01 ' + convert(char(8),par_hora_limite56 ,108))
      from tb_parametro
    
--    SET DATEFIRST 1
    
    ------------------------
    -- Pega dia da Semana --
    ------------------------
    
    select @diaSemana = 8
      From tb_feriado
     where fer_data = convert(datetime, @DataExibicao, 103)
    
    if @diaSemana is null
        select @diaSemana = datepart(dw,convert(datetime, @DataExibicao, 103))
    
    if @pre_vl_inteira_ate is null
        select b.ppr_cd 'ppr_cd',
           b.ppr_desc 'PROMOÇÃO',
           b.ppr_patrocinador 'PATROCINADOR',
           CASE 
              WHEN @ses_horario <  @par_hora_limite12 THEN
                 a.pre_vl_inteira_ate
              WHEN @ses_horario >= @par_hora_limite12 AND @ses_horario < @par_hora_limite23 THEN
                 a.pre_vl_inteira_apos
              WHEN @ses_horario >= @par_hora_limite23 AND @ses_horario < @par_hora_limite34 THEN
                 a.pre_vl_inteira3
              WHEN @ses_horario >= @par_hora_limite34 AND @ses_horario < @par_hora_limite45 THEN
                 a.pre_vl_inteira4
              WHEN @ses_horario >= @par_hora_limite45 AND @ses_horario < @par_hora_limite56 THEN
                 a.pre_vl_inteira5
              WHEN @ses_horario >= @par_hora_limite56 THEN
                 a.pre_vl_inteira6
           END 'INTEIRA',
           CASE               
              WHEN @ses_horario <  @par_hora_limite12 THEN
                 a.pre_vl_meia_ate
              WHEN @ses_horario >= @par_hora_limite12 AND @ses_horario < @par_hora_limite23 THEN
                 a.pre_vl_meia_apos
              WHEN @ses_horario >= @par_hora_limite23 AND @ses_horario < @par_hora_limite34 THEN
                 a.pre_vl_meia3
              WHEN @ses_horario >= @par_hora_limite34 AND @ses_horario < @par_hora_limite45 THEN
                 a.pre_vl_meia4
              WHEN @ses_horario >= @par_hora_limite45 AND @ses_horario < @par_hora_limite56 THEN
                 a.pre_vl_meia5
              WHEN @ses_horario >= @par_hora_limite56 THEN
                 a.pre_vl_meia6
           END 'MEIA'          
          from tb_preco a,
               tb_prog_preco b
         Where a.ppr_cd = b.ppr_cd
           and b.ppr_flg_promocao = 1
           and convert(datetime, @DataExibicao, 103) between b.ppr_dt_ini and b.ppr_dt_fim
           and a.fil_cd IN (0, @fil_cd)
           and a.pre_dia_semana = @diaSemana
           and a.pre_dt_des is null
           and b.ppr_dt_des is null
           
     SELECT @Erro = @@ERROR
   
     IF @Erro <> 0
        BEGIN
           SELECT @MsgErr = description
           FROM master..sysmessages
           WHERE error = @Erro
         
           RETURN
        END

GO

--****************************************************
CREATE PROCEDURE upVALOR_EM_CAIXA
    (@cxa_cd       int,
     @valorCaixa   money OUTPUT,
     @Erro         int OUTPUT,
     @MsgErr       varchar(255) OUTPUT)
AS
    declare @valorRecebido money,
            @ValorSangria  money
            
    SELECT @valorRecebido = ISNULL(SUM(tb_venda_ingresso.ing_valor), 0)
    FROM tb_venda_ingresso
    WHERE tb_venda_ingresso.ope_cd IN (SELECT tb_operacao.ope_cd
                                       FROM tb_caixa_movto,
                                            tb_operacao
                                       WHERE tb_caixa_movto.cxa_cd = tb_operacao.cxa_cd
                                       AND   tb_caixa_movto.cxp_cd = tb_operacao.cxp_cd
                                       AND   tb_caixa_movto.cxa_cd = @cxa_cd
                                       AND   tb_caixa_movto.cxp_dt_fechamento IS NULL
                                       AND   tb_operacao.ope_dt_des IS NULL)
    AND tb_venda_ingresso.ing_dt_canc IS NULL
    
    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
        BEGIN
           SELECT @MsgErr = description
           FROM master..sysmessages
           WHERE error = @Erro
         
           RETURN
        END

    SELECT @valorRecebido = @valorRecebido + ISNULL(SUM(tb_venda_combo.vcb_valor * tb_venda_combo.vcb_qtde), 0)
    FROM tb_venda_combo
    WHERE tb_venda_combo.ope_cd IN (SELECT tb_operacao.ope_cd
                                    FROM tb_caixa_movto,
                                         tb_operacao
                                    WHERE tb_caixa_movto.cxa_cd = tb_operacao.cxa_cd
                                    AND   tb_caixa_movto.cxp_cd = tb_operacao.cxp_cd
                                    AND   tb_caixa_movto.cxa_cd = @cxa_cd
                                    AND   tb_caixa_movto.cxp_dt_fechamento IS NULL
                                    AND   tb_operacao.ope_dt_des IS NULL)
    AND tb_venda_combo.vcb_dt_canc IS NULL
    
    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
        BEGIN
           SELECT @MsgErr = description
           FROM master..sysmessages
           WHERE error = @Erro
         
           RETURN
        END

    SELECT @valorRecebido = @valorRecebido + ISNULL(SUM(tb_operacao.ope_valor), 0)
    FROM tb_caixa_movto,
         tb_operacao
    WHERE tb_caixa_movto.cxa_cd = tb_operacao.cxa_cd
    AND   tb_caixa_movto.cxp_cd = tb_operacao.cxp_cd
    AND   tb_caixa_movto.cxa_cd = @cxa_cd
    AND   tb_caixa_movto.cxp_dt_fechamento IS NULL
    AND   tb_operacao.ope_dt_des IS NULL
    AND   tb_operacao.opt_cd = 2 -- Fundo de Caixa
    
    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
        BEGIN
           SELECT @MsgErr = description
           FROM master..sysmessages
           WHERE error = @Erro
         
           RETURN
        END
        
        
    SELECT @ValorSangria = ISNULL(SUM(tb_operacao.ope_valor), 0)
    FROM tb_caixa_movto,
         tb_operacao
    WHERE tb_caixa_movto.cxa_cd = tb_operacao.cxa_cd
    AND   tb_caixa_movto.cxp_cd = tb_operacao.cxp_cd
    AND   tb_caixa_movto.cxa_cd = @cxa_cd
    AND   tb_caixa_movto.cxp_dt_fechamento IS NULL
    AND   tb_operacao.opt_cd IN (5, 6)
    AND   tb_operacao.ope_dt_des IS NULL
    
    SELECT @Erro = @@ERROR
    
    IF @Erro <> 0
        BEGIN
           SELECT @MsgErr = description
           FROM master..sysmessages
           WHERE error = @Erro
         
           RETURN
        END
    
    SELECT @valorCaixa = @valorRecebido - @ValorSangria

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upExpurgo') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upExpurgo
GO

--****************************************************
CREATE PROCEDURE upExpurgo
	(@dias   int,
         @Erro   int OUTPUT,
         @MsgErr varchar(255) OUTPUT)
AS

DECLARE @dtExpurgo DATETIME,
        @dtHoje    DATETIME

SET NOCOUNT OFF

SELECT @dtHoje = CONVERT(DATETIME, CONVERT(CHAR(10), GETDATE(), 103), 103)

SELECT @dtExpurgo = DATEADD(day, -1 * @dias, @dtHoje)

/*Movimento*/

/*tb_caixa_movto
  tb_operacao
  tb_pagamento
  tb_venda_combo
  tb_venda_ingresso
  tb_num_talao*/
  
BEGIN TRANSACTION 
                             
DELETE FROM tb_pagamento
WHERE tb_pagamento.ope_cd IN (SELECT tb_operacao.ope_cd
                              FROM tb_operacao
                              WHERE tb_operacao.cxp_cd IN (SELECT tb_caixa_movto.cxp_cd
                                                           FROM tb_caixa_movto
                                                           WHERE tb_caixa_movto.cxp_dt_abertura <= @dtExpurgo))
SELECT @Erro = @@ERROR
  
IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE tb_venda_combo
WHERE tb_venda_combo.ope_cd IN (SELECT tb_operacao.ope_cd
                                FROM tb_operacao
                                WHERE tb_operacao.cxp_cd IN (SELECT tb_caixa_movto.cxp_cd
                                                             FROM tb_caixa_movto
                                                             WHERE tb_caixa_movto.cxp_dt_abertura <= @dtExpurgo))

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_venda_ingresso
WHERE tb_venda_ingresso.ope_cd IN (SELECT tb_operacao.ope_cd
                                   FROM tb_operacao
                                   WHERE tb_operacao.cxp_cd IN (SELECT tb_caixa_movto.cxp_cd
                                                                FROM tb_caixa_movto
                                                                WHERE tb_caixa_movto.cxp_dt_abertura <= @dtExpurgo))

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_num_talao
WHERE tb_num_talao.ope_cd IN (SELECT tb_operacao.ope_cd
                              FROM tb_operacao
                              WHERE tb_operacao.cxp_cd IN (SELECT tb_caixa_movto.cxp_cd
                                                           FROM tb_caixa_movto
                                                           WHERE tb_caixa_movto.cxp_dt_abertura <= @dtExpurgo))

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_operacao
WHERE tb_operacao.cxp_cd IN (SELECT tb_caixa_movto.cxp_cd
                             FROM tb_caixa_movto
                             WHERE tb_caixa_movto.cxp_dt_abertura <= @dtExpurgo)

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_caixa_movto
WHERE tb_caixa_movto.cxp_dt_abertura <= @dtExpurgo


/*Precos*/

/*tb_prog_preco
  tb_preco*/
 
IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_preco
WHERE tb_preco.ppr_cd IN (SELECT tb_prog_preco.ppr_cd
                          FROM tb_prog_preco
                          WHERE tb_prog_preco.ppr_dt_fim IS NOT NULL
                          AND   tb_prog_preco.ppr_dt_fim <= @dtExpurgo)
  
IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_prog_preco
WHERE tb_prog_preco.ppr_dt_fim IS NOT NULL
AND   tb_prog_preco.ppr_dt_fim <= @dtExpurgo



/*Filmes/programacao*/

/*tb_filme
  tb_copia
  tb_sessao
  tb_programacao
  tb_sessao_real*/

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_sessao_real
WHERE tb_sessao_real.fil_cd IN (SELECT tb_filme.fil_cd
                                FROM tb_filme
                                WHERE tb_filme.fil_dt_fim IS NOT NULL
                                AND   tb_filme.fil_dt_fim <= @dtExpurgo)
  
IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_sessao
WHERE tb_sessao.fil_cd IN (SELECT tb_filme.fil_cd
                           FROM tb_filme
                           WHERE tb_filme.fil_dt_fim IS NOT NULL
                           AND   tb_filme.fil_dt_fim <= @dtExpurgo)

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_sessao
WHERE tb_sessao.prg_cd IN (SELECT tb_programacao.prg_cd
                           FROM tb_programacao
                           WHERE tb_programacao.prg_dt_fim IS NOT NULL
                           AND   tb_programacao.prg_dt_fim <= @dtExpurgo)

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_programacao
WHERE tb_programacao.prg_dt_fim IS NOT NULL
AND   tb_programacao.prg_dt_fim <= @dtExpurgo


IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_copia_filme
WHERE tb_copia_filme.fil_cd IN (SELECT tb_filme.fil_cd
                                FROM tb_filme
                                WHERE tb_filme.fil_dt_fim IS NOT NULL
                                AND   tb_filme.fil_dt_fim <= @dtExpurgo)
  
IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_filme
WHERE tb_filme.fil_dt_fim IS NOT NULL
AND   tb_filme.fil_dt_fim <= @dtExpurgo
  
IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

/*Programacao Combos*/  

/*tb_prog_combo*/
  
DELETE FROM tb_prog_combo
WHERE tb_prog_combo.pcb_dt_fim IS NOT NULL
AND   tb_prog_combo.pcb_dt_fim <= @dtExpurgo
  
IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

/*Outros*/  
/*tb_sala_lugar*/

DELETE FROM tb_sala_lugar
WHERE tb_sala_lugar.sal_dt_fim IS NOT NULL
AND   tb_sala_lugar.sal_dt_fim <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

/*Boletim/Bordero */

/*tb_bol_catraca
  tb_sis_log
  tb_bol_tp_ingr
  tb_bol_param
  tb_bol_filme
  tb_bol_distrib
  tb_bol_sala
  tb_bol_cin
  tb_bol_empr
  tb_boletim
  tb_bol_talao
  tb_bol_sessao
  tb_bol_pag_tp
  tb_bol_ope_tp
  tb_bol_ingre2
  tb_bol_ingre
  tb_ctrl_boletim
  tb_aux_boletim
  tb_bol_catraca_sala*/

DELETE FROM tb_sis_log
WHERE tb_sis_log.slg_data <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

--DELETE FROM tb_bol_ingre2
--WHERE tb_bol_ingre2.bol_dt_mov <= @dtExpurgo

DELETE FROM tb_bol_ingre
WHERE tb_bol_ingre.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_bol_talao
WHERE tb_bol_talao.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_bol_sessao
WHERE tb_bol_sessao.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_ctrl_boletim
WHERE tb_ctrl_boletim.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_aux_boletim
WHERE tb_aux_boletim.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_boletim
WHERE tb_boletim.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_bol_catraca_sala
WHERE tb_bol_catraca_sala.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_bol_catraca
WHERE tb_bol_catraca.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_bol_param
WHERE tb_bol_param.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_bol_filme
WHERE tb_bol_filme.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_bol_distrib
WHERE tb_bol_distrib.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_bol_sala
WHERE tb_bol_sala.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_bol_cin
WHERE tb_bol_cin.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_bol_empr
WHERE tb_bol_empr.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

COMMIT TRANSACTION 

GO


SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upExpurgoCentral') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upExpurgoCentral
GO


--****************************************************
CREATE PROCEDURE upExpurgoCentral
	(@dias   int,
         @Erro   int OUTPUT,
         @MsgErr varchar(255) OUTPUT)
AS

DECLARE @dtExpurgo DATETIME,
        @dtHoje    DATETIME


SELECT @dtHoje = CONVERT(DATETIME, CONVERT(CHAR(10), GETDATE(), 103), 103)

SELECT @dtExpurgo = DATEADD(day, (-1 * @dias), @dtHoje)

BEGIN TRANSACTION 

/*Boletim/Bordero */

/*tb_bol_catraca
  tb_sis_log
  tb_bol_tp_ingr
  tb_bol_param
  tb_bol_filme
  tb_bol_distrib
  tb_bol_sala
  tb_bol_cin
  tb_bol_empr
  tb_boletim
  tb_bol_talao
  tb_bol_sessao
  tb_bol_pag_tp
  tb_bol_ope_tp
  tb_bol_ingre2
  tb_bol_ingre
  tb_ctrl_boletim
  tb_aux_boletim
  tb_bol_catraca_sala*/

DELETE FROM tb_sis_log
WHERE tb_sis_log.slg_data <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

--DELETE FROM tb_bol_ingre2
--WHERE tb_bol_ingre2.bol_dt_mov <= @dtExpurgo

DELETE FROM tb_bol_ingre
WHERE tb_bol_ingre.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_bol_talao
WHERE tb_bol_talao.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_bol_sessao
WHERE tb_bol_sessao.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_ctrl_boletim
WHERE tb_ctrl_boletim.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_aux_boletim
WHERE tb_aux_boletim.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_boletim
WHERE tb_boletim.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_bol_catraca_sala
WHERE tb_bol_catraca_sala.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_bol_catraca
WHERE tb_bol_catraca.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_bol_param
WHERE tb_bol_param.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_bol_filme
WHERE tb_bol_filme.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_bol_distrib
WHERE tb_bol_distrib.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_bol_sala
WHERE tb_bol_sala.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_bol_cin
WHERE tb_bol_cin.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

DELETE FROM tb_bol_empr
WHERE tb_bol_empr.bol_dt_mov <= @dtExpurgo

IF @Erro <> 0
   BEGIN
      ROLLBACK TRANSACTION 
      
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro
         
      RETURN
   END

COMMIT TRANSACTION 

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upFECHAMENTO_CAIXA_VALOR') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upFECHAMENTO_CAIXA_VALOR
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upFECHAMENTO_CAIXA_BILHETE') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upFECHAMENTO_CAIXA_BILHETE
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upFECHAMENTO_CAIXA_COMBO') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upFECHAMENTO_CAIXA_COMBO
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upFECHAMENTO_CAIXA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upFECHAMENTO_CAIXA
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upFECHAMENTO_CAIXA1') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upFECHAMENTO_CAIXA1
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upFECHAMENTO_CAIXA_SELECAO') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upFECHAMENTO_CAIXA_SELECAO
GO
 
--****************************************************
CREATE PROCEDURE upFECHAMENTO_CAIXA_VALOR
   (@cxa_cd          int,
    @cxp_dt_abertura datetime,
    @Erro            int OUTPUT,
    @MsgErr          varchar(255) OUTPUT)
AS
   SELECT 0              AS 'pgt_cd', 
         'TROCO INICIAL' AS 'pgt_desc', 
         1               AS 'sinal', 
         d.ope_valor     AS 'pag_valor'
   FROM TB_CAIXA_MOVTO a INNER JOIN TB_OPERACAO d ON d.cxp_cd = a.cxp_cd
   WHERE d.opt_cd          = 2
   AND   a.cxa_cd          = @cxa_cd
   AND   a.cxp_dt_abertura = @cxp_dt_abertura
   AND   d.ope_dt_des IS NULL
   UNION
   SELECT f.pgt_cd                   AS 'pgt_cd',
          'TOTAL ' + f.pgt_desc      AS 'pgt_desc', 
          1                          AS 'sinal', 
          ISNULL(SUM(e.pag_valor),0) AS 'pag_valor'
   FROM ((TB_CAIXA_MOVTO a INNER JOIN TB_OPERACAO d ON a.cxp_cd = d.cxp_cd)
         INNER JOIN TB_PAGAMENTO e ON d.ope_cd = e.ope_cd)
        INNER JOIN TB_PAGAMENTO_TIPO f ON e.pgt_cd = f.pgt_cd
   WHERE a.cxa_cd          = @cxa_cd
   AND   a.cxp_dt_abertura = @cxp_dt_abertura
   AND   ((d.opt_cd IN (1, 3, 4)
   AND     d.ope_dt_des IS NULL)
   OR     (d.opt_cd IN (1)
   AND     d.ope_dt_des IS NOT NULL))
   GROUP BY f.pgt_cd, 
            f.pgt_desc
   UNION
   SELECT 10                         AS 'pgt_cd', 
          'SANGRIA (-)'              AS 'pgt_desc',
          -1                         AS 'sinal', 
          ISNULL(SUM(d.ope_valor),0) AS 'pag_valor'
   FROM (TB_CAIXA_MOVTO a INNER JOIN TB_OPERACAO d ON a.cxp_cd = d.cxp_cd)
        LEFT JOIN TB_PAGAMENTO e ON d.ope_cd = e.ope_cd
   WHERE d.opt_cd          = 5
   AND   a.cxa_cd          = @cxa_cd
   AND   a.cxp_dt_abertura = @cxp_dt_abertura
   AND   d.ope_dt_des IS NULL
   UNION
   SELECT 10                      AS 'pgt_cd', 
          'DEVOLUCAO (-)'         AS 'pgt_desc', 
          -1                      AS 'sinal', 
          ISNULL(SUM(a.valor), 0) AS 'pag_valor'
   FROM (SELECT e.ing_valor AS 'valor'
         FROM (TB_CAIXA_MOVTO a INNER JOIN TB_OPERACAO d ON a.cxp_cd = d.cxp_cd)
              INNER JOIN TB_VENDA_INGRESSO e ON d.ope_cd = e.ope_cd
         WHERE a.cxa_cd          = @cxa_cd
         AND   a.cxp_dt_abertura = @cxp_dt_abertura
         AND   e.ing_dt_canc IS NOT NULL
         UNION ALL
         SELECT e.vcb_qtde * e.vcb_valor AS 'valor'
         FROM (TB_CAIXA_MOVTO a INNER JOIN TB_OPERACAO d ON a.cxp_cd = d.cxp_cd)
              INNER JOIN TB_VENDA_COMBO e ON d.ope_cd = e.ope_cd
         WHERE a.cxa_cd          = @cxa_cd
         AND   a.cxp_dt_abertura = @cxp_dt_abertura
         AND   e.vcb_dt_canc IS NOT NULL) a
      
     SELECT @Erro = @@ERROR
   
     IF @Erro <> 0
        BEGIN
           SELECT @MsgErr = description
           FROM master..sysmessages
           WHERE error = @Erro
         
           RETURN
        END

GO

--****************************************************
CREATE PROCEDURE upFECHAMENTO_CAIXA
   (@cxa_cd       int,
    @cxp_dt_abertura  datetime,
         @Erro             int OUTPUT,
         @MsgErr           varchar(255) OUTPUT)
AS
   SELECT a.cxa_cd, 
          a.cxp_dt_abertura, 
          a.usu_abertura, 
          a.cxp_dt_fechamento, 
          a.usu_fechamento,
          b.usu_nm 'usu_nm_abertura', 
          c.usu_nm 'usu_nm_fechamento'
     FROM TB_CAIXA_MOVTO a,
          TB_USUARIO b,
          TB_USUARIO c
          --TB_OPERACAO d
    WHERE b.usu_cd = a.usu_abertura
      AND c.usu_cd = a.usu_fechamento
      AND a.cxa_cd = @cxa_cd
           AND a.cxp_dt_abertura = @cxp_dt_abertura
--    AND convert(char(11), a.cxp_dt_abertura, 103) + convert(char(8), a.cxp_dt_abertura, 108) = convert(char(11), @cxp_dt_abertura, 103) + convert(char(8), @cxp_dt_abertura, 108)
  
     SELECT @Erro = @@ERROR
   
     IF @Erro <> 0
        BEGIN
           SELECT @MsgErr = description
           FROM master..sysmessages
           WHERE error = @Erro
         
           RETURN
        END

GO

--****************************************************
CREATE PROCEDURE upFECHAMENTO_CAIXA_BILHETE
   (@cxa_cd       int,
    @cxp_dt_abertura  datetime,
         @Erro             int OUTPUT,
         @MsgErr           varchar(255) OUTPUT)
AS
   SELECT 0 'ordem', f.igt_cd, 'TOTAL ' + f.igt_desc 'igt_desc', count(1) 'qtde' --, ISNULL(SUM(e.ing_valor),0) ing_valor
     FROM TB_CAIXA_MOVTO a,
          TB_OPERACAO d,
          TB_VENDA_INGRESSO e,
          TB_INGRESSO_TIPO f
    WHERE d.cxp_cd = a.cxp_cd
      AND e.ope_cd = d.ope_cd
      AND f.igt_cd = e.igt_cd
      AND d.opt_cd = 1
      AND a.cxa_cd = @cxa_cd
      AND convert(char(10), e.sre_data, 103) = convert(char(10), a.cxp_dt_abertura, 103)
           AND a.cxp_dt_abertura = @cxp_dt_abertura
      AND d.ope_dt_des IS NULL
      AND e.ing_dt_canc IS NULL
   GROUP BY f.igt_cd, f.igt_desc
   UNION
   SELECT 1 'ordem', f.igt_cd, 'TOTAL ' + f.igt_desc + ' ANT' 'igt_desc', count(1) 'qtde' --, ISNULL(SUM(e.ing_valor),0) ing_valor
     FROM TB_CAIXA_MOVTO a,
          TB_OPERACAO d,
          TB_VENDA_INGRESSO e,
          TB_INGRESSO_TIPO f
    WHERE d.cxp_cd = a.cxp_cd
      AND e.ope_cd = d.ope_cd
      AND f.igt_cd = e.igt_cd
      AND d.opt_cd = 1
      AND a.cxa_cd = @cxa_cd
      AND convert(char(10), e.sre_data, 103) <> convert(char(10), a.cxp_dt_abertura, 103)
           AND a.cxp_dt_abertura = @cxp_dt_abertura
      AND d.ope_dt_des IS NULL
      AND e.ing_dt_canc IS NULL
   GROUP BY f.igt_cd, f.igt_desc
   
     SELECT @Erro = @@ERROR
   
     IF @Erro <> 0
        BEGIN
           SELECT @MsgErr = description
           FROM master..sysmessages
           WHERE error = @Erro
         
           RETURN
        END

GO

--****************************************************
CREATE PROCEDURE upFECHAMENTO_CAIXA_COMBO
   (@cxa_cd       int,
    @cxp_dt_abertura  datetime,
         @Erro             int OUTPUT,
         @MsgErr           varchar(255) OUTPUT)
AS
   SELECT TB_COMBO.cbo_cd, 
          TB_COMBO.cbo_nm,
          ISNULL(SUM(TB_VENDA_COMBO.vcb_qtde),0)  'qtde',
          ISNULL(SUM(TB_VENDA_COMBO.vcb_valor * TB_VENDA_COMBO.vcb_qtde),0) 'valor'
        FROM TB_CAIXA_MOVTO,
             TB_OPERACAO,
             TB_VENDA_COMBO,
             TB_COMBO
        WHERE TB_CAIXA_MOVTO.cxa_cd = @cxa_cd
        AND   TB_CAIXA_MOVTO.cxp_dt_abertura = @cxp_dt_abertura
        AND   TB_CAIXA_MOVTO.cxp_cd          = TB_OPERACAO.cxp_cd
        AND   TB_OPERACAO.ope_cd             = TB_VENDA_COMBO.ope_cd
        AND   TB_VENDA_COMBO.cbo_cd          = TB_COMBO.cbo_cd 
        AND   TB_OPERACAO.ope_dt_des IS NULL
        AND   TB_VENDA_COMBO.vcb_dt_canc IS NULL
        GROUP BY TB_COMBO.cbo_cd, TB_COMBO.cbo_nm
   
     SELECT @Erro = @@ERROR
   
     IF @Erro <> 0
        BEGIN
           SELECT @MsgErr = description
           FROM master..sysmessages
           WHERE error = @Erro
         
           RETURN
        END

GO

--****************************************************
CREATE PROCEDURE upFECHAMENTO_CAIXA_SELECAO
   (@cxp_dt_abertura  datetime,
         @Erro             int OUTPUT,
         @MsgErr           varchar(255) OUTPUT)
AS
   SELECT convert(bit,0), a.cxa_cd 'Caixa', 
          convert(char(11), a.cxp_dt_abertura, 103) + convert(char(8), a.cxp_dt_abertura, 108) 'Data Abertura',
          b.usu_nm 'Usuário Abertura', 
          convert(char(11), a.cxp_dt_fechamento, 103) + convert(char(8), a.cxp_dt_fechamento, 108) 'Data Fechamento',
          c.usu_nm 'Usuário Fechamento'
     FROM TB_CAIXA_MOVTO a,
          TB_USUARIO b,
          TB_USUARIO c
    WHERE b.usu_cd = a.usu_abertura
      AND c.usu_cd = a.usu_fechamento
      AND convert(char(11), a.cxp_dt_abertura, 103) = convert(char(11), @cxp_dt_abertura, 103)
   ORDER BY a.cxp_dt_abertura DESC
   
     SELECT @Erro = @@ERROR
   
     IF @Erro <> 0
        BEGIN
           SELECT @MsgErr = description
           FROM master..sysmessages
           WHERE error = @Erro
         
           RETURN
        END

GO

--****************************************************
CREATE PROCEDURE upFECHAMENTO_CAIXA1
   (@cxa_cd       int,
    @cxp_dt_abertura  datetime,
         @Erro             int OUTPUT,
         @MsgErr           varchar(255) OUTPUT)
AS
   SELECT a.cxa_cd, 
          a.cxp_dt_abertura, 
          a.usu_abertura, 
          b.usu_nm 'usu_nm_abertura'
     FROM TB_CAIXA_MOVTO a,
          TB_USUARIO b
    WHERE b.usu_cd = a.usu_abertura
      AND a.cxa_cd = @cxa_cd
           AND a.cxp_dt_abertura = @cxp_dt_abertura
  
     SELECT @Erro = @@ERROR
   
     IF @Erro <> 0
        BEGIN
           SELECT @MsgErr = description
           FROM master..sysmessages
           WHERE error = @Erro
         
           RETURN
        END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upDATA_SERVIDOR') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upDATA_SERVIDOR
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upHORA_SERVIDOR') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upHORA_SERVIDOR
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufDiaSemana]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufDiaSemana]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[upPeriodoDia]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[upPeriodoDia]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[upDataRef]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[upDataRef]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufCNPJ]')  and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufCNPJ]
GO

--****************************************************
CREATE PROCEDURE upDATA_SERVIDOR
AS
    SELECT CONVERT(char(10),GETDATE(),103) 'data'
GO

--****************************************************
CREATE PROCEDURE upHORA_SERVIDOR
AS
    SELECT CONVERT(char(10),GETDATE(),108) 'hora'
GO


--****************************************************
CREATE  FUNCTION ufDiaSemana (@DataRef datetime)
RETURNS int
AS
   BEGIN	
      DECLARE @diaSemana smallint

      SELECT @diaSemana = 8
      FROM tb_feriado
      WHERE fer_data = convert(datetime, @DataRef, 103)

      IF @diaSemana IS NULL
         SELECT @diaSemana = datepart(dw, convert(datetime, @DataRef, 103))

      RETURN (@diaSemana)
   END	
GO

--****************************************************
CREATE PROCEDURE upPeriodoDia(@DataRef    datetime,
                              @DataIniPer datetime  OUTPUT,
                              @DataFimPer datetime  OUTPUT)
AS

   DECLARE @HoraMaxSes  datetime

   SELECT @HoraMaxSes = par_hora_max_ses
   FROM tb_parametro
     
   SELECT @DataIniPer = convert(datetime, convert(char(10),@DataRef,103) + ' ' + convert(char(10),@HoraMaxSes,108), 103)
   SELECT @DataFimPer = convert(datetime, convert(char(10),DATEADD(Day, 1, @DataRef),103) + ' ' + convert(char(10),@HoraMaxSes,108), 103)

GO

--****************************************************
CREATE PROCEDURE upDataRef(@DataRef datetime OUTPUT)
AS

   DECLARE @HoraMaxSes datetime,
           @dtRef      datetime,
           @DataAux    datetime

   SELECT @HoraMaxSes = par_hora_max_ses
   FROM tb_parametro
   
   SELECT @dtRef = GETDATE()
     
   SELECT @DataAux = convert(datetime, convert(char(10),@dtRef,103) + ' ' + convert(char(10),@HoraMaxSes,108), 103)
   
   IF @dtRef > @DataAux 
      SELECT @DataRef = convert(datetime, convert(char(10), @dtRef, 103), 103)
   ELSE
      SELECT @DataRef = DATEADD(day, -1, convert(datetime, convert(char(10), @DataAux, 103), 103))

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

--****************************************************
CREATE FUNCTION ufCNPJ (@CNPJ VARCHAR(18))
RETURNS VARCHAR(18)
AS
   BEGIN	
      DECLARE @ret CHAR(18)
      
      SELECT @ret = SUBSTRING(@CNPJ,1,2) + '.' + SUBSTRING(@CNPJ,3,3) + '.' + SUBSTRING(@CNPJ,6,3) + '/' + SUBSTRING(@CNPJ,9,4) + '-' + SUBSTRING(@CNPJ,13,2)
      
      RETURN @ret
   END	
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CAIXA_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CAIXA_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CAIXA_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CAIXA_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CAIXA_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CAIXA_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CAIXA_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CAIXA_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CAIXA_GRID') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CAIXA_GRID
GO

--****************************************************
CREATE PROCEDURE upTB_CAIXA_D
	(@cxa_cd 	int,
         @TipoExclusao  char(1),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
    DECLARE @n int

    SELECT @n = COUNT(*)
    FROM TB_CAIXA_MOVTO
    WHERE cxa_cd = @cxa_cd
    AND   cxp_status <> 2
    
    IF @n > 0 
      BEGIN
         SELECT @MsgErr = 'Este caixa esta aberto'
         SELECT @Erro   = 99
         
         RETURN
      END

    IF @TipoExclusao = 'L'
	 UPDATE TB_CAIXA 
	    SET cxa_dt_des = GETDATE()
	  WHERE cxa_cd	= @cxa_cd
    ELSE
	DELETE TB_CAIXA 
	 WHERE cxa_cd = @cxa_cd
	 
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CAIXA_GRID
	(@Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   SELECT a.cxa_cd, b.cin_cd, b.cin_nm 'Cinema', a.cxa_desc 'Caixa'
     FROM TB_CAIXA a,
          TB_CINEMA b
    WHERE a.cin_cd = b.cin_cd
      AND a.cxa_dt_des IS NULL
      AND b.cin_dt_des IS NULL
    ORDER BY cxa_cd  
      
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE  PROCEDURE upTB_CAIXA_I
	(@cin_cd 	int,
         @cxa_desc 	varchar(50),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
INSERT INTO TB_CAIXA 
	 ( cin_cd,
         cxa_desc,
 	 cxa_dt_inc) 
VALUES 
	( @cin_cd,
         @cxa_desc,
	 GETDATE())
	 
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CAIXA_S
	(@cxa_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   IF @cxa_cd IS NULL
       SELECT a.*
         FROM TB_CAIXA a,
              TB_CINEMA b
        WHERE a.cin_cd = b.cin_cd
          AND a.cxa_dt_des IS NULL
          AND b.cin_dt_des IS NULL
   ELSE
      SELECT * 
        FROM TB_CAIXA
       WHERE cxa_cd = @cxa_cd
       
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CAIXA_U
	(@cxa_cd 	int,
	 @cin_cd 	int,
         @cxa_desc 	varchar(50),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
 UPDATE TB_CAIXA 
    SET cin_cd = @cin_cd,
        cxa_desc = @cxa_desc,
	cxa_dt_alt = GETDATE()
  WHERE cxa_cd	= @cxa_cd
  
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CAIXA_MOVTO_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CAIXA_MOVTO_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CAIXA_MOVTO_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CAIXA_MOVTO_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CAIXA_MOVTO_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CAIXA_MOVTO_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CAIXA_MOVTO_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CAIXA_MOVTO_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CAIXA_MOVTO_GRID') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CAIXA_MOVTO_GRID
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CAIXA_MOVTO_SELECAO') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CAIXA_MOVTO_SELECAO
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CAIXA_MOVTO_STATUS') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CAIXA_MOVTO_STATUS
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CAIXA_MOVTO_ABRE') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CAIXA_MOVTO_ABRE
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CAIXA_MOVTO_FECHA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CAIXA_MOVTO_FECHA
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CAIXA_MOVTO_SUSPENDE') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CAIXA_MOVTO_SUSPENDE
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CAIXA_MOVTO_DATA_ABERTURA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CAIXA_MOVTO_DATA_ABERTURA
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CAIXA_MOVTO_REABRE') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CAIXA_MOVTO_REABRE
GO

--****************************************************
CREATE PROCEDURE upTB_CAIXA_MOVTO_ABRE
	(@cxa_cd 		int,
	 @usu_abertura		int,
	 @cxp_talao		bit,
	 @cxp_cd                int OUTPUT,
         @Erro       		int OUTPUT,
         @MsgErr      		varchar(255) OUTPUT)
AS 
DECLARE @data datetime
SELECT @data = convert(datetime,convert(char(11),GETDATE())) + convert(datetime,convert(char(8),GETDATE(),108))
INSERT INTO TB_CAIXA_MOVTO 
	 (cxa_cd,
	  cxp_dt_abertura,
 	  usu_abertura,
	  cxp_status,
	  cxp_talao )
VALUES 
	(@cxa_cd,
	 @data,
 	 @usu_abertura,
	 0,
	 @cxp_talao)
	 
   SELECT @cxp_cd = @@IDENTITY
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CAIXA_MOVTO_D
	(@cxp_cd                int,
         @Erro          	int OUTPUT,
         @MsgErr        	varchar(255) OUTPUT)
AS
   DELETE TB_CAIXA_MOVTO 
    WHERE cxp_cd = @cxp_cd
    
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CAIXA_MOVTO_DATA_ABERTURA
	(@cxa_cd 		int,
	 @Erro         		int OUTPUT,
         @MsgErr        	varchar(255) OUTPUT)
AS
    SELECT cxp_dt_abertura
      FROM TB_CAIXA_MOVTO
     WHERE cxa_cd = @cxa_cd
       AND cxp_dt_fechamento IS NULL
       
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CAIXA_MOVTO_FECHA
	(@cxp_cd         int,
	 @usu_fechamento int,
         @Erro           int OUTPUT,
         @MsgErr         varchar(255) OUTPUT)
AS 
DECLARE @data datetime
SELECT @data = convert(datetime,convert(char(11),GETDATE())) + convert(datetime,convert(char(8),GETDATE(),108))
 UPDATE TB_CAIXA_MOVTO 
    SET cxp_dt_fechamento = @data,
	usu_fechamento    = @usu_fechamento,
	cxp_status        = 2
  WHERE cxp_cd = @cxp_cd
  
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CAIXA_MOVTO_GRID
	(@Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   SELECT a.cxa_cd, 
          a.cxp_status,
          b.cxa_desc 'Caixa', 
          a.cxp_dt_abertura 'Abertura', 
          a.usu_abertura 'Usuário Abriu', 
          a.cxp_dt_fechamento 'Fechamento', 
          a.usu_fechamento 'Usuário Fechou', 
          case when a.cxp_status = 0 then
		'Caixa em Operação'
	  else
		'Caixa Suspenso'
	  end 'Status do Caixa',
          case when a.cxp_talao = 0 then
		'normal'
	  else
		'TALÃO'
	  end 'Tipo',
	  a.cxp_cd
     FROM TB_CAIXA_MOVTO a,
          TB_CAIXA b
    WHERE a.cxa_cd = b.cxa_cd
      AND b.cxa_dt_des IS NULL
      
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CAIXA_MOVTO_I
	(@cxa_cd 		int,
	 @usu_abertura		int,
	 @cxp_dt_fechamento	datetime,
	 @usu_fechamento	int,
	 @cxp_status		smallint,
	 @cxp_talao		bit,
	 @cxp_cd                int OUTPUT,
         @Erro       		int OUTPUT,
         @MsgErr      		varchar(255) OUTPUT)
AS 
DECLARE @data datetime
SELECT @data = convert(datetime,convert(char(11),GETDATE())) + convert(datetime,convert(char(8),GETDATE(),108))
INSERT INTO TB_CAIXA_MOVTO 
	 (cxa_cd,
	  cxp_dt_abertura,
 	  usu_abertura,
	  cxp_dt_fechamento,
	  usu_fechamento, 
	  cxp_status,
	  cxp_talao)
VALUES 
	(@cxa_cd,
	 @data,
 	 @usu_abertura,
	 @cxp_dt_fechamento,
	 @usu_fechamento,
	 @cxp_status,
	 @cxp_talao)
   
   SELECT @cxp_cd = @@IDENTITY
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CAIXA_MOVTO_REABRE
	(@cxp_cd int,
         @Erro   int OUTPUT,
         @MsgErr varchar(255) OUTPUT)
AS 
 UPDATE TB_CAIXA_MOVTO 
    SET cxp_status = 0
  WHERE cxp_cd = @cxp_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CAIXA_MOVTO_S
	(@cxa_cd 		int,
	 @cxp_dt_abertura	datetime,
         @Erro         		int OUTPUT,
         @MsgErr        	varchar(255) OUTPUT)
AS
   IF @cxa_cd IS NULL AND @cxp_dt_abertura IS NULL
       SELECT *
         FROM TB_CAIXA_MOVTO
   ELSE
       IF @cxa_cd IS NOT NULL AND @cxp_dt_abertura IS NULL
          SELECT * 
            FROM TB_CAIXA_MOVTO
           WHERE cxa_cd = @cxa_cd
       ELSE
           IF @cxa_cd IS NOT NULL AND @cxp_dt_abertura IS NOT NULL
              SELECT * 
                FROM TB_CAIXA_MOVTO
               WHERE cxa_cd = @cxa_cd
                 AND cxp_dt_abertura = @cxp_dt_abertura
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CAIXA_MOVTO_SELECAO
	(@Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   SELECT b.cxa_cd, 
          a.usu_abertura, 
          a.cxp_status,
          b.cxa_desc 'Caixa', 
          c.usu_nm 'Usuário',
          a.cxp_dt_abertura 'Data Abertura', 
          case when a.cxp_status = 0 then
		'Caixa em Operação'
               when a.cxp_status = 1 then
		'Caixa Suspenso'
	  else
		'Caixa Livre'
	  end 'Status do Caixa',
          case when a.cxp_talao = 0 then
		'normal'
	  else
		'TALÃO'
	  end 'Tipo',
	  a.cxp_cd
     FROM TB_CAIXA_MOVTO a,
          TB_CAIXA b,
          TB_USUARIO c
    WHERE b.cxa_cd = a.cxa_cd
      AND c.usu_cd = a.usu_abertura
      AND b.cxa_dt_des IS NULL
      AND a.cxp_status <> 2
   UNION
   SELECT cxa_cd, null 'usu_abertura', null 'cxp_status',
          cxa_desc 'Caixa', 
          null 'Usuário' ,
          null 'Data Abertura', 
          'Caixa Livre' 'Status do Caixa',
          null 'Tipo',
	  null 'cxp_cd'
     FROM TB_CAIXA 
    WHERE cxa_dt_des IS NULL
      AND cxa_cd NOT IN ( SELECT cxa_cd 
                            FROM TB_CAIXA_MOVTO 
                           WHERE cxp_status <> 2 ) 
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CAIXA_MOVTO_STATUS
	(@usu_abertura		int,
	 @cxp_status		smallint,
         @Erro         		int OUTPUT,
         @MsgErr        	varchar(255) OUTPUT)
AS
   IF @usu_abertura IS NOT NULL AND @cxp_status IS NOT NULL
      SELECT * 
        FROM TB_CAIXA_MOVTO
       WHERE usu_abertura = @usu_abertura
         AND cxp_status = @cxp_status
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CAIXA_MOVTO_SUSPENDE
	(@cxp_cd     int,
	 @cxp_status smallint,
	 @Erro       int OUTPUT,
         @MsgErr     varchar(255) OUTPUT)
AS 
 UPDATE TB_CAIXA_MOVTO 
    SET cxp_status = @cxp_status
  WHERE cxp_cd = @cxp_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CAIXA_MOVTO_U
	(@cxp_cd                int,
	 @usu_abertura		int,
	 @cxp_dt_fechamento	datetime,
	 @usu_fechamento	int,
	 @cxp_status		smallint,
	 @cxp_talao		bit,
         @Erro          	int OUTPUT,
         @MsgErr        	varchar(255) OUTPUT)
AS 
 UPDATE TB_CAIXA_MOVTO 
    SET usu_abertura      = @usu_abertura,
	cxp_dt_fechamento = @cxp_dt_fechamento,
	usu_fechamento    = @usu_fechamento,
	cxp_status        = @cxp_status,
	cxp_talao         = @cxp_talao
  WHERE cxp_cd = @cxp_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO


SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CATRACA_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CATRACA_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CATRACA_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CATRACA_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CATRACA_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CATRACA_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CATRACA_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CATRACA_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CATRACA_GRID') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CATRACA_GRID
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CATRACA_SALA_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CATRACA_SALA_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CATRACA_SALA_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CATRACA_SALA_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CATRACA_SALA_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CATRACA_SALA_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CATRACA_CONTADOR') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CATRACA_CONTADOR
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CATRACA_CONTADOR_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CATRACA_CONTADOR_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CATRACA_CONTADOR_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CATRACA_CONTADOR_S
GO

--****************************************************
CREATE PROCEDURE upTB_CATRACA_CONTADOR_S
        (@cat_cd       int,
         @ctc_fim_cont int OUTPUT,
         @Erro         int OUTPUT,
         @MsgErr       varchar(255) OUTPUT)
AS 
   declare @data1  datetime

   SELECT @data1 = MAX(ctc_dt)
   FROM TB_CATRACA_CONT
   WHERE cat_cd = @cat_cd

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

   SELECT @ctc_fim_cont = ctc_fim_cont
   FROM TB_CATRACA_CONT
   WHERE cat_cd = @cat_cd
   AND   ctc_dt = @data1

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CATRACA_CONTADOR_U
        (@cat_cd       int,
         @ctc_fim_cont int,
         @Erro         int OUTPUT,
         @MsgErr       varchar(255) OUTPUT)
AS 
   declare @data1  datetime

   SELECT @data1 = MAX(ctc_dt)
   FROM TB_CATRACA_CONT
   WHERE cat_cd = @cat_cd

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

   UPDATE TB_CATRACA_CONT
   SET ctc_fim_cont = @ctc_fim_cont 
   WHERE cat_cd = @cat_cd
   AND   ctc_dt = @data1

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CATRACA_CONTADOR
        (@cat_cd    int,
         @Erro      int OUTPUT,
         @MsgErr    varchar(255) OUTPUT)
AS 
   declare @par_hora_max_ses datetime,
           @data1        datetime,
           @data2        datetime,
           @ctc_ini_cont int,
           @ctc_fim_cont int,
           @contAux      int
           
   SELECT @par_hora_max_ses = par_hora_max_ses
   FROM TB_PARAMETRO
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END
      
   SELECT @data1 = GETDATE()
   SELECT @data2 = CONVERT(datetime, CONVERT(char(8), @data1, 108), 108)
   SELECT @data1 = CONVERT(datetime, CONVERT(CHAR(10), @data1, 103), 103)
   
   IF @data2 <= @par_hora_max_ses 
      SELECT @data1 = DATEADD(Day, -1, @data1)
      
   SELECT @ctc_ini_cont = ctc_ini_cont,
          @ctc_fim_cont = ctc_fim_cont
   FROM TB_CATRACA_CONT
   WHERE cat_cd = @cat_cd
   AND   ctc_dt = @data1
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END
   
   IF @ctc_ini_cont IS NULL
      BEGIN
         SELECT @data2 = MAX(ctc_dt)
         FROM TB_CATRACA_CONT
         WHERE cat_cd = @cat_cd
         
         SELECT @Erro = @@ERROR
    
         IF @Erro <> 0
            BEGIN
               SELECT @MsgErr = description
               FROM master..sysmessages
               WHERE error = @Erro
         
               RETURN
            END

         SELECT @ctc_ini_cont = ctc_ini_cont,
                @ctc_fim_cont = ctc_fim_cont
         FROM TB_CATRACA_CONT
         WHERE cat_cd = @cat_cd
         AND   ctc_dt = @data2
      
         IF @ctc_fim_cont IS NOT NULL
            BEGIN
               IF @ctc_fim_cont + 1 > 99999
                  SELECT @contAux = 1
               ELSE
                  SELECT @contAux = @ctc_fim_cont + 1
            
               INSERT INTO TB_CATRACA_CONT
               (cat_cd, ctc_dt, ctc_ini_cont, ctc_fim_cont)
               VALUES (@cat_cd, @data1, @contAux, @contAux)
            END
         ELSE
            BEGIN
               INSERT INTO TB_CATRACA_CONT
               (cat_cd, ctc_dt, ctc_ini_cont, ctc_fim_cont)
               VALUES (@cat_cd, @data1, 1, 1)
            END
      END
   ELSE
      BEGIN
         UPDATE TB_CATRACA_CONT
         SET ctc_fim_cont = CASE 
                               WHEN ctc_fim_cont + 1 > 99999 THEN 1
                               ELSE ctc_fim_cont + 1
                            END
         WHERE cat_cd = @cat_cd
         AND   ctc_dt = @data1
      
      SELECT @Erro = @@ERROR
   
      IF @Erro <> 0
         BEGIN
            SELECT @MsgErr = description
            FROM master..sysmessages
            WHERE error = @Erro
         
            RETURN
         END
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CATRACA_D
   (@cat_cd    int,
         @TipoExclusao  char(1),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
    IF @TipoExclusao = 'L'
       UPDATE TB_CATRACA 
       SET cat_dt_des = GETDATE()
       WHERE cat_cd = @cat_cd
       
       SELECT @Erro = @@ERROR
   
       IF @Erro <> 0
          BEGIN
             SELECT @MsgErr = description
             FROM master..sysmessages
             WHERE error = @Erro
         
             RETURN
          END
    ELSE
      BEGIN
         BEGIN TRANSACTION
         
         DELETE TB_CATRACA_SALA
         WHERE cat_cd = @cat_cd
         
         SELECT @Erro = @@ERROR
   
         IF @Erro <> 0
            BEGIN
               SELECT @MsgErr = description
               FROM master..sysmessages
               WHERE error = @Erro
         
               ROLLBACK TRANSACTION
         
               RETURN
            END
 
         DELETE TB_CATRACA_CONT
         WHERE cat_cd = @cat_cd
         
         SELECT @Erro = @@ERROR
   
         IF @Erro <> 0
            BEGIN
               SELECT @MsgErr = description
               FROM master..sysmessages
               WHERE error = @Erro
         
               ROLLBACK TRANSACTION
         
               RETURN
            END
       
         DELETE TB_CATRACA 
         WHERE cat_cd = @cat_cd
         
         SELECT @Erro = @@ERROR
   
         IF @Erro <> 0
            BEGIN
               SELECT @MsgErr = description
               FROM master..sysmessages
               WHERE error = @Erro
         
               ROLLBACK TRANSACTION
         
               RETURN
            END
            
         COMMIT TRANSACTION    
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CATRACA_GRID
   (@Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   SELECT b.cin_cd, a.cat_cd, b.cin_nm 'Cinema', a.cat_nm 'Nome'
     FROM TB_CATRACA a,
     TB_CINEMA b
    WHERE a.cin_cd = b.cin_cd
      AND a.cat_dt_des IS NULL
      AND b.cin_dt_des IS NULL
      
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CATRACA_I
   (@cat_cd  int,
    @cin_cd  int,
    @cat_nm  varchar(50),
    @Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS 
   declare @catCd int
   
   SELECT @catCd = cat_cd
   FROM TB_CATRACA
   WHERE cat_cd = @cat_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END
      
   IF @catCd IS NOT NULL   
      BEGIN
         SELECT @Erro   = 1
         SELECT @MsgErr = 'Já existe catraca com este número'
         
         RETURN
      END
   INSERT INTO TB_CATRACA 
       (cat_cd,
        cin_cd,
        cat_nm,
        cat_dt_inc) 
   VALUES 
      (@cat_cd,
       @cin_cd,
       @cat_nm,
       GETDATE())
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO


--****************************************************
CREATE PROCEDURE upTB_CATRACA_S
(@cat_cd    int,
     @Erro          int OUTPUT,
     @MsgErr        varchar(255) OUTPUT)
AS
	IF @cat_cd IS NULL
		SELECT *
		  FROM TB_CATRACA 
		 WHERE cat_dt_des IS NULL
	ELSE
	  SELECT * 
		 FROM TB_CATRACA
		WHERE cat_cd = @cat_cd
	
	SELECT @Erro = @@ERROR

	IF @Erro <> 0
	  BEGIN
		  SELECT @MsgErr = description
		  FROM master..sysmessages
		  WHERE error = @Erro

		  RETURN
	  END

GO

--****************************************************
CREATE PROCEDURE upTB_CATRACA_SALA_D
   (@cat_cd    int,
    @sal_cd    int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   IF @sal_cd IS NULL
     DELETE TB_CATRACA_SALA
     WHERE cat_cd = @cat_cd
   ELSE
     DELETE TB_CATRACA_SALA
     WHERE cat_cd = @cat_cd
     AND   sal_cd = @sal_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CATRACA_SALA_I
   (@cat_cd    int,
         @sal_cd  int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
INSERT INTO TB_CATRACA_SALA
    (cat_cd,
     sal_cd) 
VALUES 
   (@cat_cd,
    @sal_cd)
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CATRACA_SALA_S
   (@cat_cd    int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   SELECT TB_CATRACA_SALA.sal_cd,
          TB_SALA.sal_desc
   FROM TB_CATRACA_SALA,
        TB_SALA
   WHERE TB_CATRACA_SALA.cat_cd = @cat_cd
   AND   TB_CATRACA_SALA.sal_cd = TB_SALA.sal_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CATRACA_U
   (@cat_cd    int,
         @cin_cd  int,
         @cat_nm  varchar(50),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
 UPDATE TB_CATRACA 
    SET cat_nm = @cat_nm,
   cin_cd = @cin_cd,
   cat_dt_alt = GETDATE()
  WHERE cat_cd = @cat_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CINEMA_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CINEMA_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CINEMA_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CINEMA_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CINEMA_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CINEMA_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_CINEMA_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_CINEMA_S
GO

--****************************************************
CREATE PROCEDURE upTB_CINEMA_D
	(@cin_cd 	int,
         @TipoExclusao  char(1),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
    IF @TipoExclusao = 'L'
	 UPDATE TB_CINEMA 
	    SET cin_dt_des = GETDATE()
	  WHERE cin_cd	= @cin_cd
    ELSE
	DELETE TB_CINEMA 
	 WHERE cin_cd = @cin_cd
	 
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CINEMA_I
	(@cin_cd 	int,
	 @emp_cd 	int,
         @cin_nm 	varchar(50),
	 @cin_cnpj 	char(14),
	 @cin_inscricao	char(12),
	 @cin_end 	varchar(50),
	 @cin_num_end 	int,
	 @cin_cmp_end 	varchar(20),
	 @cin_brr_end 	varchar(50),
	 @cin_cid_end 	varchar(50),
	 @cin_uf_end 	varchar(2),
	 @cin_cep_end 	char(8),
	 @cin_tel	varchar(20),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
INSERT INTO TB_CINEMA 
	 (cin_cd,
	  emp_cd,
          cin_nm,
	  cin_cnpj,
	  cin_inscricao,
	  cin_end,
	  cin_num_end,
	  cin_cmp_end,
	  cin_brr_end,
	  cin_cid_end,
	  cin_uf_end,
	  cin_cep_end,
	  cin_tel,
 	  cin_dt_inc) 
VALUES 
	(@cin_cd,
	 @emp_cd,
         @cin_nm,
	 @cin_cnpj,
	 @cin_inscricao,
	 @cin_end,
	 @cin_num_end,
	 @cin_cmp_end,
	 @cin_brr_end,
	 @cin_cid_end,
	 @cin_uf_end,
	 @cin_cep_end,
	 @cin_tel,
	 GETDATE())
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CINEMA_S
	(@cin_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   IF @cin_cd IS NULL
       SELECT a.*
         FROM TB_CINEMA a,
              TB_EMPRESA b
        WHERE a.emp_cd = b.emp_cd
          AND a.cin_dt_des IS NULL
          AND b.emp_dt_des IS NULL
   ELSE
      SELECT * 
        FROM TB_CINEMA
       WHERE cin_cd = @cin_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_CINEMA_U
	(@cin_cd 	int,
	 @emp_cd 	int,
         @cin_nm 	varchar(50),
	 @cin_cnpj 	char(14),
	 @cin_inscricao	char(12),
	 @cin_end 	varchar(50),
	 @cin_num_end 	int,
	 @cin_cmp_end 	varchar(20),
	 @cin_brr_end 	varchar(50),
	 @cin_cid_end 	varchar(50),
	 @cin_uf_end 	varchar(2),
	 @cin_cep_end 	char(8),
	 @cin_tel 	varchar(20),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
 UPDATE TB_CINEMA 
    SET emp_cd = @emp_cd,
        cin_nm = @cin_nm,
	cin_cnpj = @cin_cnpj,
	cin_inscricao = @cin_inscricao,
	cin_end = @cin_end,
	cin_num_end = @cin_num_end,
	cin_cmp_end = @cin_cmp_end,
	cin_brr_end = @cin_brr_end,
	cin_cid_end = @cin_cid_end,
	cin_uf_end = @cin_uf_end,
	cin_cep_end = @cin_cep_end,
	cin_dt_alt = GETDATE(),
	cin_tel = @cin_tel
  WHERE cin_cd	= @cin_cd
if @@ROWCOUNT = 0 
   BEGIN
      EXEC upTB_CINEMA_I
         @cin_cd,
         @emp_cd,
         @cin_nm,
	 @cin_cnpj,
	 @cin_inscricao,
	 @cin_end,
	 @cin_num_end,
	 @cin_cmp_end,
	 @cin_brr_end,
	 @cin_cid_end,
	 @cin_uf_end,
	 @cin_cep_end,
	 @cin_tel,
         @Erro,
         @MsgErr
   END
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_COMBO_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_COMBO_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_COMBO_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_COMBO_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_COMBO_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_COMBO_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_COMBO_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_COMBO_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_COMBO_GRID') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_COMBO_GRID
GO

--****************************************************
CREATE PROCEDURE upTB_COMBO_D
	(@cbo_cd 	int,
         @TipoExclusao  char(1),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
    IF @TipoExclusao = 'L'
	 UPDATE TB_COMBO 
	    SET cbo_dt_des = GETDATE()
	  WHERE cbo_cd	= @cbo_cd
    ELSE
	DELETE TB_COMBO 
	 WHERE cbo_cd = @cbo_cd

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_COMBO_GRID
	(@Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   SELECT cbo_cd, cbo_nm 'Combo', cbo_desc 'Descrição'
     FROM TB_COMBO
    WHERE cbo_dt_des IS NULL
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_COMBO_I
	(@cbo_nm 	varchar(20),
         @cbo_desc 	varchar(255),
	 @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
INSERT INTO TB_COMBO 
	 ( cbo_nm,
	 cbo_desc,
 	 cbo_dt_inc) 
VALUES 
	( @cbo_nm,
	 @cbo_desc,
	 GETDATE())
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_COMBO_S
	(@cbo_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   IF @cbo_cd IS NULL
       SELECT *
         FROM TB_COMBO
        WHERE cbo_dt_des IS NULL
   ELSE
      SELECT * 
        FROM TB_COMBO
       WHERE cbo_cd = @cbo_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_COMBO_U
	(@cbo_cd 	int,
	 @cbo_nm 	varchar(20),
         @cbo_desc 	varchar(255),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
 UPDATE TB_COMBO 
    SET cbo_desc = @cbo_desc,
	cbo_dt_alt = GETDATE()
  WHERE cbo_cd	= @cbo_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_DISTRIBUIDORA_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_DISTRIBUIDORA_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_DISTRIBUIDORA_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_DISTRIBUIDORA_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_DISTRIBUIDORA_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_DISTRIBUIDORA_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_DISTRIBUIDORA_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_DISTRIBUIDORA_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_DISTRIBUIDORA_GRID') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_DISTRIBUIDORA_GRID
GO

--****************************************************
CREATE PROCEDURE upTB_DISTRIBUIDORA_D
	(@dis_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   DELETE TB_DISTRIBUIDORA 
   WHERE dis_cd = @dis_cd

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_DISTRIBUIDORA_GRID
	(@Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   SELECT dis_cd 'Código', dis_nm 'Distribuidora'
     FROM TB_DISTRIBUIDORA
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_DISTRIBUIDORA_I
	(@dis_cd 	int,
         @dis_nm 	varchar(50),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
INSERT INTO TB_DISTRIBUIDORA
	 (dis_cd,
          dis_nm) 
VALUES 	(@dis_cd,
         @dis_nm)
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_DISTRIBUIDORA_S
	(@dis_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   IF @dis_cd IS NULL
       SELECT *
         FROM TB_DISTRIBUIDORA
   ELSE
      SELECT * 
        FROM TB_DISTRIBUIDORA
       WHERE dis_cd = @dis_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_DISTRIBUIDORA_U
	(@dis_cd 	int,
         @dis_nm 	varchar(50),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
 UPDATE TB_DISTRIBUIDORA 
    SET dis_nm = @dis_nm
  WHERE dis_cd = @dis_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_EMPRESA_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_EMPRESA_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_EMPRESA_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_EMPRESA_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_EMPRESA_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_EMPRESA_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_EMPRESA_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_EMPRESA_S
GO

--****************************************************
CREATE PROCEDURE upTB_EMPRESA_D
	(@emp_cd 	int,
         @TipoExclusao  char(1),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
    IF @TipoExclusao = 'L'
	 UPDATE TB_EMPRESA 
	    SET emp_dt_des = GETDATE()
	  WHERE emp_cd	= @emp_cd
    ELSE
	DELETE TB_EMPRESA 
	 WHERE emp_cd = @emp_cd

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_EMPRESA_I
	(@emp_cd 	int,
	 @emp_nm 	varchar(50),
	 @emp_cnpj 	char(14),
	 @emp_inscricao	char(12),
	 @emp_end 	varchar(50),
	 @emp_num_end 	int,
	 @emp_cmp_end 	varchar(20),
	 @emp_brr_end 	varchar(50),
	 @emp_cid_end 	varchar(50),
	 @emp_uf_end 	varchar(2),
	 @emp_cep_end 	char(8),
	 @emp_tel	varchar(20),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
INSERT INTO TB_EMPRESA 
	 (emp_cd,
	  emp_nm,
	  emp_cnpj,
	  emp_inscricao,
	  emp_end,
	  emp_num_end,
	  emp_cmp_end,
	  emp_brr_end,
	  emp_cid_end,
	  emp_uf_end,
	  emp_cep_end,
	  emp_tel,
 	  emp_dt_inc) 
VALUES 
	(@emp_cd,
	 @emp_nm,
	 @emp_cnpj,
	 @emp_inscricao,
	 @emp_end,
	 @emp_num_end,
	 @emp_cmp_end,
	 @emp_brr_end,
	 @emp_cid_end,
	 @emp_uf_end,
	 @emp_cep_end,
	 @emp_tel,
	 GETDATE())
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_EMPRESA_S
	(@emp_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   IF @emp_cd IS NULL
       SELECT * 
         FROM TB_EMPRESA        
        WHERE emp_dt_des IS NULL
   ELSE
      SELECT * 
        FROM TB_EMPRESA
       WHERE emp_cd = @emp_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_EMPRESA_U
	(@emp_cd 	int,
	 @emp_nm 	varchar(50),
	 @emp_cnpj 	char(14),
	 @emp_inscricao	char(12),
	 @emp_end 	varchar(50),
	 @emp_num_end 	int,
	 @emp_cmp_end 	varchar(20),
	 @emp_brr_end 	varchar(50),
	 @emp_cid_end 	varchar(50),
	 @emp_uf_end 	varchar(2),
	 @emp_cep_end 	char(8),
	 @emp_tel 	varchar(20),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
 UPDATE TB_EMPRESA 
    SET emp_nm = @emp_nm,
	emp_cnpj = @emp_cnpj,
	emp_inscricao = @emp_inscricao,
	emp_end = @emp_end,
	emp_num_end = @emp_num_end,
	emp_cmp_end = @emp_cmp_end,
	emp_brr_end = @emp_brr_end,
	emp_cid_end = @emp_cid_end,
	emp_uf_end = @emp_uf_end,
	emp_cep_end = @emp_cep_end,
	emp_dt_alt = GETDATE(),
	emp_tel = @emp_tel
  WHERE emp_cd	= @emp_cd
if @@ROWCOUNT = 0 
   BEGIN
      EXEC upTB_EMPRESA_I
         @emp_cd,
         @emp_nm,
	 @emp_cnpj,
	 @emp_inscricao,
	 @emp_end,
	 @emp_num_end,
	 @emp_cmp_end,
	 @emp_brr_end,
	 @emp_cid_end,
	 @emp_uf_end,
	 @emp_cep_end,
	 @emp_tel,
         @Erro,
         @MsgErr
   END
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_FERIADO_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_FERIADO_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_FERIADO_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_FERIADO_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_FERIADO_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_FERIADO_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_FERIADO_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_FERIADO_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_FERIADO_GRID') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_FERIADO_GRID
GO

--****************************************************
CREATE PROCEDURE upTB_FERIADO_D
	(@fer_data	 datetime,
	 @Erro           int OUTPUT,
         @MsgErr         varchar(255) OUTPUT)
AS
   DELETE TB_FERIADO 
    WHERE fer_data = @fer_data

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_FERIADO_GRID
	 ( @Erro           int OUTPUT,
           @MsgErr         varchar(255) OUTPUT)
AS
   SELECT fer_data 'Data', 
          fer_desc 'Feriado'
     FROM TB_FERIADO
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_FERIADO_I
	(@fer_data	datetime,
	 @fer_desc	varchar(20),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
INSERT INTO TB_FERIADO 
	 ( fer_data,
         fer_desc,
	 fer_dt_inc) 
VALUES 
	( @fer_data,	
         @fer_desc,
	 GETDATE())
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_FERIADO_S
	(@fer_data	datetime,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
    IF @fer_data IS NULL
       SELECT * 
         FROM TB_FERIADO
    ELSE
       SELECT * 
         FROM TB_FERIADO
        WHERE fer_data = @fer_data
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_FERIADO_U
	(@fer_data	datetime,
	 @fer_desc	varchar(20),
         @Erro        	int OUTPUT,
         @MsgErr       	varchar(255) OUTPUT)
AS 
 UPDATE TB_FERIADO 
    SET fer_desc = @fer_desc
  WHERE fer_data = @fer_data
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_FILME_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_FILME_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_FILME_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_FILME_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_FILME_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_FILME_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_COPIA_FILME_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_COPIA_FILME_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_COPIA_FILME_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_COPIA_FILME_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_COPIA_FILME_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_COPIA_FILME_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_FILME_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_FILME_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_FILME_GRID') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_FILME_GRID
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.vwTB_FILME_CARTAZ') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view dbo.vwTB_FILME_CARTAZ
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_FILME_CARTAZ') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_FILME_CARTAZ
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_FILME_CARTAZ2') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_FILME_CARTAZ2
GO

--****************************************************
CREATE VIEW vwTB_FILME_CARTAZ
AS

    SELECT DISTINCT c.fil_cd, c.fil_nm, c.fil_censura, 
           c.fil_duracao, b.prg_dt_ini, b.prg_dt_fim, 
           a.ses_dia_semana, c.fil_dt_ini, c.fil_dt_fim, 
           a.sal_cd
      FROM TB_SESSAO a,
           TB_PROGRAMACAO b,
           TB_FILME c
     WHERE a.prg_cd = b.prg_cd
       AND c.fil_cd = a.fil_cd
       AND a.ses_dt_des IS NULL
       AND b.prg_dt_des IS NULL
       AND c.fil_dt_des IS NULL

GO

--****************************************************
CREATE PROCEDURE upTB_COPIA_FILME_D
	(@fil_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   DELETE TB_COPIA_FILME 
   WHERE fil_cd = @fil_cd

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_COPIA_FILME_I
	(@fil_cd 	int,
	 @cfi_cd        int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
INSERT INTO TB_COPIA_FILME 
	 (fil_cd,
 	  cfi_cd) 
VALUES 
	(@fil_cd,
	 @cfi_cd)
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_COPIA_FILME_S
        (@fil_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
    SELECT * 
    FROM TB_COPIA_FILME
    WHERE fil_cd = @fil_cd
    ORDER BY cfi_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
/*
declare @erro int, @msgerr varchar(255), @data datetime
select @data = convert(datetime, '16/10/2005', 103 )
exec upTB_FILME_CARTAZ @data, @erro, @msgerr
*/
--*******************************************************
CREATE PROCEDURE upTB_FILME_CARTAZ
	(@DataExibicao	datetime,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
    declare @diaSemana    smallint
--    SET DATEFIRST 1
    ------------------------
    -- Pega dia da Semana --
    ------------------------
    
    select @diaSemana = 8
      From tb_feriado
     where fer_data = convert(datetime, @DataExibicao, 103)
    
    if @diaSemana is null
        select @diaSemana = dbo.ufDiaSemana(@DataExibicao)
        
    SELECT a.fil_cd, a.sal_cd, a.fil_censura, a.fil_duracao, 
	   b.sal_desc + ' - ' + a.fil_nm + ' - ' +
	   case fil_censura
	      when 0 then 'Livre'
              else CONVERT(VARCHAR, fil_censura) + ' Anos'
           end 'Sala - Filme - Classificação'
      FROM vwTB_FILME_CARTAZ a,
	   TB_SALA b
     WHERE a.sal_cd = b.sal_cd
       AND @DataExibicao between a.prg_dt_ini and a.prg_dt_fim
       AND a.ses_dia_semana = @diaSemana
       AND @DataExibicao between a.fil_dt_ini and a.fil_dt_fim
     ORDER BY b.sal_desc, a.fil_nm
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
/*
declare @erro int, @msgerr varchar(255), @data datetime
select @data = convert(datetime, '16/10/2005', 103 )
exec upTB_FILME_CARTAZ2 @data, @erro, @msgerr
*/
--*******************************************************
CREATE PROCEDURE upTB_FILME_CARTAZ2
	(@DataExibicao	datetime,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS

   DECLARE @HoraMaxSes     datetime,
           @DataIniPer     datetime,
           @DataFimPer     datetime

   SELECT @HoraMaxSes = par_hora_max_ses
   FROM tb_parametro
	  
   SELECT @DataIniPer = CONVERT(datetime, CONVERT(char(10),@DataExibicao,103) + ' ' + CONVERT(char(10),@HoraMaxSes,108), 103)
   SELECT @DataFimPer = CONVERT(datetime, CONVERT(char(10),DATEADD(Day, 1, @DataExibicao),103) + ' ' + CONVERT(char(10),@HoraMaxSes,108), 103)

    SELECT DISTINCT tb_sessao.sal_cd,
                    tb_sessao.fil_cd, 
                    tb_sala.sal_desc + ' - ' + tb_filme.fil_nm AS 'Sala - Filme',
                    'N' AS ses_excl
      FROM tb_sessao,
           tb_programacao,
           tb_filme,
           tb_sala
     WHERE tb_sessao.prg_cd = tb_programacao.prg_cd
       AND tb_sessao.fil_cd = tb_filme.fil_cd
       AND tb_sessao.sal_cd = tb_sala.sal_cd
       AND tb_sessao.ses_dt_des      IS NULL
       AND tb_programacao.prg_dt_des IS NULL
       AND tb_filme.fil_dt_des       IS NULL
       AND @DataExibicao between tb_programacao.prg_dt_ini and tb_programacao.prg_dt_fim
       AND tb_sessao.ses_dia_semana = dbo.ufDiaSemana(@DataExibicao)  
    UNION	
    SELECT DISTINCT tb_venda_ingresso.sal_cd,
                    tb_venda_ingresso.fil_cd, 
                    tb_sala.sal_desc + ' - ' + tb_filme.fil_nm AS 'Sala - Filme',
                    'S' AS ses_excl
    FROM tb_venda_ingresso,
         tb_filme,
         tb_sala
     WHERE tb_venda_ingresso.fil_cd   = tb_filme.fil_cd
     AND   tb_venda_ingresso.sal_cd   = tb_sala.sal_cd
     AND   tb_venda_ingresso.sre_data = @DataExibicao
     AND   tb_venda_ingresso.ing_dt_canc IS NULL
     --AND   tb_venda_ingresso.ing_dt_venda BETWEEN @DataIniPer AND @DataFimPer
       AND convert(varchar(10),tb_venda_ingresso.fil_cd) + 
           convert(varchar(3),tb_venda_ingresso.sal_cd) NOT IN
               (SELECT DISTINCT convert(varchar(10),tb_sessao.fil_cd) + 
                                convert(varchar(3),tb_sessao.sal_cd)
                FROM tb_sessao,
                     tb_programacao
                WHERE tb_sessao.prg_cd = tb_programacao.prg_cd
                AND tb_sessao.ses_dt_des      IS NULL
                AND tb_programacao.prg_dt_des IS NULL
                AND @DataExibicao between tb_programacao.prg_dt_ini and tb_programacao.prg_dt_fim
                AND tb_sessao.ses_dia_semana = dbo.ufDiaSemana(@DataExibicao))  
    UNION
    SELECT DISTINCT tb_venda_ingresso.sal_cd,
                    tb_venda_ingresso.fil_cd, 
                    tb_sala.sal_desc + ' - ' + tb_filme.fil_nm AS 'Sala - Filme',
                    'P' AS ses_excl
    FROM tb_venda_ingresso,
         tb_filme,
         tb_sala
     WHERE tb_venda_ingresso.fil_cd   = tb_filme.fil_cd
     AND   tb_venda_ingresso.sal_cd   = tb_sala.sal_cd
     AND   tb_venda_ingresso.sre_data <> @DataExibicao
     AND   tb_venda_ingresso.ing_dt_canc IS NULL
     AND   tb_venda_ingresso.ing_dt_venda BETWEEN @DataIniPer AND @DataFimPer
     AND convert(varchar(10),tb_venda_ingresso.fil_cd) + 
         convert(varchar(3),tb_venda_ingresso.sal_cd) NOT IN
               (SELECT DISTINCT convert(varchar(10),tb_sessao.fil_cd) + 
                                convert(varchar(3),tb_sessao.sal_cd)
                FROM tb_sessao,
                     tb_programacao
                WHERE tb_sessao.prg_cd = tb_programacao.prg_cd
                AND tb_sessao.ses_dt_des      IS NULL
                AND tb_programacao.prg_dt_des IS NULL
                AND @DataExibicao BETWEEN tb_programacao.prg_dt_ini and tb_programacao.prg_dt_fim
                AND tb_sessao.ses_dia_semana = dbo.ufDiaSemana(@DataExibicao))  
     AND convert(varchar(10),tb_venda_ingresso.fil_cd) + 
         convert(varchar(3),tb_venda_ingresso.sal_cd) IN
               (SELECT DISTINCT convert(varchar(10),tb_sessao.fil_cd) + 
                                convert(varchar(3),tb_sessao.sal_cd)
                FROM tb_sessao,
                     tb_programacao
                WHERE tb_sessao.prg_cd = tb_programacao.prg_cd
                AND tb_sessao.ses_dt_des      IS NULL
                AND tb_programacao.prg_dt_des IS NULL
                AND tb_venda_ingresso.sre_data BETWEEN tb_programacao.prg_dt_ini and tb_programacao.prg_dt_fim
                AND tb_sessao.ses_dia_semana = dbo.ufDiaSemana(tb_venda_ingresso.sre_data))  
    UNION
    SELECT DISTINCT tb_venda_ingresso.sal_cd,
                    tb_venda_ingresso.fil_cd, 
                    tb_sala.sal_desc + ' - ' + tb_filme.fil_nm AS 'Sala - Filme',
                    'Q' AS ses_excl
    FROM tb_venda_ingresso,
         tb_filme,
         tb_sala
     WHERE tb_venda_ingresso.fil_cd   = tb_filme.fil_cd
     AND   tb_venda_ingresso.sal_cd   = tb_sala.sal_cd
     AND   tb_venda_ingresso.sre_data <> @DataExibicao
     AND   tb_venda_ingresso.ing_dt_canc IS NULL
     AND   tb_venda_ingresso.ing_dt_venda BETWEEN @DataIniPer AND @DataFimPer
     AND convert(varchar(10),tb_venda_ingresso.fil_cd) + 
         convert(varchar(3),tb_venda_ingresso.sal_cd) NOT IN
               (SELECT DISTINCT convert(varchar(10),tb_sessao.fil_cd) + 
                                convert(varchar(3),tb_sessao.sal_cd)
                FROM tb_sessao,
                     tb_programacao
                WHERE tb_sessao.prg_cd = tb_programacao.prg_cd
                AND tb_sessao.ses_dt_des      IS NULL
                AND tb_programacao.prg_dt_des IS NULL
                AND @DataExibicao BETWEEN tb_programacao.prg_dt_ini and tb_programacao.prg_dt_fim
                AND tb_sessao.ses_dia_semana = dbo.ufDiaSemana(@DataExibicao))  
     AND convert(varchar(10),tb_venda_ingresso.fil_cd) + 
         convert(varchar(3),tb_venda_ingresso.sal_cd) NOT IN
               (SELECT DISTINCT convert(varchar(10),tb_sessao.fil_cd) + 
                                convert(varchar(3),tb_sessao.sal_cd)
                FROM tb_sessao,
                     tb_programacao
                WHERE tb_sessao.prg_cd = tb_programacao.prg_cd
                AND tb_sessao.ses_dt_des      IS NULL
                AND tb_programacao.prg_dt_des IS NULL
                AND tb_venda_ingresso.sre_data BETWEEN tb_programacao.prg_dt_ini and tb_programacao.prg_dt_fim
                AND tb_sessao.ses_dia_semana = dbo.ufDiaSemana(tb_venda_ingresso.sre_data))  
    
     AND convert(varchar(10),tb_venda_ingresso.fil_cd) + 
         convert(varchar(3),tb_venda_ingresso.sal_cd) NOT IN
               (SELECT DISTINCT convert(varchar(10),tb_sessao_real.fil_cd) + 
                                convert(varchar(3),tb_sessao_real.sal_cd)
                FROM tb_sessao_real,
                     tb_filme,
                     tb_sala
                WHERE tb_sessao_real.fil_cd   = tb_filme.fil_cd
                AND tb_sessao_real.sal_cd   = tb_sala.sal_cd
                AND tb_sessao_real.sre_data = @DataExibicao
                AND tb_sessao_real.sre_dt_des IS NULL
                AND tb_filme.fil_dt_des       IS NULL
                AND convert(varchar(10),tb_sessao_real.fil_cd) + 
                    convert(varchar(3),tb_sessao_real.sal_cd) NOT IN
                          (SELECT DISTINCT convert(varchar(10),tb_sessao.fil_cd) + 
                                           convert(varchar(3),tb_sessao.sal_cd)
                           FROM tb_sessao,
                                tb_programacao
                           WHERE tb_sessao.prg_cd = tb_programacao.prg_cd
                           AND tb_sessao.ses_dt_des      IS NULL
                           AND tb_programacao.prg_dt_des IS NULL
                           AND @DataExibicao between tb_programacao.prg_dt_ini and tb_programacao.prg_dt_fim
                           AND tb_sessao.ses_dia_semana = dbo.ufDiaSemana(@DataExibicao)))

   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_FILME_D
	(@fil_cd 	    int,
     @TipoExclusao  char(1),
     @Erro          int OUTPUT,
     @MsgErr        varchar(255) OUTPUT)
AS
	DECLARE @Qtdefil INT
	
    IF @TipoExclusao = 'L'
    	BEGIN
			UPDATE TB_FILME 
			SET fil_dt_des = GETDATE()
			WHERE fil_cd	= @fil_cd

			SELECT @Erro = @@ERROR

			IF @Erro <> 0
				  BEGIN
					 SELECT @MsgErr = description
					 FROM master..sysmessages
					 WHERE error = @Erro

					 RETURN
				  END
		END
    ELSE
    	BEGIN
    		SELECT  @Qtdefil = COUNT(fil_cd)
    		FROM tb_sessao
    		WHERE fil_cd = @fil_cd
    		AND   ses_dt_des IS NULL
    		
    		IF @Qtdefil > 0 
    			BEGIN
			         SELECT @Erro = 99
			         SELECT @MsgErr = 'Existe programação para este filme'
    			
    				 RETURN
    			END

    		SELECT  @Qtdefil = COUNT(fil_cd)
    		FROM tb_venda_ingresso
    		WHERE fil_cd = @fil_cd
    		
    		IF @Qtdefil > 0 
    			BEGIN
			         SELECT @Erro = 99
			         SELECT @MsgErr = 'Existe ingresso vendidos para este filme'
    			
    				 RETURN
    			END
    			
    		BEGIN TRANSACTION	

    		DELETE FROM tb_sessao
    		WHERE fil_cd = @fil_cd
    		AND   ses_dt_des IS NOT NULL

			SELECT @Erro = @@ERROR

			IF @Erro <> 0
				BEGIN
					 SELECT @MsgErr = description
					 FROM master..sysmessages
					 WHERE error = @Erro
					 
					 ROLLBACK TRANSACTION

					 RETURN
				END

    		DELETE FROM tb_copia_filme
    		WHERE fil_cd = @fil_cd

			SELECT @Erro = @@ERROR

			IF @Erro <> 0
				BEGIN
					 SELECT @MsgErr = description
					 FROM master..sysmessages
					 WHERE error = @Erro
					 
					 ROLLBACK TRANSACTION

					 RETURN
				END
    		
			DELETE tb_filme 
			WHERE fil_cd = @fil_cd

			SELECT @Erro = @@ERROR

			IF @Erro <> 0
				BEGIN
					 SELECT @MsgErr = description
					 FROM master..sysmessages
					 WHERE error = @Erro
					 
					 ROLLBACK TRANSACTION

					 RETURN
				END
				
			COMMIT TRANSACTION
		END

GO

--****************************************************
CREATE PROCEDURE upTB_FILME_GRID
	(@Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   SELECT TB_FILME.fil_cd, 
          TB_FILME.fil_durac_trai, 
          TB_DISTRIBUIDORA.dis_cd, 
          TB_FILME.fil_id_nacio,
          SUBSTRING(REPLICATE('0', 8 - LEN(LTRIM(STR(TB_FILME.fil_cd)))) + LTRIM(STR(TB_FILME.fil_cd)), 1, 4) 'Ano',
          SUBSTRING(REPLICATE('0', 8 - LEN(LTRIM(STR(TB_FILME.fil_cd)))) + LTRIM(STR(TB_FILME.fil_cd)), 5, 4) 'Código',
          TB_FILME.fil_nm 'Filme', 
          TB_DISTRIBUIDORA.dis_nm 'Distribuidora',
	  TB_FILME.fil_censura 'Censura', 
	  TB_FILME.fil_duracao 'Duração', 
	  TB_FILME.fil_dt_ini 'Início', 
	  TB_FILME.fil_dt_fim 'Término'
     FROM TB_FILME,
          TB_DISTRIBUIDORA
    WHERE TB_FILME.fil_dt_des IS NULL
    AND   TB_FILME.dis_cd = TB_DISTRIBUIDORA.dis_cd
   ORDER BY  fil_dt_ini
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_FILME_I
	(@fil_cd                int,
	 @fil_nm 		varchar(50),
	 @fil_censura 		smallint,
         @fil_duracao		smallint,
	 @fil_dt_ini		datetime,
	 @fil_dt_fim		datetime,
	 @fil_durac_trai        smallint,
	 @dis_cd                int,
	 @fil_id_nacio          varchar(1),
         @Erro         		int OUTPUT,
         @MsgErr       		varchar(255) OUTPUT)
AS 
INSERT INTO TB_FILME 
	 (fil_cd,
	  fil_nm,
	  fil_censura,
	  fil_duracao,
	  fil_dt_ini,
	  fil_dt_fim,
	  fil_durac_trai,
	  dis_cd,
	  fil_id_nacio,
 	  fil_dt_inc) 
VALUES 
	(@fil_cd,
	 @fil_nm,
	 @fil_censura,
	 @fil_duracao,
	 @fil_dt_ini,
	 @fil_dt_fim,
	 @fil_durac_trai,
	 @dis_cd,
	 @fil_id_nacio,
	 GETDATE())
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
      	 IF @Erro = 2627 
      	    BEGIN
               SELECT @MsgErr = 'Já existe filme com este código'
         
               RETURN
            END
         ELSE
            BEGIN
               SELECT @MsgErr = description
               FROM master..sysmessages
               WHERE error = @Erro

               RETURN
            END
      END

GO

--****************************************************
CREATE PROCEDURE upTB_FILME_S
	(@fil_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   IF @fil_cd IS NULL
       SELECT * 
         FROM TB_FILME
        WHERE fil_dt_des IS NULL
        ORDER BY fil_dt_ini
   ELSE
      SELECT * 
        FROM TB_FILME
       WHERE fil_cd = @fil_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_FILME_U
	(@fil_cd 		int,
         @fil_nm 		varchar(50),
	 @fil_censura 		smallint,
         @fil_duracao		smallint,
	 @fil_dt_ini		datetime,
	 @fil_dt_fim		datetime,
	 @fil_durac_trai        smallint,
	 @dis_cd                int,
	 @fil_id_nacio          varchar(1),
         @Erro          	int OUTPUT,
         @MsgErr        	varchar(255) OUTPUT)
AS 
 UPDATE TB_FILME 
    SET fil_nm         = @fil_nm,
	fil_censura    = @fil_censura,
        fil_duracao    = @fil_duracao,
	fil_dt_ini     = @fil_dt_ini,
	fil_dt_fim     = @fil_dt_fim,
	fil_durac_trai = @fil_durac_trai,
	dis_cd         = @dis_cd,
	fil_id_nacio   = @fil_id_nacio,
	fil_dt_alt     = GETDATE()
  WHERE fil_cd	= @fil_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_OPERACAO_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_OPERACAO_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_OPERACAO_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_OPERACAO_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_OPERACAO_I_T') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_OPERACAO_I_T
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_OPERACAO_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_OPERACAO_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_OPERACAO_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_OPERACAO_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_OPERACAO_CANCEL') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_OPERACAO_CANCEL
GO

--****************************************************
CREATE PROCEDURE upTB_OPERACAO_CANCEL
	(@ope_cd 	int,
	 @ope_mot_des	varchar(50),
         @Erro        	int OUTPUT,
         @MsgErr       	varchar(255) OUTPUT)
AS 
 UPDATE TB_OPERACAO 
    SET ope_dt_des = GETDATE(),
        ope_mot_des = @ope_mot_des
  WHERE ope_cd = @ope_cd

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_OPERACAO_D
	(@ope_cd 	int,
         @TipoExclusao  char(1),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
    IF @TipoExclusao = 'L'
	 UPDATE TB_OPERACAO 
	    SET ope_dt_des = GETDATE()
	  WHERE ope_cd	= @ope_cd
    ELSE
	DELETE TB_OPERACAO 
	 WHERE ope_cd = @ope_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_OPERACAO_I
	(@ope_cd 	int OUTPUT,
         @cxa_cd	int,
         @cxp_cd        int,
	 @opt_cd	int,
	 @ope_valor 	money,
	 @Erro         	int OUTPUT,
         @MsgErr       	varchar(255) OUTPUT)
AS 
    INSERT INTO TB_OPERACAO 
	 (cxa_cd,
	  cxp_cd,
	  opt_cd,
	  ope_valor,
	  ope_dt_operacao) 
    VALUES 
	(@cxa_cd,
	 @cxp_cd,
	 @opt_cd,
	 @ope_valor,
	 GETDATE())
   
   SELECT @ope_cd = @@IDENTITY
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_OPERACAO_I_T
	(@ope_cd 	int OUTPUT,
         @cxa_cd	int,
         @cxp_cd        int,
	 @opt_cd	int,
	 @ope_valor 	money,
	 @dt_operacao   datetime,
	 @Erro         	int OUTPUT,
         @MsgErr       	varchar(255) OUTPUT)
AS 
    DECLARE @dataOper datetime
    
    SELECT @dataOper = convert(datetime, convert(char(10),@dt_operacao,103) + ' ' + convert(char(10),GETDATE(),108), 103)
    INSERT INTO TB_OPERACAO 
	 (cxa_cd,
	  cxp_cd,
	  opt_cd,
	  ope_valor,
	  ope_dt_operacao) 
    VALUES 
	(@cxa_cd,
	 @cxp_cd,
	 @opt_cd,
	 @ope_valor,
	 @dataOper)
   
   SELECT @ope_cd = @@IDENTITY
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_OPERACAO_S
	(@ope_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   IF @ope_cd IS NULL
       SELECT * 
         FROM TB_OPERACAO
        WHERE ope_dt_des IS NULL
   ELSE
      SELECT * 
        FROM TB_OPERACAO
       WHERE ope_cd = @ope_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_OPERACAO_U
	(@ope_cd 	int,
         @cxa_cd	int,
         @cxp_cd        int,
	 @opt_cd	int,
	 @ope_valor 	money,
	 @Erro         	int OUTPUT,
         @MsgErr       	varchar(255) OUTPUT)
AS 
 UPDATE TB_OPERACAO 
    SET cxa_cd    = @cxa_cd,
        cxp_cd    = @cxp_cd,
	opt_cd    = @opt_cd,
	ope_valor = @ope_valor
  WHERE ope_cd	= @ope_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PAGAMENTO_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PAGAMENTO_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PAGAMENTO_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PAGAMENTO_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PAGAMENTO_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PAGAMENTO_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PAGAMENTO_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PAGAMENTO_S
GO

--****************************************************
CREATE PROCEDURE upTB_PAGAMENTO_D
	(@ope_cd 	int,
	 @pgt_cd	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   DELETE TB_PAGAMENTO 
    WHERE ope_cd = @ope_cd
      AND pgt_cd = @pgt_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_PAGAMENTO_I
	(@ope_cd 	int,
	 @pgt_cd	int,
	 @pag_valor	money,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
INSERT INTO TB_PAGAMENTO 
	 ( ope_cd,
	 pgt_cd,
	 pag_valor) 
VALUES 
	( @ope_cd,
	 @pgt_cd,
	 @pag_valor) 
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_PAGAMENTO_S
	(@ope_cd 	int,
	 @pgt_cd	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   IF @ope_cd IS NULL AND @pgt_cd IS NULL
       SELECT *
         FROM TB_PAGAMENTO
   ELSE
       IF @ope_cd IS NOT NULL AND @pgt_cd IS NULL
          SELECT * 
            FROM TB_PAGAMENTO
           WHERE ope_cd = @ope_cd
       ELSE
           IF @ope_cd IS NOT NULL AND @pgt_cd IS NOT NULL
              SELECT * 
                FROM TB_PAGAMENTO
               WHERE ope_cd = @ope_cd
                 AND pgt_cd = @pgt_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_PAGAMENTO_U
	(@ope_cd 	int,
	 @pgt_cd	int,
	 @pag_valor	money,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
 UPDATE TB_PAGAMENTO 
    SET pag_valor = @pag_valor
  WHERE ope_cd	= @ope_cd
    AND pgt_cd = @pgt_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PARAMETRO_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PARAMETRO_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PARAMETRO_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PARAMETRO_S
GO

--****************************************************
CREATE PROCEDURE upTB_PARAMETRO_S
         ( @Erro          int OUTPUT,
           @MsgErr        varchar(255) OUTPUT)
AS
   SELECT * 
     FROM TB_PARAMETRO        
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_PARAMETRO_U
	(@par_tmp_ses 		smallint,
	 @par_hora_max_ses 	datetime,
	 @par_hora_limite 	datetime,
         @par_hora_limite12     datetime,
         @par_hora_limite23     datetime,
         @par_hora_limite34     datetime,
         @par_hora_limite45     datetime,
         @par_hora_limite56     datetime,
	 @par_imp_cod_barra	bit,
	 @par_imp_lotacao	bit,
	 @par_imp_endereco	bit,
	 @par_imp_CNPJ		bit,
	 @par_imp_IE		bit,
	 @par_imp_tck		bit,
	 @par_custo_ingresso	money,
	 @par_imposto_mun	money,
	 @par_direitos_aut	money,
	 @par_outros		money,
	 @par_perc_meias        money,
	 @par_perc_cortesias    money,
	 @par_msg1              varchar(40),
	 @par_msg2              varchar(40),
	 @par_msg3              varchar(40),
	 @par_imp_MFIM		bit,
         @Erro          	int OUTPUT,
         @MsgErr        	varchar(255) OUTPUT)
AS 
 UPDATE TB_PARAMETRO 
    SET par_tmp_ses        = @par_tmp_ses,
	par_hora_max_ses   = @par_hora_max_ses,
	par_hora_limite    = @par_hora_limite,
	par_hora_limite12  = @par_hora_limite12,
	par_hora_limite23  = @par_hora_limite23,
	par_hora_limite34  = @par_hora_limite34,
	par_hora_limite45  = @par_hora_limite45,
	par_hora_limite56  = @par_hora_limite56,
	par_imp_cod_barra  = @par_imp_cod_barra,
	par_imp_lotacao    = @par_imp_lotacao,
	par_imp_endereco   = @par_imp_endereco,
	par_imp_CNPJ       = @par_imp_CNPJ,
	par_imp_IE         = @par_imp_IE,
	par_imp_tck        = @par_imp_tck,
	par_custo_ingresso = @par_custo_ingresso,
	par_imposto_mun    = @par_imposto_mun,
	par_direitos_aut   = @par_direitos_aut,
	par_outros         = @par_outros,
	par_perc_meias     = @par_perc_meias,
	par_perc_cortesias = @par_perc_cortesias,
	par_msg1           = @par_msg1,
	par_msg2           = @par_msg2,
	par_msg3           = @par_msg3,
	par_imp_MFIM       = @par_imp_MFIM
if @@ROWCOUNT = 0 
   BEGIN
      INSERT INTO TB_PARAMETRO 
	(par_tmp_ses, 
	 par_hora_max_ses, 
	 par_hora_limite, 
 	 par_hora_limite12,
 	 par_hora_limite23,
	 par_hora_limite34,
	 par_hora_limite45,
	 par_hora_limite56,
	 par_imp_cod_barra, 
	 par_imp_lotacao, 
	 par_imp_endereco, 
	 par_imp_CNPJ, 
	 par_imp_IE, 
	 par_imp_tck, 
	 par_custo_ingresso,
	 par_imposto_mun, 
	 par_direitos_aut, 
	 par_outros,
	 par_perc_meias,
	 par_perc_cortesias,
	 par_msg1,
	 par_msg2,
	 par_msg3,
	 par_imp_MFIM)
      VALUES 
	(@par_tmp_ses, 
	 @par_hora_max_ses, 
	 @par_hora_limite, 
 	 @par_hora_limite12,
 	 @par_hora_limite23,
	 @par_hora_limite34,
	 @par_hora_limite45,
	 @par_hora_limite56,
	 @par_imp_cod_barra, 
	 @par_imp_lotacao, 
	 @par_imp_endereco, 
	 @par_imp_CNPJ, 
	 @par_imp_IE, 
	 @par_imp_tck, 
	 @par_custo_ingresso,
	 @par_imposto_mun, 
	 @par_direitos_aut, 
	 @par_outros,
	 @par_perc_meias,
	 @par_perc_cortesias,
	 @par_msg1,
	 @par_msg2,
	 @par_msg3,
	 @par_imp_MFIM)
   END
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_POLTRONAS_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_POLTRONAS_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_POLTRONAS_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_POLTRONAS_S
GO

--****************************************************
CREATE PROCEDURE upTB_POLTRONAS_U
	(@sal_cd           int,
         @pol_tp_numeracao int,
         @pol_num_pri_col  varchar(4),
         @pol_num_filas    int,
         @pol_num_colunas  int,
         @pol_num_horiz    int,
         @pol_num_vert     int,
         @pol_poltronas    int,
         @pol_mat_poltr    varchar(2704),
         @Erro             int OUTPUT,
         @MsgErr           varchar(255) OUTPUT)
AS 
 UPDATE TB_POLTRONAS 
 SET pol_tp_numeracao = @pol_tp_numeracao,
     pol_num_pri_col  = @pol_num_pri_col,
     pol_num_filas    = @pol_num_filas,
     pol_num_colunas  = @pol_num_colunas,
     pol_num_horiz    = @pol_num_horiz,
     pol_num_vert     = @pol_num_vert,
     pol_poltronas    = @pol_poltronas,
     pol_mat_poltr    = @pol_mat_poltr
 WHERE sal_cd	= @sal_cd

IF @@ROWCOUNT = 0 
   BEGIN
      INSERT INTO TB_POLTRONAS
             (sal_cd,
              pol_tp_numeracao,
              pol_num_pri_col,
              pol_num_filas,
              pol_num_colunas,
              pol_num_horiz,
              pol_num_vert,
              pol_poltronas,
              pol_mat_poltr)
      VALUES (@sal_cd,
              @pol_tp_numeracao,
              @pol_num_pri_col,
              @pol_num_filas,
              @pol_num_colunas,
              @pol_num_horiz,
              @pol_num_vert,
              @pol_poltronas,
              @pol_mat_poltr)
   END
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_POLTRONAS_S
	(@sal_cd  int,
	 @Erro    int OUTPUT,
         @MsgErr  varchar(255) OUTPUT)
AS 

   SELECT sal_cd,
          pol_tp_numeracao,
          pol_num_pri_col,
          pol_num_filas,
          pol_num_colunas,
          pol_num_horiz,
          pol_num_vert,
          pol_poltronas,
          pol_mat_poltr 
   FROM TB_POLTRONAS        
   WHERE sal_cd = @sal_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PRECO_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PRECO_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PRECO_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PRECO_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PRECO_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PRECO_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PRECO_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PRECO_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PRECO_GRID') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PRECO_GRID
GO

--****************************************************
CREATE PROCEDURE upTB_PRECO_D
	(@ppr_cd	 int,
	 @fil_cd	 int,
         @pre_periodo	 smallint,
	 @TipoExclusao   char(1),
         @Erro           int OUTPUT,
         @MsgErr         varchar(255) OUTPUT)
AS
    IF @TipoExclusao = 'L'
	 UPDATE TB_PRECO 
	    SET pre_dt_des = GETDATE()
	  WHERE ppr_cd = @ppr_cd
	    AND fil_cd = @fil_cd
	    AND pre_periodo = @pre_periodo
    ELSE
	 DELETE TB_PRECO 
	  WHERE ppr_cd = @ppr_cd
	    AND fil_cd = @fil_cd
	    AND pre_periodo = @pre_periodo

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_PRECO_GRID
	(@ppr_cd	 int,
	 @Erro           int OUTPUT,
         @MsgErr         varchar(255) OUTPUT)
AS
   SELECT --convert(char(10),b.prg_dt_ini,103) + ' - ' + convert(char(10),b.prg_dt_fim,103) 'Programação',
          a.ppr_cd, 
          a.fil_cd,
	  c.fil_nm,
	  a.pre_periodo,
          a.pre_dia_semana,
          a.pre_vl_inteira_ate,
          a.pre_vl_inteira_apos,
          a.pre_vl_inteira3,
          a.pre_vl_inteira4,
          a.pre_vl_inteira5,
          a.pre_vl_inteira6,
          a.pre_vl_meia_ate,
          a.pre_vl_meia_apos,
          a.pre_vl_meia3,
          a.pre_vl_meia4,
          a.pre_vl_meia5,
          a.pre_vl_meia6,
          a.pre_Promocao
     FROM (TB_PRECO a INNER JOIN TB_PROG_PRECO b ON a.ppr_cd = b.ppr_cd)
          LEFT JOIN TB_FILME c ON a.fil_cd = c.fil_cd
    WHERE a.pre_dt_des IS NULL
      AND b.ppr_dt_des IS NULL
      AND a.ppr_cd = @ppr_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_PRECO_I
	(@ppr_cd	 	int,
	 @fil_cd	 	int,
         @pre_periodo	 	smallint,
         @pre_dia_semana 	smallint,
	 @pre_vl_inteira_ate	money,
	 @pre_vl_inteira_apos	money,
         @pre_vl_inteira3	money,
         @pre_vl_inteira4	money,
         @pre_vl_inteira5	money,
         @pre_vl_inteira6	money,
	 @pre_vl_meia_ate 	money,
	 @pre_vl_meia_apos	money,
         @pre_vl_meia3          money,
         @pre_vl_meia4          money,
         @pre_vl_meia5  	money,
         @pre_vl_meia6  	money,
         @pre_promocao		bit,
         @Erro           	int OUTPUT,
         @MsgErr         	varchar(255) OUTPUT)
AS 
INSERT INTO TB_PRECO 
	 ( ppr_cd,
	 fil_cd,
	 pre_periodo,
	 pre_dia_semana,
         pre_vl_inteira_ate,
	 pre_vl_inteira_apos,
         pre_vl_inteira3,
         pre_vl_inteira4,
         pre_vl_inteira5,
         pre_vl_inteira6,
	 pre_vl_meia_ate,
 	 pre_vl_meia_apos,
         pre_vl_meia3,
         pre_vl_meia4,
         pre_vl_meia5,
         pre_vl_meia6,
         pre_Promocao,
	 pre_dt_inc) 
VALUES 
	(@ppr_cd,	
	 @fil_cd,
	 @pre_periodo,
	 @pre_dia_semana,
         @pre_vl_inteira_ate,
	 @pre_vl_inteira_apos,
         @pre_vl_inteira3,
         @pre_vl_inteira4,
         @pre_vl_inteira5,
         @pre_vl_inteira6,
	 @pre_vl_meia_ate,
 	 @pre_vl_meia_apos,
         @pre_vl_meia3,
         @pre_vl_meia4,
         @pre_vl_meia5,
         @pre_vl_meia6,
         @pre_promocao,
	 GETDATE())
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_PRECO_S
	(@ppr_cd	 int,
	 @fil_cd	 int,
         @pre_periodo	 smallint,
         @Erro           int OUTPUT,
         @MsgErr         varchar(255) OUTPUT)
AS
    IF @ppr_cd IS NULL AND @fil_cd IS NULL 
       SELECT * 
         FROM TB_PRECO
        WHERE pre_dt_des IS NULL
    ELSE
        IF @ppr_cd IS NOT NULL AND @fil_cd IS NOT NULL AND @pre_periodo IS NOT NULL
            SELECT * 
              FROM TB_PRECO
             WHERE ppr_cd = @ppr_cd
               AND fil_cd = @fil_cd
               AND pre_periodo = @pre_periodo
        ELSE
            IF @ppr_cd IS NOT NULL AND @fil_cd IS NOT NULL AND @pre_periodo IS NULL
                SELECT *
                  FROM TB_PRECO
                 WHERE ppr_cd = @ppr_cd
                   AND fil_cd = @fil_cd
            ELSE
                IF @ppr_cd IS NOT NULL AND @fil_cd IS NULL 
                    SELECT * 
                      FROM TB_PRECO a
                     WHERE ppr_cd = @ppr_cd 
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_PRECO_U
	(@ppr_cd		int,
	 @fil_cd	 	int,
         @pre_periodo	 	smallint,
         @pre_dia_semana 	smallint,
	 @pre_vl_inteira_ate	money,
	 @pre_vl_inteira_apos	money,
         @pre_vl_inteira3	money,
         @pre_vl_inteira4	money,
         @pre_vl_inteira5	money,
         @pre_vl_inteira6	money,
	 @pre_vl_meia_ate 	money,
	 @pre_vl_meia_apos	money,
         @pre_vl_meia3          money,
         @pre_vl_meia4          money,
         @pre_vl_meia5  	money,
         @pre_vl_meia6  	money,
         @pre_promocao		bit,
         @Erro        		int OUTPUT,
         @MsgErr       		varchar(255) OUTPUT)
AS 
 UPDATE TB_PRECO 
    SET pre_vl_inteira_ate  = @pre_vl_inteira_ate,
        pre_vl_inteira_apos = @pre_vl_inteira_apos,
        pre_vl_inteira3     = @pre_vl_inteira3,
        pre_vl_inteira4     = @pre_vl_inteira4,
        pre_vl_inteira5     = @pre_vl_inteira5,
        pre_vl_inteira6     = @pre_vl_inteira6,
        pre_vl_meia_ate     = @pre_vl_meia_ate,
	pre_vl_meia_apos    = @pre_vl_meia_apos,
        pre_vl_meia3        = @pre_vl_meia3,
        pre_vl_meia4        = @pre_vl_meia4,
        pre_vl_meia5        = @pre_vl_meia5,
        pre_vl_meia6        = @pre_vl_meia6,
        pre_Promocao		= @pre_promocao,
        pre_dt_alt          = GETDATE()
  WHERE ppr_cd = @ppr_cd
    AND fil_cd = @fil_cd
    AND pre_periodo = @pre_periodo
    AND pre_dia_semana = @pre_dia_semana
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PROGRAMACAO_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PROGRAMACAO_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PROGRAMACAO_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PROGRAMACAO_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PROGRAMACAO_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PROGRAMACAO_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PROGRAMACAO_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PROGRAMACAO_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PROGRAMACAO_GRID') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PROGRAMACAO_GRID
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_VERIFICA_COPIAS') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_VERIFICA_COPIAS
GO

--****************************************************
CREATE PROCEDURE upTB_PROGRAMACAO_D
	(@prg_cd 	int,
         @TipoExclusao  char(1),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
    IF @TipoExclusao = 'L'
	 UPDATE TB_PROGRAMACAO 
	    SET prg_dt_des = GETDATE()
	  WHERE prg_cd	= @prg_cd
    ELSE
	DELETE TB_PROGRAMACAO 
	 WHERE prg_cd = @prg_cd

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_PROGRAMACAO_GRID
	(@Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   SELECT prg_cd, prg_dt_ini 'Início', prg_dt_fim 'Término'
     FROM TB_PROGRAMACAO
    WHERE prg_dt_des IS NULL
   ORDER BY prg_dt_fim 
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_PROGRAMACAO_I
	(@prg_dt_ini	datetime,
	 @prg_dt_fim	datetime,
	 @copiaSess     int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 

declare @qtdeProg int,
        @prgCd1   int,
        @prgCd2   int

SELECT @qtdeProg = COUNT(*)
FROM TB_PROGRAMACAO
WHERE @prg_dt_ini between prg_dt_ini AND prg_dt_fim
OR    @prg_dt_fim between prg_dt_ini AND prg_dt_fim
OR    (@prg_dt_ini <= prg_dt_ini AND @prg_dt_fim >= prg_dt_fim)
OR    @prg_dt_ini = prg_dt_ini
OR    @prg_dt_fim = prg_dt_fim

IF @qtdeProg <> 0
   BEGIN
      SELECT @Erro = 99
      SELECT @MsgErr = 'Ocorreu sobreposição de período de programação'
      
      RETURN
   END
   
IF @copiaSess = 1 
   SELECT @prgCd2 = prg_cd
   FROM TB_PROGRAMACAO
   WHERE prg_dt_ini = (SELECT MAX(prg_dt_ini) 
                       FROM TB_PROGRAMACAO
                       WHERE prg_dt_des IS NULL
                       AND   prg_dt_ini <= @prg_dt_ini)

INSERT INTO TB_PROGRAMACAO 
	 (prg_dt_ini,
	  prg_dt_fim,
 	  prg_dt_inc) 
VALUES 
	(@prg_dt_ini,
	 @prg_dt_fim,
	 GETDATE())
	 
SELECT @prgCd1 = @@IDENTITY	 
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END
      
IF @copiaSess = 1 
   INSERT INTO TB_SESSAO (prg_cd,
                          sal_cd,
                          fil_cd,
                          ses_cd,
                          ses_periodo,
                          ses_horario,
                          ses_dia_semana,
                          ses_dt_inc,
                          ses_dt_alt,
                          ses_dt_des,
                          ses_mot_des,
                          ses_pre_estreia)
   SELECT @prgCd1,
          tb_sessao.sal_cd,
          tb_sessao.fil_cd,
          tb_sessao.ses_cd,
          tb_sessao.ses_periodo,
          tb_sessao.ses_horario,
          tb_sessao.ses_dia_semana,
          tb_sessao.ses_dt_inc,
          tb_sessao.ses_dt_alt,
          tb_sessao.ses_dt_des,
          tb_sessao.ses_mot_des,
          tb_sessao.ses_pre_estreia
   FROM tb_sessao,
        tb_filme
   WHERE tb_sessao.prg_cd = @prgCd2
   AND   tb_sessao.fil_cd = tb_filme.fil_cd
   AND   tb_sessao.ses_dt_des IS NULL
   AND   (tb_filme.fil_dt_fim between @prg_dt_ini         and @prg_dt_fim
   OR     @prg_dt_ini         between tb_filme.fil_dt_ini and tb_filme.fil_dt_fim)

GO

--****************************************************
CREATE PROCEDURE upTB_PROGRAMACAO_S
	(@prg_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   IF @prg_cd IS NULL
       SELECT * 
         FROM TB_PROGRAMACAO
        WHERE prg_dt_des IS NULL
   ELSE
      SELECT * 
        FROM TB_PROGRAMACAO
       WHERE prg_cd = @prg_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_PROGRAMACAO_U
	(@prg_cd 	int,
         @prg_dt_ini	datetime,
	 @prg_dt_fim	datetime,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
declare @qtdeProg int
SELECT @qtdeProg = COUNT(*)
FROM TB_PROGRAMACAO
WHERE (@prg_dt_ini between prg_dt_ini AND prg_dt_fim
       OR @prg_dt_fim between prg_dt_ini AND prg_dt_fim
       OR (@prg_dt_ini <= prg_dt_ini AND @prg_dt_fim >= prg_dt_fim)
       OR @prg_dt_ini = prg_dt_ini
       OR @prg_dt_fim = prg_dt_fim)
AND prg_cd <> @prg_cd       
IF @qtdeProg <> 0
   BEGIN
      SELECT @Erro = 99
      SELECT @MsgErr = 'Ocorreu sobreposição de período de programação'
      
      RETURN
   END
 UPDATE TB_PROGRAMACAO 
    SET prg_dt_ini = @prg_dt_ini,
	prg_dt_fim = @prg_dt_fim,
	prg_dt_alt = GETDATE()
  WHERE prg_cd	= @prg_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_VERIFICA_COPIAS
	(@prg_dt_ini	datetime,
	 @prg_dt_fim	datetime,
	 @QtdeFil       int OUTPUT,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 

declare @prgCd2  int

SELECT @prgCd2 = prg_cd
FROM TB_PROGRAMACAO
WHERE prg_dt_ini = (SELECT MAX(prg_dt_ini) 
                    FROM TB_PROGRAMACAO
                    WHERE prg_dt_des IS NULL
                    AND   prg_dt_ini <= @prg_dt_ini)
   
SELECT @QtdeFil = COUNT(*)
FROM tb_sessao,
     tb_filme
WHERE tb_sessao.prg_cd = @prgCd2
AND   tb_sessao.fil_cd = tb_filme.fil_cd
AND   tb_sessao.ses_dt_des IS NULL
AND   (tb_filme.fil_dt_fim between @prg_dt_ini         and @prg_dt_fim
OR     @prg_dt_ini         between tb_filme.fil_dt_ini and tb_filme.fil_dt_fim)

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PROG_COMBO_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PROG_COMBO_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PROG_COMBO_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PROG_COMBO_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PROG_COMBO_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PROG_COMBO_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PROG_COMBO_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PROG_COMBO_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PROG_COMBO_GRID') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PROG_COMBO_GRID
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PROG_COMBO_DATA') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PROG_COMBO_DATA
GO

--****************************************************
CREATE PROCEDURE upTB_PROG_COMBO_D
	(@cbo_cd 	int,
	 @pcb_dt_ini	datetime,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   DELETE TB_PROG_COMBO 
    WHERE cbo_cd = @cbo_cd
      AND pcb_dt_ini = @pcb_dt_ini

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_PROG_COMBO_DATA
	(@DataExibicao	datetime,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
--SET DATEFIRST 1
    select a.cbo_cd, a.cbo_nm, a.cbo_desc, b.pcb_valor 
      from tb_combo a,
           tb_prog_combo b
     where a.cbo_cd = b.cbo_cd
       and a.cbo_dt_des is null
       and convert(datetime, @DataExibicao, 103) between b.pcb_dt_ini and b.pcb_dt_fim
     order by a.cbo_nm
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_PROG_COMBO_GRID
	(@Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   SELECT a.cbo_cd, b.cbo_nm 'Combo', a.pcb_dt_ini 'Início', a.pcb_dt_fim 'Término', a.pcb_valor 'Valor'
     FROM TB_PROG_COMBO a,
          TB_COMBO b
    WHERE a.cbo_cd = b.cbo_cd
      AND b.cbo_dt_des IS NULL
   ORDER BY a.pcb_dt_ini
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_PROG_COMBO_I
	(@cbo_cd 	int,
	 @pcb_dt_ini	datetime,
	 @pcb_dt_fim	datetime,
	 @pcb_valor	money,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
INSERT INTO TB_PROG_COMBO 
	 ( cbo_cd,
	 pcb_dt_ini,
	 pcb_dt_fim,
 	 pcb_valor) 
VALUES 
	( @cbo_cd,
	 @pcb_dt_ini,
	 @pcb_dt_fim,
 	 @pcb_valor) 
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_PROG_COMBO_S
	(@cbo_cd 	int,
	 @pcb_dt_ini	datetime,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   IF @cbo_cd IS NULL AND @pcb_dt_ini IS NULL
       SELECT *
         FROM TB_PROG_COMBO
   ELSE
       IF @cbo_cd IS NOT NULL AND @pcb_dt_ini IS NULL
          SELECT * 
            FROM TB_PROG_COMBO
           WHERE cbo_cd = @cbo_cd
       ELSE
           IF @cbo_cd IS NOT NULL AND @pcb_dt_ini IS NOT NULL
              SELECT * 
                FROM TB_PROG_COMBO
               WHERE cbo_cd = @cbo_cd
                 AND pcb_dt_ini = @pcb_dt_ini
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_PROG_COMBO_U
	(@cbo_cd 	int,
	 @pcb_dt_ini	datetime,
	 @pcb_dt_fim	datetime,
	 @pcb_valor	money,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
 UPDATE TB_PROG_COMBO 
    SET pcb_dt_fim = @pcb_dt_fim,
	pcb_valor = @pcb_valor
  WHERE cbo_cd	= @cbo_cd
    AND pcb_dt_ini = @pcb_dt_ini
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PROG_PRECO_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PROG_PRECO_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PROG_PRECO_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PROG_PRECO_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PROG_PRECO_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PROG_PRECO_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PROG_PRECO_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PROG_PRECO_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_PROG_PRECO_GRID') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_PROG_PRECO_GRID
GO

--****************************************************
CREATE PROCEDURE upTB_PROG_PRECO_D
	(@ppr_cd 	int,
         @TipoExclusao  char(1),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
    IF @TipoExclusao = 'L'
	 UPDATE TB_PROG_PRECO 
	    SET ppr_dt_des = GETDATE()
	  WHERE ppr_cd	= @ppr_cd
    ELSE
	DELETE TB_PROG_PRECO 
	 WHERE ppr_cd = @ppr_cd
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_PROG_PRECO_GRID
	(@Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   SELECT ppr_cd, ppr_dt_ini 'Início', 
	  ppr_dt_fim 'Término',
	  ppr_flg_promocao 'Promoção', 
	  ppr_desc 'Descrição', 
	  ppr_patrocinador 'Patrocinador'
     FROM TB_PROG_PRECO
    WHERE ppr_dt_des IS NULL
   ORDER BY ppr_dt_ini 
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_PROG_PRECO_I
	(@ppr_dt_ini		datetime,
	 @ppr_dt_fim		datetime,
	 @ppr_flg_promocao 	bit,
	 @ppr_desc		varchar(50),
	 @ppr_patrocinador 	varchar(50),
         @Erro         		int OUTPUT,
         @MsgErr       		varchar(255) OUTPUT)
AS 
INSERT INTO TB_PROG_PRECO 
	 ( ppr_dt_ini,
	 ppr_dt_fim,
	 ppr_flg_promocao,
	 ppr_desc,
	 ppr_patrocinador,
 	 ppr_dt_inc) 
VALUES 
	( @ppr_dt_ini,
	 @ppr_dt_fim,
	 @ppr_flg_promocao,
	 @ppr_desc,
	 @ppr_patrocinador,
	 GETDATE())
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_PROG_PRECO_S
	(@ppr_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   IF @ppr_cd IS NULL
       SELECT * 
         FROM TB_PROG_PRECO
        WHERE ppr_dt_des IS NULL
   ELSE
      SELECT * 
        FROM TB_PROG_PRECO
       WHERE ppr_cd = @ppr_cd
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_PROG_PRECO_U
	(@ppr_cd 		int,
         @ppr_dt_ini		datetime,
	 @ppr_dt_fim		datetime,
	 @ppr_flg_promocao 	bit,
	 @ppr_desc		varchar(50),
	 @ppr_patrocinador 	varchar(50),
         @Erro         		int OUTPUT,
         @MsgErr       		varchar(255) OUTPUT)
AS 
 UPDATE TB_PROG_PRECO 
    SET ppr_dt_ini = @ppr_dt_ini,
	ppr_dt_fim = @ppr_dt_fim,
	ppr_flg_promocao = @ppr_flg_promocao,
	ppr_desc = @ppr_desc,
	ppr_patrocinador = @ppr_patrocinador,
	ppr_dt_alt = GETDATE()
  WHERE ppr_cd	= @ppr_cd
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SALA_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SALA_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SALA_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SALA_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SALA_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SALA_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SALA_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SALA_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SALA_GRID') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SALA_GRID
GO

--****************************************************
CREATE PROCEDURE upTB_SALA_D
	(@sal_cd 	int,
         @TipoExclusao  char(1),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
    IF @TipoExclusao = 'L'
	 UPDATE TB_SALA 
	    SET sal_dt_des = GETDATE()
	  WHERE sal_cd	= @sal_cd
    ELSE
	DELETE TB_SALA 
	 WHERE sal_cd = @sal_cd

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_SALA_GRID
	(@Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   SELECT a.sal_cd, b.cin_cd, b.cin_nm 'Cinema', a.sal_desc 'Sala', a.sal_lugares 'Lugares'
     FROM TB_SALA a,
          TB_CINEMA b,
          TB_EMPRESA c
    WHERE b.emp_cd = c.emp_cd
      AND a.cin_cd = b.cin_cd
      AND a.sal_dt_des IS NULL
      AND b.cin_dt_des IS NULL
      AND c.emp_dt_des IS NULL
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_SALA_I
	(@cin_cd 	int,
         @sal_desc 	varchar(50),
	 @sal_lugares 	smallint,	 
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
INSERT INTO TB_SALA 
	 (cin_cd,
          sal_desc,
	  sal_lugares,
 	  sal_dt_inc) 
VALUES 
	(@cin_cd,
         @sal_desc,
	 @sal_lugares,
	 GETDATE())
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_SALA_S
	(@sal_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   IF @sal_cd IS NULL
       SELECT a.*
         FROM TB_SALA a,
              TB_CINEMA b,
              TB_EMPRESA c
        WHERE b.emp_cd = c.emp_cd
          AND a.cin_cd = b.cin_cd
          AND a.sal_dt_des IS NULL
          AND b.cin_dt_des IS NULL
          AND c.emp_dt_des IS NULL
   ELSE
      SELECT * 
        FROM TB_SALA
       WHERE sal_cd = @sal_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_SALA_U
	(@sal_cd 	int,
	 @cin_cd 	int,
         @sal_desc 	varchar(50),
	 @sal_lugares 	smallint,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
 UPDATE TB_SALA 
    SET cin_cd       = @cin_cd,
        sal_desc     = @sal_desc,
	sal_lugares  = @sal_lugares,
	sal_dt_alt   = GETDATE()
  WHERE sal_cd	= @sal_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SALA_LUGAR_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SALA_LUGAR_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SALA_LUGAR_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SALA_LUGAR_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SALA_LUGAR_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SALA_LUGAR_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SALA_LUGAR_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SALA_LUGAR_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SALA_LUGAR_GRID') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SALA_LUGAR_GRID
GO

--****************************************************
CREATE PROCEDURE upTB_SALA_LUGAR_D
	(@sal_cd 	int,
	 @sal_dt_ini	datetime,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   DELETE TB_SALA_LUGAR 
    WHERE sal_cd = @sal_cd
      AND sal_dt_ini = @sal_dt_ini

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_SALA_LUGAR_GRID
	(@Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   SELECT a.sal_cd, b.sal_desc 'Sala', a.sal_dt_ini 'Início', a.sal_dt_fim 'Término', 
          a.sal_lugares 'Lugares', a.sal_mot_alt 'Motivo'
     FROM TB_SALA_LUGAR a,
          TB_SALA b
    WHERE a.sal_cd = b.sal_cd
      AND b.sal_dt_des IS NULL
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_SALA_LUGAR_I
	(@sal_cd 	int,
	 @sal_dt_ini	datetime,
	 @sal_dt_fim	datetime,
	 @sal_lugares	smallint,
	 @sal_mot_alt	varchar(50),
	 @usu_cd	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
INSERT INTO TB_SALA_LUGAR 
	 ( sal_cd,
	 sal_dt_ini,
	 sal_dt_fim,
 	 sal_lugares,
	 sal_mot_alt,
	 usu_cd) 
VALUES 
	( @sal_cd,
	 @sal_dt_ini,
	 @sal_dt_fim,
 	 @sal_lugares,
	 @sal_mot_alt,
	 @usu_cd)  
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_SALA_LUGAR_S
	(@sal_cd 	int,
	 @sal_dt_ini	datetime,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   IF @sal_cd IS NULL AND @sal_dt_ini IS NULL
       SELECT *
         FROM TB_SALA_LUGAR
   ELSE
       IF @sal_cd IS NOT NULL AND @sal_dt_ini IS NULL
          SELECT * 
            FROM TB_SALA_LUGAR
           WHERE sal_cd = @sal_cd
       ELSE
           IF @sal_cd IS NOT NULL AND @sal_dt_ini IS NOT NULL
              SELECT * 
                FROM TB_SALA_LUGAR
               WHERE sal_cd = @sal_cd
                 AND sal_dt_ini = @sal_dt_ini
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_SALA_LUGAR_U
	(@sal_cd 	int,
	 @sal_dt_ini	datetime,
	 @sal_dt_fim	datetime,
	 @sal_lugares	smallint,
	 @sal_mot_alt	varchar(50),
	 @usu_cd	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
 UPDATE TB_SALA_LUGAR 
    SET sal_dt_fim = @sal_dt_fim,
	sal_lugares = @sal_lugares,
	sal_mot_alt = @sal_mot_alt,
	usu_cd = @usu_cd
  WHERE sal_cd	= @sal_cd
    AND sal_dt_ini = @sal_dt_ini
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SESSAO_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SESSAO_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SESSAO_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SESSAO_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SESSAO_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SESSAO_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SESSAO_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SESSAO_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SESSAO_GRID') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SESSAO_GRID
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SESSAO_COLISAO') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SESSAO_COLISAO
GO

--****************************************************
CREATE PROCEDURE upTB_SESSAO_COLISAO
	(@prg_cd	 int,
	 @sal_cd	 int,
	 @fil_cd	 int,
	 @ses_dia_semana smallint,
	 @ses_horario	 datetime,
	 @Erro           int OUTPUT,
         @MsgErr         varchar(255) OUTPUT)
AS
SET NOCOUNT ON
DECLARE @par_tmp_ses     smallint,
        @fil_duracao     smallint,
        @ses_horario_out datetime,
        @fil_nm_out      varchar(50)
SELECT @ses_horario = dateadd(s, 1, @ses_horario )
SELECT @par_tmp_ses = par_tmp_ses 
  FROM tb_parametro
SELECT @fil_duracao = fil_duracao
  FROM tb_filme
 WHERE fil_cd = @fil_cd
SELECT @fil_nm_out = b.fil_nm, 
       @ses_horario_out = a.ses_horario
  FROM tb_sessao a, tb_filme b
 WHERE a.fil_cd = b.fil_cd
   AND a.prg_CD = @prg_cd
   AND a.sal_cd = @sal_cd
   AND a.fil_cd <> @fil_cd
   AND a.ses_dia_semana = @ses_dia_semana
   AND (    ( a.ses_horario <= @ses_horario AND DateAdd(n, b.fil_duracao + @par_tmp_ses, a.ses_horario) >= @ses_horario )
         OR ( a.ses_horario >= @ses_horario AND DateAdd(n, @fil_duracao + @par_tmp_ses, @ses_horario) >= a.ses_horario )
       )

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END
   IF @fil_nm_out IS NOT NULL
       SELECT @Erro = 1
       SELECT @MsgErr = 'Colisão com Filme ' + char(34) + @fil_nm_out + char(34) + ' no horário das ' + convert(char(5),@ses_horario_out, 108)
       RETURN

GO

--****************************************************
CREATE PROCEDURE upTB_SESSAO_D
	(@prg_cd	 int,
	 @sal_cd	 int,
	 @fil_cd	 int,
	 @ses_periodo	 smallint,
	 @TipoExclusao   char(1),
         @Erro           int OUTPUT,
         @MsgErr         varchar(255) OUTPUT)
AS
    IF @TipoExclusao = 'L'
	 UPDATE TB_SESSAO 
	    SET ses_dt_des = GETDATE()
	  WHERE prg_cd = @prg_cd
	    AND sal_cd = @sal_cd
	    AND fil_cd = @fil_cd
	    AND ses_periodo = @ses_periodo
    ELSE
	 DELETE TB_SESSAO 
	  WHERE prg_cd = @prg_cd
	    AND sal_cd = @sal_cd
	    AND fil_cd = @fil_cd
	    AND ses_periodo = @ses_periodo
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_SESSAO_GRID
	(@prg_cd	 int,
	 @Erro           int OUTPUT,
         @MsgErr         varchar(255) OUTPUT)
AS
   SELECT --convert(char(10),b.prg_dt_ini,103) + ' - ' + convert(char(10),b.prg_dt_fim,103) 'Programação',
          a.prg_cd, 
          a.sal_cd,
          a.fil_cd,
	  a.ses_periodo,
          c.sal_desc,
          d.fil_nm,
          a.ses_cd,
          a.ses_dia_semana,
          a.ses_horario,
          a.ses_pre_estreia
     FROM TB_SESSAO a,
          TB_PROGRAMACAO b,
          TB_SALA c,
          TB_FILME d
    WHERE a.prg_cd = b.prg_cd
      AND a.sal_cd = c.sal_cd
      AND a.fil_cd = d.fil_cd
      AND a.ses_dt_des IS NULL
      AND b.prg_dt_des IS NULL
      AND c.sal_dt_des IS NULL
      AND d.fil_dt_des IS NULL
      AND a.prg_cd = @prg_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_SESSAO_I
	(@prg_cd	  int,
	 @sal_cd	  int,
	 @fil_cd	  int,
	 @ses_periodo	  smallint,
	 @ses_cd	  smallint,
	 @ses_dia_semana  smallint,
	 @ses_horario	  datetime,
	 @ses_pre_estreia varchar(1),
         @Erro            int OUTPUT,
         @MsgErr          varchar(255) OUTPUT)
AS 
INSERT INTO TB_SESSAO 
	 ( prg_cd,
	 sal_cd,
	 fil_cd,
         ses_periodo,
	 ses_cd,
	 ses_dia_semana,
 	 ses_horario,
 	 ses_pre_estreia,
	 ses_dt_inc) 
VALUES 
	(@prg_cd,
	 @sal_cd,
	 @fil_cd,
         @ses_periodo,
	 @ses_cd,
	 @ses_dia_semana,
 	 @ses_horario,
 	 @ses_pre_estreia,
	 GETDATE())
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_SESSAO_S
	(@prg_cd	  int,
	 @sal_cd	  int,
	 @fil_cd	  int,
	 @ses_periodo     smallint,
	 @Erro            int OUTPUT,
         @MsgErr          varchar(255) OUTPUT)
AS
    IF @prg_cd IS NULL AND @sal_cd IS NULL AND @fil_cd IS NULL AND @ses_periodo IS NULL
       SELECT * 
         FROM TB_SESSAO
        WHERE ses_dt_des IS NULL
    ELSE
        IF @prg_cd IS NOT NULL AND @sal_cd IS NOT NULL AND @fil_cd IS NOT NULL AND @ses_periodo IS NOT NULL
            SELECT * 
              FROM TB_SESSAO
             WHERE prg_cd = @prg_cd
               AND sal_cd = @sal_cd
               AND fil_cd = @fil_cd
               AND ses_periodo = @ses_periodo
        ELSE
            IF @prg_cd IS NOT NULL AND @sal_cd IS NOT NULL AND @fil_cd IS NOT NULL AND @ses_periodo IS NULL
                SELECT * 
                  FROM TB_SESSAO
                 WHERE prg_cd = @prg_cd
                   AND sal_cd = @sal_cd
                   AND fil_cd = @fil_cd
            ELSE
                IF @prg_cd IS NOT NULL AND @sal_cd IS NOT NULL AND @fil_cd IS NULL AND @ses_periodo IS NULL
                    SELECT * 
                      FROM TB_SESSAO
                     WHERE prg_cd = @prg_cd
                       AND sal_cd = @sal_cd
                ELSE
                    IF @prg_cd IS NOT NULL AND @sal_cd IS NULL AND @fil_cd IS NULL AND @ses_periodo IS NULL
                        SELECT * 
                          FROM TB_SESSAO
                         WHERE prg_cd = @prg_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_SESSAO_U
	(@prg_cd	  int,
	 @sal_cd	  int,
	 @fil_cd	  int,
	 @ses_periodo	  smallint,
	 @ses_cd	  smallint,
	 @ses_dia_semana  smallint,
	 @ses_horario	  datetime,
	 @ses_pre_estreia varchar(1),
         @Erro            int OUTPUT,
         @MsgErr          varchar(255) OUTPUT)
AS 
 UPDATE TB_SESSAO 
    SET ses_horario     = @ses_horario,
        ses_pre_estreia = @ses_pre_estreia,
	ses_dt_alt      = GETDATE()
  WHERE prg_cd         = @prg_cd
    AND sal_cd         = @sal_cd
    AND fil_cd         = @fil_cd
    AND ses_periodo    = @ses_periodo
    AND ses_cd         = @ses_cd
    AND ses_dia_semana = @ses_dia_semana
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SESSAO_AUX_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SESSAO_AUX_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SESSAO_AUX_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SESSAO_AUX_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SESSAO_AUX_Lot') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SESSAO_AUX_Lot
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SESSAO_AUX_U_P') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SESSAO_AUX_U_P
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SESSAO_AUX_P_P') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SESSAO_AUX_P_P
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SESSAO_AUX_L_P') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SESSAO_AUX_L_P
GO

--****************************************************
CREATE PROCEDURE upTB_SESSAO_AUX_D
   (@cxa_cd      int,
    @Erro        int OUTPUT,
    @MsgErr      varchar(255) OUTPUT)
AS
   DELETE TB_SESSAO_AUX 
    WHERE cxa_cd      = @cxa_cd 

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_SESSAO_AUX_Lot
   (@sal_cd          int,
    @fil_cd          int,
    @sre_data        datetime,
    @sre_horario     datetime,
    @sre_lugares     int OUTPUT,
    @sea_lugares_sel int OUTPUT,
    @sea_lugares_ven int OUTPUT,
    @sea_inteiras    int OUTPUT,
    @sre_inteiras    int OUTPUT,
    @sea_meias       int OUTPUT,
    @sre_meias       int OUTPUT,
    @sea_cortesias   int OUTPUT,
    @sre_cortesias   int OUTPUT,
    @Erro            int OUTPUT,
    @MsgErr          varchar(255) OUTPUT)
AS
   DECLARE @aux1  int,
           @aux2  int,
           @aux3  int
   
   SELECT @aux1          = TB_SESSAO_REAL.sre_lugares,
          @aux2          = TB_SESSAO_REAL.sre_lugares_vendidos,
          @sre_inteiras  = TB_SESSAO_REAL.sre_inteiras,
          @sre_meias     = TB_SESSAO_REAL.sre_meias,
          @sre_cortesias = TB_SESSAO_REAL.sre_cortesias
    FROM TB_SESSAO_REAL
    WHERE TB_SESSAO_REAL.sal_cd      = @sal_cd
      AND TB_SESSAO_REAL.fil_cd      = @fil_cd
      AND TB_SESSAO_REAL.sre_data    = @sre_data
      AND TB_SESSAO_REAL.sre_horario = @sre_horario
      
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END
      
   SELECT @sre_lugares = @aux1

   SELECT @aux3 = sal_lugares
   FROM tb_sala_lugar
   WHERE @sre_data between sal_dt_ini and sal_dt_fim
   AND   sal_cd = @sal_cd
   
   IF @aux3 IS NULL
      SELECT @aux3 = TB_SALA.sal_lugares
      FROM TB_SALA
      WHERE TB_SALA.sal_cd = @sal_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

   IF @aux3 IS NULL 
      SELECT @aux1 = 0
   ELSE
      SELECT @aux1 = @aux3
         
   IF @aux2 IS NULL 
      SELECT @aux2 = 0
      
   SELECT @sea_lugares_ven = @aux1 - @aux2
   
   SELECT @sea_lugares_sel = SUM(sea_lugares_sel),
          @sea_inteiras    = SUM(sea_inteiras),
          @sea_meias       = SUM(sea_meias),
          @sea_cortesias   = SUM(sea_cortesias)
    FROM TB_SESSAO_AUX
    WHERE sal_cd      = @sal_cd
      AND fil_cd      = @fil_cd
      AND sre_data    = @sre_data
      AND sre_horario = @sre_horario
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END
      
   IF @sea_lugares_sel IS NULL 
      SELECT @sea_lugares_sel = 0

   IF @sre_lugares IS NULL
      SELECT @sre_lugares = @aux3

   IF @sea_inteiras IS NULL
      SELECT @sea_inteiras =0

   IF @sea_meias IS NULL
      SELECT @sea_meias = 0

   IF @sea_cortesias IS NULL
      SELECT @sea_cortesias = 0

   IF @sre_inteiras IS NULL
      SELECT @sre_inteiras = 0

   IF @sre_meias IS NULL
      SELECT @sre_meias = 0

   IF @sre_cortesias IS NULL
      SELECT @sre_cortesias = 0
      
GO

--****************************************************
CREATE PROCEDURE upTB_SESSAO_AUX_U
   (@sal_cd          int,
    @fil_cd          int,
    @sre_data        datetime,
    @sre_horario     datetime,
    @cxa_cd          int,
    @sea_lugares_sel int,
    @sea_inteiras    int, 
    @sea_meias       int, 
    @sea_cortesias   int,
    @Erro            int OUTPUT,
    @MsgErr          varchar(255) OUTPUT)
AS 
   DECLARE @cxa_cd_aux  int
   
   SELECT @cxa_cd_aux = cxa_cd
     FROM TB_SESSAO_AUX
    WHERE sal_cd      = @sal_cd
      AND fil_cd      = @fil_cd
      AND sre_data    = @sre_data
      AND sre_horario = @sre_horario
      AND cxa_cd      = @cxa_cd 
   
   IF @cxa_cd_aux IS NULL
      INSERT INTO TB_SESSAO_AUX (sal_cd, fil_cd, sre_data, sre_horario, cxa_cd, sea_lugares_sel,
                                 sea_inteiras, sea_meias, sea_cortesias)
      VALUES (@sal_cd, @fil_cd, @sre_data, @sre_horario, @cxa_cd, @sea_lugares_sel,
              @sea_inteiras, @sea_meias, @sea_cortesias)
   ELSE
      UPDATE TB_SESSAO_AUX 
         SET sea_lugares_sel = @sea_lugares_sel,
             sea_inteiras    = @sea_inteiras, 
	     sea_meias	     = @sea_meias, 
	     sea_cortesias   = @sea_cortesias
       WHERE sal_cd      = @sal_cd
         AND fil_cd      = @fil_cd
         AND sre_data    = @sre_data
         AND sre_horario = @sre_horario
         AND cxa_cd      = @cxa_cd 
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_SESSAO_AUX_U_P
	(@sal_cd	int,
	 @fil_cd	int,
	 @sre_data	datetime,
	 @sre_horario	datetime,
	 @sap_mat_poltr varchar(2704),
	 @lcaixa        varchar(1),
         @Erro		int OUTPUT,
         @MsgErr	varchar(255) OUTPUT)
AS 
 DECLARE @sap_mat_poltr_aux varchar(2704),
         @n                 int,
         @i                 int
 
 SELECT @sap_mat_poltr_aux = sap_mat_poltr
 FROM TB_SESSAO_AUX_P
 WHERE sal_cd      = @sal_cd
 AND   fil_cd      = @fil_cd
 AND   sre_data    = @sre_data
 AND   sre_horario = @sre_horario
 
 IF @sap_mat_poltr_aux IS NULL
     INSERT TB_SESSAO_AUX_P 
     (sal_cd, fil_cd, sre_data, sre_horario, sap_mat_poltr)
     VALUES
     (@sal_cd, @fil_cd, @sre_data, @sre_horario, @sap_mat_poltr)
 ELSE
    BEGIN
      SELECT @n = LEN(RTRIM(@sap_mat_poltr_aux))
      SELECT @i = 1
 
      WHILE @i <= @n
         BEGIN
            IF SUBSTRING(@sap_mat_poltr_aux,@i,1) <> '8' AND SUBSTRING(@sap_mat_poltr,@i,1) = @lcaixa
               BEGIN
                  SELECT @sap_mat_poltr_aux = SUBSTRING(@sap_mat_poltr_aux, 1, @i-1) + @lcaixa + SUBSTRING(@sap_mat_poltr_aux, @i+1, @n)
               END
     
            IF SUBSTRING(@sap_mat_poltr_aux,@i,1) = @lcaixa AND SUBSTRING(@sap_mat_poltr,@i,1) <> @lcaixa
               BEGIN
                  SELECT @sap_mat_poltr_aux = SUBSTRING(@sap_mat_poltr_aux, 1, @i-1) + SUBSTRING(@sap_mat_poltr,@i,1) + SUBSTRING(@sap_mat_poltr_aux, @i+1, @n)
               END
               
            SELECT @i = @i + 1
         END
     
      UPDATE TB_SESSAO_AUX_P 
      SET sap_mat_poltr = @sap_mat_poltr_aux
      WHERE sal_cd      = @sal_cd
      AND   fil_cd      = @fil_cd
      AND   sre_data    = @sre_data
      AND   sre_horario = @sre_horario
   END
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_SESSAO_AUX_P_P
   (@sal_cd        int,
    @fil_cd        int,
    @sre_data      datetime,
    @sre_horario   datetime,
    @sap_mat_poltr varchar(2704) OUTPUT,
    @Erro          int OUTPUT,
    @MsgErr        varchar(255) OUTPUT)
AS
   DECLARE @aux1  int,
           @aux2  int,
           @aux3  int
   
   SELECT  @sap_mat_poltr = sap_mat_poltr
   FROM TB_SESSAO_AUX_P
   WHERE sal_cd      = @sal_cd
   AND   fil_cd      = @fil_cd
   AND   sre_data    = @sre_data
   AND   sre_horario = @sre_horario
      
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

   IF @sap_mat_poltr IS NULL
      BEGIN
         SELECT  @sap_mat_poltr = sre_mat_poltr
         FROM TB_SESSAO_REAL
         WHERE sal_cd      = @sal_cd
         AND   fil_cd      = @fil_cd
         AND   sre_data    = @sre_data
         AND   sre_horario = @sre_horario
            
         SELECT @Erro = @@ERROR
   
         IF @Erro <> 0
            BEGIN
               SELECT @MsgErr = description
               FROM master..sysmessages
               WHERE error = @Erro
         
              RETURN
           END
      END
      
GO


--****************************************************
CREATE PROCEDURE upTB_SESSAO_AUX_L_P
   (@sal_cd        int,
    @fil_cd        int,
    @sre_data      datetime,
    @sre_horario   datetime,
    @lcaixa        varchar(1),
    @Erro          int OUTPUT,
    @MsgErr        varchar(255) OUTPUT)
AS
 DECLARE @sre_mat_poltr_aux varchar(2704),
         @sap_mat_poltr_aux varchar(2704),
         @n                 int,
         @i                 int
 
 SELECT @sre_mat_poltr_aux = sre_mat_poltr
 FROM TB_SESSAO_REAL
 WHERE sal_cd      = @sal_cd
 AND   fil_cd      = @fil_cd
 AND   sre_data    = @sre_data
 AND   sre_horario = @sre_horario
 
 SELECT @sap_mat_poltr_aux = sap_mat_poltr
 FROM TB_SESSAO_AUX_P
 WHERE sal_cd      = @sal_cd
 AND   fil_cd      = @fil_cd
 AND   sre_data    = @sre_data
 AND   sre_horario = @sre_horario

 SELECT @n = LEN(RTRIM(@sre_mat_poltr_aux))
 SELECT @i = 1
 
 WHILE @i <= @n
    BEGIN

       IF SUBSTRING(@sap_mat_poltr_aux,@i,1) = @lcaixa
          BEGIN
             SELECT @sap_mat_poltr_aux = SUBSTRING(@sap_mat_poltr_aux, 1, @i-1) + SUBSTRING(@sre_mat_poltr_aux,@i,1) + SUBSTRING(@sap_mat_poltr_aux, @i+1, @n)
          END
          
       SELECT @i = @i + 1
    END

 UPDATE TB_SESSAO_AUX_P 
 SET sap_mat_poltr = @sap_mat_poltr_aux
 WHERE sal_cd      = @sal_cd
 AND   fil_cd      = @fil_cd
 AND   sre_data    = @sre_data
 AND   sre_horario = @sre_horario
      
 SELECT @Erro = @@ERROR
   
 IF @Erro <> 0
    BEGIN
       SELECT @MsgErr = description
       FROM master..sysmessages
       WHERE error = @Erro
       
       RETURN
    END

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SESSAO_REAL_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SESSAO_REAL_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SESSAO_REAL_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SESSAO_REAL_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SESSAO_REAL_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SESSAO_REAL_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SESSAO_REAL_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SESSAO_REAL_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SESSAO_REAL_S1') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SESSAO_REAL_S1
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SESSAO_REAL_U1') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SESSAO_REAL_U1
GO

--****************************************************
CREATE PROCEDURE upTB_SESSAO_REAL_D
	(@sal_cd	 int,
	 @fil_cd	 int,
	 @sre_data	 datetime,
	 @sre_horario	 datetime,
	 @TipoExclusao   char(1),
         @Erro           int OUTPUT,
         @MsgErr         varchar(255) OUTPUT)
AS
    IF @TipoExclusao = 'L'
	 UPDATE TB_SESSAO_REAL 
	    SET sre_dt_des = GETDATE()
	  WHERE sal_cd      = @sal_cd
	    AND fil_cd      = @fil_cd
	    AND sre_data    = @sre_data
	    AND sre_horario = @sre_horario
    ELSE
	 DELETE TB_SESSAO_REAL 
	  WHERE sal_cd      = @sal_cd
	    AND fil_cd      = @fil_cd
	    AND sre_data    = @sre_data
	    AND sre_horario = @sre_horario

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_SESSAO_REAL_I
	(@sal_cd	int,
	 @fil_cd	int,
	 @sre_data	datetime,
	 @sre_horario	datetime,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
--    SET DATEFIRST 1
    DECLARE @count          int,
            @lugares        int,
            @ses_dia_semana int,
            @sre_poltronas  int,
            @sre_mat_poltr  varchar(2702)
    
    
    SELECT @count = count(*)
      FROM TB_SESSAO_REAL
     WHERE sal_cd      = @sal_cd
       AND fil_cd      = @fil_cd
       AND sre_data    = @sre_data
       AND sre_horario = @sre_horario
       
    IF @count = 0
	BEGIN
		SELECT @lugares = sal_lugares 
		FROM TB_SALA_LUGAR
		WHERE sal_cd = @sal_cd
		AND @sre_data BETWEEN sal_dt_ini AND sal_dt_fim
		
		IF @lugares IS NULL
		   SELECT @lugares = sal_lugares
		   FROM TB_SALA
		   WHERE sal_cd = @sal_cd
		   AND sal_dt_des IS NULL

		SELECT @sre_poltronas = pol_poltronas,
		       @sre_mat_poltr = pol_mat_poltr
		FROM tb_poltronas
		WHERE sal_cd = @sal_cd
		
		SELECT @count = COUNT(*)
		FROM TB_FERIADO
		WHERE fer_data = convert(datetime, @sre_data, 103)
		
		IF @count > 0 
		   SELECT @ses_dia_semana = 8
		ELSE
		   SELECT @ses_dia_semana = datepart(dw,convert(datetime, @sre_data, 103))

		INSERT INTO TB_SESSAO_REAL 
		     (sal_cd, fil_cd, sre_data, sre_horario, sre_lugares, sre_lugares_vendidos,
		      sre_inteiras, sre_meias, sre_cortesias, sre_pre_estreia, sre_poltronas, sre_mat_poltr) 
		SELECT sal_cd, fil_cd, @sre_data, ses_horario, @lugares, 
		       0, 0, 0, 0, ses_pre_estreia, @sre_poltronas, @sre_mat_poltr
		FROM TB_SESSAO
		WHERE prg_cd = (SELECT prg_cd 
		                FROM TB_PROGRAMACAO
		                WHERE @sre_data BETWEEN prg_dt_ini AND prg_dt_fim 
		                AND   prg_dt_des IS NULL )
		AND sal_cd         = @sal_cd
		AND fil_cd         = @fil_cd
		AND ses_dia_semana = @ses_dia_semana
		AND ses_horario    = @sre_horario
		AND ses_dt_des IS NULL
	   END
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_SESSAO_REAL_S
	(@sal_cd	int,
	 @fil_cd	int,
	 @sre_data	datetime,
	 @sre_horario	datetime,
	 @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
    IF @sal_cd IS NULL AND @fil_cd IS NULL AND @sre_data IS NULL AND @sre_horario IS NULL
       SELECT * 
         FROM TB_SESSAO_REAL
        WHERE sre_dt_des IS NULL
    ELSE
        IF @sal_cd IS NOT NULL AND @fil_cd IS NOT NULL AND @sre_data IS NOT NULL AND @sre_horario IS NOT NULL
            SELECT * 
              FROM TB_SESSAO_REAL
             WHERE sal_cd      = @sal_cd
               AND fil_cd      = @fil_cd
               AND sre_data    = @sre_data
               AND sre_horario = @sre_horario
        ELSE
            IF @sal_cd IS NOT NULL AND @fil_cd IS NOT NULL AND @sre_data IS NOT NULL AND @sre_horario IS NULL
                 SELECT * 
                   FROM TB_SESSAO_REAL
                  WHERE sal_cd   = @sal_cd
                    AND fil_cd   = @fil_cd
                    AND sre_data = @sre_data
             ELSE
                 IF @sal_cd IS NOT NULL AND @fil_cd IS NOT NULL AND @sre_data IS NULL AND @sre_horario IS NULL
                     SELECT * 
                       FROM TB_SESSAO_REAL
                      WHERE sal_cd = @sal_cd
                        AND fil_cd = @fil_cd
                 ELSE
                     IF @sal_cd IS NOT NULL AND @fil_cd IS NULL AND @sre_data IS NULL AND @sre_horario IS NULL
                         SELECT * 
                           FROM TB_SESSAO_REAL
                          WHERE sal_cd = @sal_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_SESSAO_REAL_U
	(@sal_cd		int,
	 @fil_cd		int,
	 @sre_data		datetime,
	 @sre_horario		datetime,
	 @sre_lugares_vendidos	smallint,
	 @sre_inteiras          int, 
	 @sre_meias             int, 
	 @sre_cortesias         int,
         @Erro			int OUTPUT,
         @MsgErr		varchar(255) OUTPUT)
AS 
 UPDATE TB_SESSAO_REAL 
 SET sre_lugares_vendidos = sre_lugares_vendidos + @sre_lugares_vendidos,
     sre_inteiras	  = sre_inteiras + @sre_inteiras, 
     sre_meias	          = sre_meias + @sre_meias, 
     sre_cortesias	  = sre_cortesias + @sre_cortesias
 WHERE sal_cd      = @sal_cd
 AND   fil_cd      = @fil_cd
 AND   sre_data    = @sre_data
 AND   sre_horario = @sre_horario
   
 SELECT @Erro = @@ERROR
   
 IF @Erro <> 0
    BEGIN
       SELECT @MsgErr = description
       FROM master..sysmessages
       WHERE error = @Erro
       
       RETURN
    END
      
GO

--****************************************************
CREATE PROCEDURE upTB_SESSAO_REAL_S1
	(@sal_cd	int,
	 @fil_cd	int,
	 @sre_data	datetime,
	 @sre_horario	datetime,
	 @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   DECLARE @sal_cdAux int
   
   SET NOCOUNT ON
   
   SELECT @sal_cdAux = NULL

   SELECT @sal_cdAux = sr.sal_cd
   FROM TB_SESSAO_REAL sr
   WHERE sr.sal_cd      = @sal_cd
   AND   sr.fil_cd      = @fil_cd
   AND   sr.sre_data    = @sre_data
   AND   sr.sre_horario = @sre_horario

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END
      
   IF @sal_cdAux IS NULL
      BEGIN
         EXECUTE upTB_SESSAO_REAL_I @sal_cd, @fil_cd, @sre_data, @sre_horario, @Erro, @MsgErr
         
         IF @Erro <> 0
            BEGIN
               SELECT @MsgErr = description
               FROM master..sysmessages
               WHERE error = @Erro
         
               RETURN
            END
      END
   
   SET NOCOUNT OFF

   SELECT s.sal_desc,
          f.fil_nm,
          p.pol_tp_numeracao,
          p.pol_num_pri_col,
          p.pol_num_filas,
          p.pol_num_colunas,
          p.pol_num_horiz,
          p.pol_num_vert,
          p.pol_poltronas,
          p.pol_mat_poltr,
          sr.sre_mat_poltr
   FROM TB_SESSAO_REAL sr,
        TB_POLTRONAS   p,
        TB_SALA        s,
        TB_FILME       f
   WHERE sr.sal_cd      = p.sal_cd
   AND   sr.sal_cd      = s.sal_cd
   AND   sr.fil_cd      = f.fil_cd
   AND   sr.sal_cd      = @sal_cd
   AND   sr.fil_cd      = @fil_cd
   AND   sr.sre_data    = @sre_data
   AND   sr.sre_horario = @sre_horario
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END
      
GO

--****************************************************
CREATE PROCEDURE upTB_SESSAO_REAL_U1
	(@sal_cd		int,
	 @fil_cd		int,
	 @sre_data		datetime,
	 @sre_horario		datetime,
	 @sre_mat_poltr         varchar(2704),
	 @lcaixa                varchar(1),
         @Erro			int OUTPUT,
         @MsgErr		varchar(255) OUTPUT)
AS 
 DECLARE @sre_mat_poltr_aux varchar(2704),
         @sap_mat_poltr_aux varchar(2704),
         @n                 int,
         @i                 int
 
 SELECT @sre_mat_poltr_aux = sre_mat_poltr
 FROM TB_SESSAO_REAL
 WHERE sal_cd      = @sal_cd
 AND   fil_cd      = @fil_cd
 AND   sre_data    = @sre_data
 AND   sre_horario = @sre_horario
 
 SELECT @sap_mat_poltr_aux = sap_mat_poltr
 FROM TB_SESSAO_AUX_P
 WHERE sal_cd      = @sal_cd
 AND   fil_cd      = @fil_cd
 AND   sre_data    = @sre_data
 AND   sre_horario = @sre_horario
 
 
 SELECT @n = LEN(RTRIM(@sre_mat_poltr_aux))
 SELECT @i = 1
 
 WHILE @i <= @n
    BEGIN
       IF SUBSTRING(@sre_mat_poltr_aux,@i,1) <> '8' AND SUBSTRING(@sre_mat_poltr,@i,1) = '9'
          BEGIN
             SELECT @sre_mat_poltr_aux = SUBSTRING(@sre_mat_poltr_aux, 1, @i-1) + '8' + SUBSTRING(@sre_mat_poltr_aux, @i+1, @n)
             SELECT @sap_mat_poltr_aux = SUBSTRING(@sap_mat_poltr_aux, 1, @i-1) + '8' + SUBSTRING(@sap_mat_poltr_aux, @i+1, @n)
          END
       --IF SUBSTRING(@sre_mat_poltr,@i,1) = '9'
       --   BEGIN
       --      SELECT @sap_mat_poltr_aux = SUBSTRING(@sap_mat_poltr_aux, 1, @i-1) + '8' + SUBSTRING(@sap_mat_poltr_aux, @i+1, @n)
       --   END
          
       SELECT @i = @i + 1
    END

 UPDATE TB_SESSAO_REAL 
 SET sre_mat_poltr        = @sre_mat_poltr_aux
 WHERE sal_cd      = @sal_cd
 AND   fil_cd      = @fil_cd
 AND   sre_data    = @sre_data
 AND   sre_horario = @sre_horario
   
 SELECT @Erro = @@ERROR
   
 IF @Erro <> 0
    BEGIN
       SELECT @MsgErr = description
       FROM master..sysmessages
       WHERE error = @Erro
       
       RETURN
    END
      
 UPDATE TB_SESSAO_AUX_P 
 SET sap_mat_poltr = @sap_mat_poltr_aux
 WHERE sal_cd      = @sal_cd
 AND   fil_cd      = @fil_cd
 AND   sre_data    = @sre_data
 AND   sre_horario = @sre_horario
      
 SELECT @Erro = @@ERROR
   
 IF @Erro <> 0
    BEGIN
       SELECT @MsgErr = description
       FROM master..sysmessages
       WHERE error = @Erro
       
       RETURN
    END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_SIS_LOG_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_SIS_LOG_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_LOG_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_LOG_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upEmpresa_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upEmpresa_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upCinema_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upCinema_S
GO

--****************************************************
CREATE PROCEDURE dbo.upTB_SIS_LOG_I
	(@usu_nm	varchar(50),
	 @slg_descricao	varchar(4000),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 

declare @emp_cd   int,
        @cin_cd   int

SELECT @emp_cd = emp_cd
FROM tb_empresa

SELECT @Erro = @@ERROR

IF @Erro <> 0
   BEGIN
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro

      RETURN
   END


SELECT @cin_cd = cin_cd
FROM tb_cinema

SELECT @Erro = @@ERROR

IF @Erro <> 0
   BEGIN
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro

      RETURN
   END

INSERT INTO tb_sis_log
        (slg_data,
         emp_cd,
         cin_cd,
         usu_nm,
         slg_descricao)
VALUES 
	(GETDATE(),
         @emp_cd,
         @cin_cd,
         @usu_nm,
         @slg_descricao)
	 
   
SELECT @Erro = @@ERROR

IF @Erro <> 0
   BEGIN
      SELECT @MsgErr = description
      FROM master..sysmessages
      WHERE error = @Erro

      RETURN
   END

GO

--****************************************************
CREATE PROCEDURE dbo.upTB_LOG_S
	(@dt_ini datetime,
	 @dt_fim datetime,
	 @emp_cd int,
	 @cin_cd int)
AS 

SELECT tb_sis_log.slg_data,
       tb_sis_log.usu_nm,
       tb_sis_log.slg_descricao
FROM tb_sis_log
WHERE tb_sis_log.emp_cd = @emp_cd
AND   tb_sis_log.cin_cd = @cin_cd
AND   tb_sis_log.slg_data BETWEEN @dt_ini AND @dt_fim
ORDER BY tb_sis_log.slg_data

GO

--****************************************************
CREATE PROCEDURE dbo.upEmpresa_S
AS 

SELECT DISTINCT emp_cd,
                emp_nm
FROM tb_bol_empr
ORDER BY emp_nm

GO

--****************************************************
CREATE PROCEDURE dbo.upCinema_S (@emp_cd int)
AS 

SELECT DISTINCT cin_cd,
                cin_nm
FROM tb_bol_cin
WHERE emp_cd = @emp_cd
ORDER BY cin_nm

GO


SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_USUARIO_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_USUARIO_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_USUARIO_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_USUARIO_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_USUARIO_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_USUARIO_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_USUARIO_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_USUARIO_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_USUARIO_LOGIN_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_USUARIO_LOGIN_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_USUARIO_SENHA_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_USUARIO_SENHA_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_USUARIO_GRID') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_USUARIO_GRID
GO

--****************************************************
CREATE PROCEDURE upTB_USUARIO_D
	(@usu_cd 	int,
         @TipoExclusao  char(1),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
    DECLARE @n int

    SELECT @n = COUNT(*)
    FROM TB_CAIXA_MOVTO
    WHERE usu_abertura = @usu_cd
    AND   cxp_status <> 2
    
    IF @n > 0 
      BEGIN
         SELECT @MsgErr = 'Existem caixas abertos para este usuário'
         SELECT @Erro   = 99
         
         RETURN
      END

    IF @TipoExclusao = 'L'
	 UPDATE TB_USUARIO 
	    SET usu_dt_des = GETDATE()
	  WHERE usu_cd	= @usu_cd
    ELSE
	BEGIN
            DELETE TB_USUARIO_PERFIL
             WHERE usu_cd = @usu_cd
	    DELETE TB_USUARIO 
	     WHERE usu_cd = @usu_cd
	END

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_USUARIO_GRID
	(@Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   SELECT a.usu_cd, a.usu_senha, b.cin_cd, d.per_cd, 
          b.cin_nm 'Cinema', a.usu_nm 'Usuário', a.usu_login 'Login', d.per_desc 'Perfil'
     FROM TB_USUARIO a,
          TB_CINEMA b,
          TB_USUARIO_PERFIL c,
          TB_PERFIL_ACESSO d
    WHERE a.cin_cd = b.cin_cd
      AND a.usu_cd = c.usu_cd
      AND c.per_cd = d.per_cd
      AND a.usu_dt_des IS NULL
      AND b.cin_dt_des IS NULL
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_USUARIO_I
	(@cin_cd 	int,
         @usu_nm 	varchar(50),
	 @per_cd	int,
         @usu_senha     char(31),
         @usu_login     varchar(20),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
DECLARE @usu_cd int
INSERT INTO TB_USUARIO 
	 ( cin_cd,
         usu_nm,
	 usu_login,
	 usu_senha,
 	 usu_dt_inc) 
VALUES 
	( @cin_cd,
         @usu_nm,
	 @usu_login,
	 @usu_senha,
	 GETDATE())
select @usu_cd = @@IDENTITY
INSERT INTO TB_USUARIO_PERFIL
	( usu_cd, 
	  per_cd )
VALUES
	( @usu_cd, 
	  @per_cd )
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_USUARIO_LOGIN_S
	(@usu_login     varchar(20),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   SELECT a.usu_cd, a.usu_nm, a.usu_senha, b.per_cd
     FROM TB_USUARIO a,
          TB_USUARIO_PERFIL b
    WHERE a.usu_cd = b.usu_cd
      AND a.usu_dt_des IS NULL
      AND a.usu_login = @usu_login
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         	
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_USUARIO_S
	(@usu_cd 	int,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   IF @usu_cd IS NULL
       SELECT a.*
         FROM TB_USUARIO a,
              TB_CINEMA b,
              TB_USUARIO_PEFIL c,
              TB_PEFIL d
        WHERE a.cin_cd = b.cin_cd
	  AND a.usu_cd = c.usu_cd
 	  AND c.per_cd = d.per_cd
          AND a.usu_dt_des IS NULL
          AND b.cin_dt_des IS NULL
   ELSE
       SELECT a.*
         FROM TB_USUARIO a,
              TB_CINEMA b,
              TB_USUARIO_PERFIL c,
              TB_PERFIL_ACESSO d
        WHERE a.cin_cd = b.cin_cd
	  AND a.usu_cd = c.usu_cd
 	  AND c.per_cd = d.per_cd
          AND a.usu_dt_des IS NULL
          AND b.cin_dt_des IS NULL
	  AND a.usu_cd = @usu_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         	
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_USUARIO_SENHA_U
	(@usu_cd 	int,
	 @usu_senha     char(31),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
 UPDATE TB_USUARIO 
    SET usu_senha = @usu_senha,
	usu_dt_alt = GETDATE()
  WHERE usu_cd	= @usu_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_USUARIO_U
	(@usu_cd 	int,
	 @cin_cd 	int,
         @usu_nm 	varchar(50),
	 @per_cd	int,
         @usu_senha     char(31),
         @usu_login     varchar(20),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
 UPDATE TB_USUARIO 
    SET cin_cd = @cin_cd,
        usu_nm = @usu_nm,
        usu_login = @usu_login,
        usu_senha = @usu_senha,
	usu_dt_alt = GETDATE()
  WHERE usu_cd	= @usu_cd
 UPDATE TB_USUARIO_PERFIL
    SET per_cd = @per_cd
  WHERE usu_cd = @usu_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_VENDA_COMBO_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_VENDA_COMBO_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_VENDA_COMBO_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_VENDA_COMBO_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_VENDA_COMBO_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_VENDA_COMBO_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_VENDA_COMBO_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_VENDA_COMBO_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_VENDA_COMBO_GRID') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_VENDA_COMBO_GRID
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_VENDA_COMBO_CANCEL') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_VENDA_COMBO_CANCEL
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_VENDA_COMBO_VER_COMBO') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_VENDA_COMBO_VER_COMBO
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_VENDA_COMBO_UTIL') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_VENDA_COMBO_UTIL
GO

--****************************************************
CREATE PROCEDURE upTB_VENDA_COMBO_CANCEL
	(@vcb_cd	bigint,
	 @vcb_mot_canc	varchar(50),
         @Erro        	int OUTPUT,
         @MsgErr       	varchar(255) OUTPUT)
AS 
 UPDATE TB_VENDA_COMBO 
    SET vcb_dt_canc = GETDATE(),
        vcb_mot_canc = @vcb_mot_canc,
	vcb_status = 9
  WHERE vcb_cd = @vcb_cd

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_VENDA_COMBO_D
	(@vcb_cd	 bigint,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
   DELETE TB_VENDA_COMBO 
    WHERE vcb_cd = @vcb_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_VENDA_COMBO_GRID
	(@vcb_cd	bigint,
         @ope_cd 	bigint,
         @num_cd        bigint,
         @tpPes         varchar(1),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
    IF @tpPes IS NOT NULL
       BEGIN
          SELECT @vcb_cd = NULL
          
          IF @tpPes = 'O'
             BEGIN
                SELECT @ope_cd = @num_cd
             END
          ELSE
             BEGIN
                IF @tpPes = 'I'
                   BEGIN
                      SELECT @ope_cd = ope_cd
                      FROM tb_venda_ingresso
                      WHERE ing_cd = @num_cd

                      SELECT @Erro = @@ERROR

                      IF @Erro <> 0
                         BEGIN
                            SELECT @MsgErr = description
                            FROM master..sysmessages
                            WHERE error = @Erro
         
                            RETURN
                         END
                   
                   END
                ELSE
                   BEGIN
                      IF @tpPes = 'C'
                         BEGIN
                            SELECT @ope_cd = ope_cd
                            FROM tb_venda_combo
                            WHERE vcb_cd = @num_cd

                            SELECT @Erro = @@ERROR

                            IF @Erro <> 0
                               BEGIN
                                  SELECT @MsgErr = description
                                  FROM master..sysmessages
                                  WHERE error = @Erro
         
                                  RETURN
                               END
                   
                         END
                      ELSE
                         BEGIN
                            SELECT @MsgErr = 'Opção invalida'
                            SELECT @Erro   = 1
                      
                            RETURN
                         END
                         
                   END
             END
       END

    IF @vcb_cd IS NOT NULL
       SELECT a.vcb_cd 'Nº Combo', 
	      a.ope_cd 'Operação', 
	      a.cbo_cd, 
	      c.cbo_desc 'Combo',
	      a.vcb_status, 
	      convert(char(10),b.ope_dt_operacao,103) 'Data da Venda', 
	      a.vcb_valor * a.vcb_qtde 'Valor',
	      a.vcb_cd
         FROM TB_VENDA_COMBO a,
	      TB_OPERACAO b,
	      TB_COMBO c
	WHERE b.ope_cd = a.ope_cd
	  AND c.cbo_cd = a.cbo_cd
          AND a.vcb_cd = @vcb_cd
          AND a.vcb_dt_canc IS NULL
    ELSE
       SELECT a.vcb_cd 'Nº Combo', 
	      a.ope_cd 'Operação', 
	      a.cbo_cd, 
	      c.cbo_desc 'Combo',
	      a.vcb_status, 
	      convert(char(10),b.ope_dt_operacao,103) 'Data da Venda', 
	      a.vcb_valor * a.vcb_qtde 'Valor', 
	      a.vcb_cd
         FROM TB_VENDA_COMBO a,
	      TB_OPERACAO b,
	      TB_COMBO c
	WHERE b.ope_cd = a.ope_cd
	  AND c.cbo_cd = a.cbo_cd
          AND a.ope_cd = @ope_cd
          AND a.vcb_dt_canc IS NULL
    
    SELECT @Erro = @@ERROR
   
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          RETURN
       END

GO

--****************************************************
CREATE PROCEDURE upTB_VENDA_COMBO_I
	(@vcb_cd	bigint,
         @ope_cd 	int,
	 @cbo_cd	int,
	 @vcb_status    tinyint,
         @vcb_qtde	smallint,
	 @vcb_valor	money,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
INSERT INTO TB_VENDA_COMBO 
	 ( vcb_cd,
	 ope_cd,
	 cbo_cd,
	 vcb_status,
         vcb_qtde,
 	 vcb_valor) 
VALUES 
	( @vcb_cd,
	 @ope_cd,
	 @cbo_cd,
	 @vcb_status,
         @vcb_qtde,
 	 @vcb_valor) 
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
         
CREATE PROCEDURE upTB_VENDA_COMBO_S
	(@vcb_cd	bigint,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
    IF @vcb_cd IS NULL
       SELECT *
         FROM TB_VENDA_COMBO
    ELSE
       SELECT *
         FROM TB_VENDA_COMBO
	WHERE vcb_cd = @vcb_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_VENDA_COMBO_U
	(@vcb_cd	bigint,
         @ope_cd 	int,
	 @cbo_cd	int,
	 @vcb_status    tinyint,
         @vcb_qtde	smallint,
	 @vcb_valor	money,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS 
 UPDATE TB_VENDA_COMBO 
    SET vcb_status = @vcb_status,
	vcb_qtde = @vcb_qtde,
	vcb_valor = @vcb_valor,
	ope_cd	= @ope_cd,
	cbo_cd = @cbo_cd
  WHERE vcb_cd = @vcb_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_VENDA_COMBO_UTIL
	(@vcb_cd	bigint,
         @Erro        	int OUTPUT,
         @MsgErr       	varchar(255) OUTPUT)
AS 
 UPDATE TB_VENDA_COMBO 
    SET vcb_status = 1
  WHERE vcb_cd = @vcb_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_VENDA_COMBO_VER_COMBO
	(@vcb_cd	bigint,
         @Erro        	int OUTPUT,
         @MsgErr       	varchar(255) OUTPUT)
AS 
   declare @par_hora_max_ses datetime,
           @ope_dt_operacao datetime,
           @vcb_status int,
           @data1 datetime,
           @data2 datetime,
           @data3 datetime
           
   SELECT  @par_hora_max_ses = par_hora_max_ses
   FROM TB_PARAMETRO
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END
           
   SELECT  @ope_dt_operacao = TB_OPERACAO.ope_dt_operacao,
           @vcb_status = TB_VENDA_COMBO.vcb_status
   FROM TB_VENDA_COMBO,
        TB_OPERACAO
   WHERE TB_VENDA_COMBO.ope_cd = TB_OPERACAO.ope_cd
   AND   TB_OPERACAO.ope_dt_des IS NULL
   AND   TB_VENDA_COMBO.vcb_cd = @vcb_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END
   IF @vcb_status IS NULL
      BEGIN
         SELECT @MsgErr = 'Combo inexistente'
         SELECT @Erro = 3
         
         RETURN
      END
      
   IF @vcb_status <> 0
      BEGIN
         SELECT @MsgErr = 'Combo utilizado'
         SELECT @Erro = 1
         
         RETURN
      END
      
   SELECT @data1 = CONVERT(datetime, CONVERT(char(8), @ope_dt_operacao, 108), 108);
   SELECT @data2 = CONVERT(datetime, CONVERT(char(10), @ope_dt_operacao, 103), 103);
   SELECT @data3 = CONVERT(datetime, CONVERT(char(10), GETDATE(), 103), 103);
   SELECT @par_hora_max_ses = CONVERT(datetime, CONVERT(char(8), @par_hora_max_ses, 108), 108);
   IF @data2 <> @data3 
      BEGIN
         IF @data1 < @par_hora_max_ses
            BEGIN
               SELECT @data3 = DATEADD(Day, -1, @data3)
               
               IF @data2 <> @data3
                  BEGIN
                     SELECT @MsgErr = 'Combo de data invalida'
                     SELECT @Erro = 2
          
                     RETURN
                  END
            END
         ELSE
            BEGIN
               SELECT @MsgErr = 'Combo de data invalida'
               SELECT @Erro = 2
      
               RETURN
            END
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_VENDA_INGRESSO_D') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_VENDA_INGRESSO_D
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_VENDA_INGRESSO_I') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_VENDA_INGRESSO_I
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_VENDA_INGRESSO_U') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_VENDA_INGRESSO_U
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_VENDA_INGRESSO_S') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_VENDA_INGRESSO_S
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_VENDA_INGRESSO_GRID') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_VENDA_INGRESSO_GRID
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_VENDA_INGRESSO_CANCEL') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_VENDA_INGRESSO_CANCEL
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_VENDA_INGRESSO_VER_INGR') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_VENDA_INGRESSO_VER_INGR
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upTB_VENDA_INGRESSO_UTIL') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upTB_VENDA_INGRESSO_UTIL
GO

--****************************************************
CREATE PROCEDURE upTB_VENDA_INGRESSO_CANCEL
	(@ing_cd	bigint,
	 @ing_mot_canc	varchar(50),
         @Erro        	int OUTPUT,
         @MsgErr       	varchar(255) OUTPUT)
AS 
 UPDATE TB_VENDA_INGRESSO 
    SET ing_dt_canc = GETDATE(),
        ing_mot_canc = @ing_mot_canc,
	ing_status = 9
  WHERE ing_cd = @ing_cd

   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_VENDA_INGRESSO_D
	(@ing_cd	 bigint,
	 @Erro           int OUTPUT,
         @MsgErr         varchar(255) OUTPUT)
AS
   DELETE TB_VENDA_INGRESSO 
    WHERE ing_cd = @ing_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_VENDA_INGRESSO_GRID
	(@ing_cd	bigint,
         @ope_cd 	bigint,
         @num_cd        bigint,
         @tpPes         varchar(1),
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
    IF @tpPes IS NOT NULL
       BEGIN
          SELECT @ing_cd = NULL
          
          IF @tpPes = 'O'
             BEGIN
                SELECT @ope_cd = @num_cd
             END
          ELSE
             BEGIN
                IF @tpPes = 'I'
                   BEGIN
                      SELECT @ope_cd = ope_cd
                      FROM tb_venda_ingresso
                      WHERE ing_cd = @num_cd

                      SELECT @Erro = @@ERROR

                      IF @Erro <> 0
                         BEGIN
                            SELECT @MsgErr = description
                            FROM master..sysmessages
                            WHERE error = @Erro
         
                            RETURN
                         END
                   
                   END
                ELSE
                   BEGIN
                      IF @tpPes = 'C'
                         BEGIN
                            SELECT @ope_cd = ope_cd
                            FROM tb_venda_combo
                            WHERE vcb_cd = @num_cd

                            SELECT @Erro = @@ERROR

                            IF @Erro <> 0
                               BEGIN
                                  SELECT @MsgErr = description
                                  FROM master..sysmessages
                                  WHERE error = @Erro
         
                                  RETURN
                               END
                   
                         END
                      ELSE
                         BEGIN
                            SELECT @MsgErr = 'Opção invalida'
                            SELECT @Erro   = 1
                      
                            RETURN
                         END
                         
                   END
             END
       END

    IF @ing_cd IS NOT NULL
       SELECT a.ing_cd 'Nº Ingresso', 
	      a.ope_cd 'Operação', 
	      b.sal_desc 'Sala',
	      c.fil_nm 'Filme', 
	      a.sre_data 'Data', 
	      convert(char(10),a.sre_horario,108) 'Sessão', 
	      convert(char(10),a.ing_dt_venda,103) 'Data da Venda', 
	      a.ing_valor 'Valor', 
	      a.sal_cd, 
	      a.fil_cd, 
	      a.sre_data, 
	      a.sre_horario,
	      a.ing_cd,
	      a.igt_cd,
	      a.ing_status, 
	      a.ing_num_talao
         FROM TB_VENDA_INGRESSO a,
	      TB_SALA b,
	      TB_FILME c
	WHERE a.sal_cd = b.sal_cd
	  AND c.fil_cd = a.fil_cd
          AND a.ing_cd = @ing_cd
          AND a.ing_dt_canc IS NULL
    ELSE
       SELECT a.ing_cd 'Nº Ingresso', 
	      a.ope_cd 'Operação', 
	      b.sal_desc 'Sala',
	      c.fil_nm 'Filme', 
	      a.sre_data 'Data', 
	      convert(char(10),a.sre_horario,108) 'Sessão', 
	      convert(char(10),a.ing_dt_venda,103) 'Data da Venda', 
	      a.ing_valor 'Valor', 
	      a.sal_cd, 
	      a.fil_cd, 
	      a.sre_data, 
	      a.sre_horario,
	      a.ing_cd,
	      a.igt_cd,
	      a.ing_status, 
	      a.ing_num_talao
         FROM TB_VENDA_INGRESSO a,
	      TB_SALA b,
	      TB_FILME c
	WHERE a.sal_cd = b.sal_cd
	  AND c.fil_cd = a.fil_cd
          AND a.ope_cd = @ope_cd
          AND a.ing_dt_canc IS NULL

    SELECT @Erro = @@ERROR
   
    IF @Erro <> 0
       BEGIN
          SELECT @MsgErr = description
          FROM master..sysmessages
          WHERE error = @Erro
         
          RETURN
       END

GO

--****************************************************
CREATE PROCEDURE upTB_VENDA_INGRESSO_I
	(@ing_cd	bigint,
	 @ope_cd	int,
	 @sal_cd	int,
	 @fil_cd	int,
	 @sre_data	datetime,
	 @sre_horario	datetime,
	 @ing_status	tinyint,
	 @igt_cd	smallint,
	 @ing_valor	money,
	 @ppr_cd	int,
	 @ing_num_talao	bigint,
	 @ing_num_ing   varchar(4),
	 @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
 
AS 
INSERT INTO TB_VENDA_INGRESSO 
	 (ing_cd,
	  ope_cd,
	  sal_cd,
	  fil_cd,
	  sre_data,
	  sre_horario,
	  ing_status,
	  ing_dt_venda,
	  igt_cd,
	  ing_valor,
	  ppr_cd,
	  ing_num_talao,
	  ing_num_ing) 
VALUES 
	(@ing_cd,
	 @ope_cd,
	 @sal_cd,
	 @fil_cd,
	 @sre_data,
	 @sre_horario,
	 @ing_status,
	 GETDATE(),
	 @igt_cd,
	 @ing_valor,
	 @ppr_cd,
	 @ing_num_talao,
	 @ing_num_ing)
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_VENDA_INGRESSO_S
	(@ing_cd	bigint,
         @Erro          int OUTPUT,
         @MsgErr        varchar(255) OUTPUT)
AS
    IF @ing_cd IS NULL
       SELECT tb_venda_ingresso.*
         FROM tb_venda_ingresso
    ELSE
       SELECT tb_venda_ingresso.*,
              tb_ingresso_tipo.igt_desc,
              tb_sala.sal_desc,
              tb_filme.fil_nm
         FROM tb_venda_ingresso,
              tb_ingresso_tipo,
              tb_sala,
              tb_filme
	WHERE tb_venda_ingresso.ing_cd = @ing_cd
	AND   tb_venda_ingresso.igt_cd = tb_ingresso_tipo.igt_cd
	AND   tb_venda_ingresso.sal_cd = tb_sala.sal_cd
	AND   tb_venda_ingresso.fil_cd = tb_filme.fil_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_VENDA_INGRESSO_U
	(@ing_cd	bigint,
	 @ope_cd	int,
	 @sal_cd	int,
	 @fil_cd	int,
	 @sre_data	datetime,
	 @sre_horario	datetime,
	 @ing_status	tinyint,
	 @igt_cd	smallint,
	 @ing_valor	money,
	 @ppr_cd	int,
	 @ing_num_talao	bigint,
	 @Erro        	int OUTPUT,
         @MsgErr       	varchar(255) OUTPUT)
AS 
 UPDATE TB_VENDA_INGRESSO 
    SET ope_cd        = @ope_cd,
	sal_cd        = @sal_cd,
	fil_cd        = @fil_cd,
	sre_data      = @sre_data,
	sre_horario   = @sre_horario,
	ing_status    = @ing_status,
	ing_valor     = @ing_valor,
	ppr_cd        = @ppr_cd,
	ing_num_talao = @ing_num_talao
  WHERE ing_cd      = @ing_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_VENDA_INGRESSO_UTIL
	(@ing_cd	bigint,
         @Erro        	int OUTPUT,
         @MsgErr       	varchar(255) OUTPUT)
AS 
   declare @sre_data datetime,
           @sre_horario datetime,
           @ing_status int,
           @data1 datetime,
           @data2 datetime,
           @data3 datetime
 UPDATE TB_VENDA_INGRESSO 
    SET ing_dt_util = GETDATE(),
	ing_status = 1
  WHERE ing_cd = @ing_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END

GO

--****************************************************
CREATE PROCEDURE upTB_VENDA_INGRESSO_VER_INGR
	(@ing_cd     bigint,
	 @cat_cd     int,
	 @tempAntes  varchar(5),
	 @tempDepois varchar(5),
         @Erro       int OUTPUT,
         @MsgErr     varchar(255) OUTPUT)
AS 
   declare @sre_data         datetime,
           @sre_horario      datetime,
           @ing_status       int,
           @data1            datetime,
           @data2            datetime,
           @data3            datetime,
           @par_hora_max_ses datetime
           
   SELECT  @par_hora_max_ses = par_hora_max_ses
   FROM tb_parametro
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END
           
   SELECT  @sre_data = TB_VENDA_INGRESSO.sre_data,
           @sre_horario = TB_VENDA_INGRESSO.sre_horario,
           @ing_status = TB_VENDA_INGRESSO.ing_status
   FROM TB_VENDA_INGRESSO,
        TB_OPERACAO,
        TB_CATRACA_SALA
   WHERE TB_VENDA_INGRESSO.ope_cd = TB_OPERACAO.ope_cd
   AND   TB_VENDA_INGRESSO.sal_cd = TB_CATRACA_SALA.sal_cd
   AND   TB_OPERACAO.ope_dt_des IS NULL
   AND   TB_CATRACA_SALA.cat_cd   = @cat_cd
   AND   TB_VENDA_INGRESSO.ing_cd = @ing_cd
   
   SELECT @Erro = @@ERROR
   
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         RETURN
      END
   IF @ing_status IS NULL
      BEGIN
         SELECT @MsgErr = 'Ingresso inexistente ou catraca errada'
         SELECT @Erro = 3
         
         RETURN
      END
      
   IF @ing_status <> 0
      BEGIN
         SELECT @MsgErr = 'Ingresso cancelado ou utilizado'
         SELECT @Erro = 1
         
         RETURN
      END
      
   SELECT @data1    = GETDATE()
   SELECT @sre_horario = CONVERT(datetime, CONVERT(CHAR(8), @sre_horario, 108 ))
   SELECT @sre_data = CONVERT(datetime, CONVERT(CHAR(10), @sre_data, 103 ) + ' ' + CONVERT(CHAR(8), @sre_horario, 108 ), 103)
   SELECT @par_hora_max_ses = CONVERT(datetime, CONVERT(CHAR(8), @par_hora_max_ses, 108 ))

   IF @sre_horario <= @par_hora_max_ses
      BEGIN
         SELECT @sre_data = DATEADD(Day, 1, @sre_data)
      END
      
   SELECT @data2 = DATEADD(Hour, -1 * CONVERT(int, SUBSTRING(@tempAntes,1,2)), @sre_data)
   SELECT @data2 = DATEADD(minute, -1 * CONVERT(int, SUBSTRING(@tempAntes,4,2)), @data2)
   
   SELECT @data3 = DATEADD(Hour, CONVERT(int, SUBSTRING(@tempDepois,1,2)), @sre_data)
   SELECT @data3 = DATEADD(minute, CONVERT(int, SUBSTRING(@tempDepois,4,2)), @data3)
   
   IF @data1 < @data2 OR @data1 > @data3
      BEGIN
         SELECT @MsgErr = 'Horario invalido'
         SELECT @Erro = 4
         
         RETURN
      END

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upMOVTOS_TRANSF') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upMOVTOS_TRANSF
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upMOVTOS_PARA_TRANSF') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upMOVTOS_PARA_TRANSF
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upLIMPA_AUX_BOLETIM') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upLIMPA_AUX_BOLETIM
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upINCLUI_DTMOVTO_AUX') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upINCLUI_DTMOVTO_AUX
GO

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.upINCLUI_DTMOVTO') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.upINCLUI_DTMOVTO
GO

--****************************************************
CREATE PROCEDURE upMOVTOS_TRANSF
   (@Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS
--   SET DATEFIRST 1
   
   DECLARE @ultDtMovto  DATETIME
   
   SELECT @ultDtMovto = MAX(bol_dt_mov)
   FROM tb_ctrl_boletim

   SELECT @Erro = @@ERROR
    
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      End

   SELECT @ultDtMovto = DATEADD(day, -7, @ultDtMovto)
    
   SELECT DISTINCT bol_dt_mov
   FROM tb_ctrl_boletim
   WHERE bol_dt_mov >= @ultDtMovto
   ORDER BY bol_dt_mov
    
   SELECT @Erro = @@ERROR
    
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      End

GO


--****************************************************
CREATE PROCEDURE upMOVTOS_PARA_TRANSF
   (@Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS
--   SET DATEFIRST 1
   
   SELECT DISTINCT bol_dt_mov
   FROM tb_boletim
   WHERE bol_dt_mov NOT IN (SELECT DISTINCT bol_dt_mov
                            FROM tb_ctrl_boletim)
   ORDER BY bol_dt_mov
    
   SELECT @Erro = @@ERROR
    
   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      End

GO

--****************************************************
CREATE PROCEDURE upLIMPA_AUX_BOLETIM
   (@Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS

--   SET DATEFIRST 1

   DELETE FROM tb_aux_boletim
   
   SELECT @Erro = @@ERROR

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

GO

--****************************************************
CREATE PROCEDURE upINCLUI_DTMOVTO_AUX
   (@DataMov datetime,
    @Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS

--   SET DATEFIRST 1

   INSERT INTO tb_aux_boletim
         (bol_dt_mov)
   VALUES (@DataMov)
   
   SELECT @Erro = @@ERROR

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END
GO

--****************************************************
CREATE PROCEDURE upINCLUI_DTMOVTO
   (@Erro    int OUTPUT,
    @MsgErr  varchar(255) OUTPUT)
AS

--   SET DATEFIRST 1

 INSERT INTO tb_ctrl_boletim (bol_dt_mov,
                              emp_cd,
                              cin_cd,
                              sal_cd,
                              fil_cd,
                              dt_envio)

   SELECT bol_dt_mov,
          emp_cd,
          cin_cd,
          sal_cd,
          fil_cd,
          GETDATE()
   FROM tb_boletim
   WHERE bol_dt_mov IN (SELECT bol_dt_mov FROM tb_aux_boletim)
   
   SELECT @Erro = @@ERROR

   IF @Erro <> 0
      BEGIN
         SELECT @MsgErr = description
         FROM master..sysmessages
         WHERE error = @Erro
         
         Return
      END

GO