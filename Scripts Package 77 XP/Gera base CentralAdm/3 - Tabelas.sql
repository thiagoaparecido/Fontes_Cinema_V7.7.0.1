
 
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
        par_custo_ingresso   NUMERIC(12,3) NULL,
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
 
 
 insert into tb_bol_pag_tp ( pgt_cd, pgt_desc ) values ( 1, 'DINHEIRO' )
 GO
 insert into tb_bol_pag_tp ( pgt_cd, pgt_desc ) values ( 2, 'CARTÃO DE DÉBITO' )
 GO
 insert into tb_bol_pag_tp ( pgt_cd, pgt_desc ) values ( 3, 'CARTÃO DE CRÉDITO' )
 GO
 insert into tb_bol_pag_tp ( pgt_cd, pgt_desc ) values ( 4, 'CHEQUE' )
 GO
 
 insert into tb_bol_ope_tp ( opt_cd, opt_desc, opt_sinal ) values ( 1, 'VENDA', 1 )
 GO
 insert into tb_bol_ope_tp ( opt_cd, opt_desc, opt_sinal ) values ( 2, 'DEPÓSITO FUNDO DE CAIXA', 1 )
 GO
 insert into tb_bol_ope_tp ( opt_cd, opt_desc, opt_sinal ) values ( 3, 'DEVOLUÇÃO DE VALOR INGRESSO', -1 )
 GO
 insert into tb_bol_ope_tp ( opt_cd, opt_desc, opt_sinal ) values ( 4, 'DEVOLUÇÃO DE VALOR COMBO', -1 )
 GO
 insert into tb_bol_ope_tp ( opt_cd, opt_desc, opt_sinal ) values ( 5, 'SAQUE PARA DEPÓSITO', -1 )
 GO
 insert into tb_bol_ope_tp ( opt_cd, opt_desc, opt_sinal ) values ( 6, 'SAQUE FUNDO DE CAIXA', -1 )
 GO
 
 insert into tb_bol_tp_ingr ( igt_cd, igt_desc ) values ( 1, 'INTEIRA' )
 GO
 insert into tb_bol_tp_ingr ( igt_cd, igt_desc ) values ( 2, 'MEIA' )
 GO
 insert into tb_bol_tp_ingr ( igt_cd, igt_desc ) values ( 3, 'PRO. - INTEIRA' )
 GO
 insert into tb_bol_tp_ingr ( igt_cd, igt_desc ) values ( 4, 'PRO. - MEIA' )
 GO
 insert into tb_bol_tp_ingr ( igt_cd, igt_desc ) values ( 9, 'CORTESIA' )
 GO
 
 