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
    
    IF @emp_cd <> 0 AND @cin_cd <> 0
       BEGIN
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
       END
    ELSE
       BEGIN
          SELECT  @qtdeBol = COUNT(*)
          FROM tb_boletim
          WHERE bol_dt_mov = @Data

          SELECT @Erro = @@ERROR
    
          IF @Erro <> 0
             BEGIN
                SELECT @MsgErr = description
                FROM master..sysmessages
                WHERE error = @Erro
               
                RETURN
             END
       END
       
    IF @qtdeBol = 0
       BEGIN
          SELECT @Erro   = 1
          SELECT @MsgErr = 'Movimento não exite'
         
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

    IF @emp_cd <> 0 AND @cin_cd <> 0
       BEGIN
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
          --
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
       END
    ELSE
       BEGIN
          DELETE FROM tb_bol_ingre
          WHERE bol_dt_mov = @Data

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
      
          --SELECT @Erro = @@ERROR
          --
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

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufCNPJ]')  and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufCNPJ]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufDiaSemana]')  and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufDiaSemana]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufNacionalidade]')  and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufNacionalidade]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufSEMANA]')  and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufSEMANA]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufSESSOES]')  and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufSESSOES]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufPREESTREIA]')  and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufPREESTREIA]
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
CREATE FUNCTION ufPREESTREIA (@DataMov    datetime,
                              @emp_cd     int,
                              @cin_cd     int,
                              @sal_cd     int,
                              @fil_cd     int)
RETURNS VARCHAR(255)
AS
   BEGIN	
      DECLARE @ret             VARCHAR(1),
              @ses_pre_estreia VARCHAR(1)
              
      SELECT  @ret = 'N'
      
      DECLARE curSessoes CURSOR
      FOR
         SELECT DISTINCT ses_pre_estreia
         FROM tb_bol_sessao
         WHERE bol_dt_mov = @DataMov
         AND   emp_cd     = @emp_cd
         AND   cin_cd     = @cin_cd
         AND   sal_cd     = @sal_cd
         AND   fil_cd     = @fil_cd
         AND   sre_data   = @DataMov
         
      OPEN curSessoes
      
      FETCH NEXT FROM curSessoes INTO @ses_pre_estreia

      IF (@@FETCH_STATUS <> -1)
         BEGIN
            SELECT  @ret = @ses_pre_estreia
         END
         
      SELECT  @ret = ISNULL(@ret, 'N')   
      
      CLOSE curSessoes
      DEALLOCATE curSessoes

      RETURN @ret
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


DROP VIEW vw_bordero_aux
go

DROP VIEW vw_bor_vlrs
go

DROP VIEW vw_bordero
go

DROP VIEW vw_boletim
go

DROP VIEW vw_bol_empr
go

DROP VIEW vw_bol_cin
go

DROP VIEW vw_bol_sala
go

DROP VIEW vw_bol_filme
go

DROP VIEW vw_bol_distrib
go

 CREATE VIEW vw_boletim as SELECT DISTINCT bol_dt_mov 
                           FROM tb_boletim
 go
 
 CREATE VIEW vw_bol_empr as SELECT DISTINCT null    AS bol_dt_mov,
                                            0       AS emp_cd,
                                            ' Todas'AS emp_nm     
                              UNION (SELECT DISTINCT bol_dt_mov,
                                                     emp_cd,
                                                     emp_nm       
                                     FROM tb_bol_empr)
 go
 
 CREATE VIEW vw_bol_cin as SELECT DISTINCT NULL     AS bol_dt_mov,
                                           0        AS emp_cd,
                                           0        AS cin_cd,
                                           ' Todos' AS cin_nm
                           UNION (SELECT DISTINCT tb_bol_cin.bol_dt_mov,
                                                  tb_boletim.emp_cd,
                                                  tb_bol_cin.cin_cd,
                                                  tb_bol_cin.cin_nm
                                  FROM tb_boletim,
                                       tb_bol_cin
                                  WHERE tb_boletim.bol_dt_mov = tb_bol_cin.bol_dt_mov
                                  AND   tb_boletim.emp_cd     = tb_bol_cin.emp_cd
                                  AND   tb_boletim.cin_cd     = tb_bol_cin.cin_cd)
 go
 
 CREATE VIEW vw_bol_sala as SELECT DISTINCT NULL    AS bol_dt_mov,
                                            0        AS emp_cd,
                                            0        AS cin_cd,
                                            0        AS sal_cd,
                                            ' Todas' AS sal_desc
                           UNION (SELECT DISTINCT tb_bol_sala.bol_dt_mov,
                                                  tb_boletim.emp_cd,
                                                  tb_boletim.cin_cd,
                                                  tb_bol_sala.sal_cd,
                                                  tb_bol_sala.sal_desc
                                  FROM tb_boletim,
                                       tb_bol_sala
                                  WHERE tb_boletim.bol_dt_mov = tb_bol_sala.bol_dt_mov
                                  AND   tb_boletim.emp_cd     = tb_bol_sala.emp_cd
                                  AND   tb_boletim.cin_cd     = tb_bol_sala.cin_cd
                                  AND   tb_boletim.sal_cd     = tb_bol_sala.sal_cd
                                  AND   tb_boletim.bol_status IN ('N', 'S'))
                                  --AND   tb_boletim.bol_status IN ('N'))
 go

 CREATE VIEW vw_bol_distrib as SELECT DISTINCT NULL     AS bol_dt_mov,
                                               0        AS dis_cd,
                                               ' Todas' AS dis_nm
                               UNION (SELECT DISTINCT tb_bol_distrib.bol_dt_mov,
                                                      tb_bol_distrib.dis_cd,
                                                      tb_bol_distrib.dis_nm
                                      FROM tb_bol_distrib)
 go

 CREATE VIEW vw_bol_filme as SELECT DISTINCT tb_bol_filme.bol_dt_mov,
                                             tb_boletim.emp_cd,
                                             tb_boletim.cin_cd,
                                             tb_boletim.sal_cd,
                                             tb_bol_filme.fil_cd,
                                             tb_bol_filme.fil_nm,
                                             tb_bol_filme.dis_cd
                             FROM tb_boletim,
                                  tb_bol_filme
                             WHERE tb_boletim.bol_dt_mov = tb_bol_filme.bol_dt_mov
                             AND   tb_boletim.emp_cd     = tb_bol_filme.emp_cd
                             AND   tb_boletim.cin_cd     = tb_bol_filme.cin_cd
                             AND   tb_boletim.fil_cd     = tb_bol_filme.fil_cd
                             AND   tb_boletim.bol_status IN ('N', 'S')
                             --AND   tb_boletim.bol_status IN ('N')
go

CREATE VIEW vw_bordero AS
SELECT tb_boletim.bol_dt_mov,
       tb_boletim.emp_cd,
       tb_boletim.cin_cd,
       tb_boletim.sal_cd,
       tb_boletim.fil_cd,
       tb_boletim.bol_dt_abertura,
       tb_boletim.bol_dt_emissao,
       tb_boletim.bol_dt_ini_per,
       tb_boletim.bol_dt_fim_per,
       tb_bol_empr.emp_nm,
       tb_bol_cin.cin_nm,
       dbo.ufCNPJ(tb_bol_cin.cin_cnpj) AS cinCnpj,
       tb_bol_cin.cin_cnpj,
       RTRIM(tb_bol_cin.cin_end) + ' ' + CONVERT(VARCHAR,tb_bol_cin.cin_num_end) + ', ' + RTRIM(tb_bol_cin.cin_cmp_end) as cinEnd,
       tb_bol_cin.cin_end,
       tb_bol_cin.cin_cmp_end,
       tb_bol_cin.cin_num_end,
       tb_bol_cin.cin_brr_end,
       tb_bol_cin.cin_cid_end,
       tb_bol_cin.cin_uf_end,
       tb_bol_cin.cin_cep_end,
       tb_bol_sala.sal_desc,
       tb_bol_sala.sal_lugares,
       tb_bol_filme.fil_nm,
       tb_bol_distrib.dis_nm  'fil_distribuidora',
       tb_bol_filme.fil_censura,
       tb_bol_filme.fil_id_nacio,
       tb_bol_filme.fil_dt_ini,
       tb_bol_filme.fil_dt_fim,
       dbo.ufNacionalidade(tb_bol_filme.fil_id_nacio) AS Nacional,
       dbo.ufSESSOES(tb_boletim.bol_dt_mov,
                     tb_boletim.emp_cd,
                     tb_boletim.cin_cd,
                     tb_boletim.sal_cd,
                     tb_boletim.fil_cd)    AS Sessoes,
       dbo.ufSEMANA(tb_boletim.bol_dt_mov,
                    tb_boletim.fil_cd)     AS Semana,
       ISNULL(custos.total_ingr, 0)        AS total_ingr,
       ISNULL(custos.custo_ingresso, 0)    AS custo_ingresso,
       ISNULL(custos.imposto_mun, 0)       AS imposto_mun,
       ISNULL(custos.direitos_aut, 0)      AS direitos_aut,
       ISNULL(custos.outros, 0)            AS outros,
       ISNULL(custos.total_ingr - 
              custos.custo_ingresso - 
              custos.imposto_mun - 
              custos.direitos_aut - 
              custos.outros, 0)            AS vlr_liquido,
       ISNULL(custos.qtdeSUso, 0)          AS qtdeSUso,
       ISNULL(custos.qtdeVend, 0)          AS qtdeVend,
       dbo.ufPREESTREIA(tb_boletim.bol_dt_mov,
                        tb_boletim.emp_cd,
                        tb_boletim.cin_cd,
                        tb_boletim.sal_cd,
                        tb_boletim.fil_cd)    AS PreEstreia,
       ISNULL(custos.qtdeInt, 0)              AS qtdeInt,
       ISNULL(custos.qtdeMeia, 0)             AS qtdeMeia,
       tb_bol_distrib.dis_cd
FROM tb_boletim LEFT OUTER JOIN
     (SELECT tot_ingr.bol_dt_mov,
             tot_ingr.emp_cd,
             tot_ingr.cin_cd,
             tot_ingr.sal_cd,
             tot_ingr.fil_cd,
             ISNULL(tot_ingr.vlr_ingr, 0)                                        AS total_ingr,
             ISNULL(tot_ingr.qtde_ingr * tb_bol_param.par_custo_ingresso, 0)     AS custo_ingresso,
             ISNULL(tot_ingr.vlr_ingr  * (tb_bol_param.par_imposto_mun/100), 0)  AS imposto_mun,
             ISNULL(tot_ingr.vlr_ingr  * (tb_bol_param.par_direitos_aut/100), 0) AS direitos_aut,
             ISNULL(tot_ingr.vlr_ingr  * (tb_bol_param.par_outros/100), 0)       AS outros,
             ISNULL(tot_ingr.qtdeSUso, 0)                                        AS qtdeSUso,
             ISNULL(tot_ingr.qtdeVend, 0)                                        AS qtdeVend,
             ISNULL(tot_ingr.qtde_int, 0)                                        AS qtdeInt,
             ISNULL(tot_ingr.qtde_meia, 0)                                       AS qtdeMeia
      FROM (SELECT tb_bol_ingre.bol_dt_mov,
                   tb_bol_ingre.emp_cd,
                   tb_bol_ingre.cin_cd,
                   tb_bol_ingre.sal_cd,
                   tb_bol_ingre.fil_cd,
                   SUM(tb_bol_ingre.bin_qtde)                          AS qtdeVend,
                   SUM(CASE 
                          WHEN tb_bol_ingre.bin_dev  = 'S' THEN 0
                          ELSE tb_bol_ingre.bin_qtde
                       END)                                            AS qtde_ingr,
                   SUM(CASE 
                          WHEN tb_bol_ingre.bin_dev  = 'S' THEN 0
                          ELSE tb_bol_ingre.ing_valor * tb_bol_ingre.bin_qtde
                       END)                                            AS vlr_ingr,
                   SUM(CASE 
                          WHEN tb_bol_ingre.ing_status = 1 THEN 0
                          ELSE tb_bol_ingre.bin_qtde
                       END)                                            AS qtdeSUso,
                       
                   SUM(CASE 
                          WHEN tb_bol_ingre.igt_cd = 2 
                          OR   tb_bol_ingre.igt_cd = 4
                          OR   tb_bol_ingre.igt_cd = 9
                          OR   tb_bol_ingre.bin_dev = 'S' THEN 0
                          ELSE tb_bol_ingre.bin_qtde
                       END)                                            AS qtde_int,
                   SUM(CASE 
                          WHEN tb_bol_ingre.igt_cd = 1 
                          OR   tb_bol_ingre.igt_cd = 3 
                          OR   tb_bol_ingre.igt_cd = 9
                          OR   tb_bol_ingre.bin_dev = 'S' THEN 0
                          ELSE tb_bol_ingre.bin_qtde
                       END)                                            AS qtde_meia
            FROM tb_bol_ingre
            WHERE tb_bol_ingre.sre_data = tb_bol_ingre.bol_dt_mov
            GROUP BY tb_bol_ingre.bol_dt_mov,
                     tb_bol_ingre.emp_cd,
                     tb_bol_ingre.cin_cd,
                     tb_bol_ingre.sal_cd,
                     tb_bol_ingre.fil_cd) tot_ingr,
            tb_bol_param
      WHERE tot_ingr.bol_dt_mov = tb_bol_param.bol_dt_mov
      AND   tot_ingr.emp_cd     = tb_bol_param.emp_cd
      AND   tot_ingr.cin_cd     = tb_bol_param.cin_cd) custos
     ON  tb_boletim.bol_dt_mov = custos.bol_dt_mov
     AND tb_boletim.emp_cd     = custos.emp_cd
     AND tb_boletim.cin_cd     = custos.cin_cd
     AND tb_boletim.sal_cd     = custos.sal_cd
     AND tb_boletim.fil_cd     = custos.fil_cd,
     tb_bol_empr,
     tb_bol_cin,
     tb_bol_sala,
     tb_bol_filme,
     tb_bol_distrib
WHERE tb_boletim.bol_dt_mov   = tb_bol_empr.bol_dt_mov
AND   tb_boletim.emp_cd       = tb_bol_empr.emp_cd
AND   tb_boletim.bol_dt_mov   = tb_bol_cin.bol_dt_mov
AND   tb_boletim.emp_cd       = tb_bol_cin.emp_cd
AND   tb_boletim.cin_cd       = tb_bol_cin.cin_cd
AND   tb_boletim.bol_dt_mov   = tb_bol_sala.bol_dt_mov
AND   tb_boletim.emp_cd       = tb_bol_sala.emp_cd
AND   tb_boletim.cin_cd       = tb_bol_sala.cin_cd
AND   tb_boletim.sal_cd       = tb_bol_sala.sal_cd
AND   tb_boletim.bol_dt_mov   = tb_bol_filme.bol_dt_mov
AND   tb_boletim.emp_cd       = tb_bol_filme.emp_cd
AND   tb_boletim.cin_cd       = tb_bol_filme.cin_cd
AND   tb_boletim.fil_cd       = tb_bol_filme.fil_cd
AND   tb_bol_filme.bol_dt_mov = tb_bol_distrib.bol_dt_mov
AND   tb_bol_filme.dis_cd     = tb_bol_distrib.dis_cd
AND   tb_boletim.bol_status IN ('N', 'S')
--AND   tb_boletim.bol_status IN ('N')
GO

CREATE VIEW vw_bor_vlrs AS
SELECT tb_bol_ingre.bol_dt_mov,
       tb_bol_ingre.emp_cd,
       tb_bol_ingre.cin_cd,
       tb_bol_ingre.sal_cd,
       tb_bol_ingre.fil_cd,
       tb_bol_ingre.igt_cd,
       tb_bol_tp_ingr.igt_desc                                      AS tipo_Ingr,
       tb_bol_ingre.ing_valor                                       AS vlr_ingr,
       SUM(tb_bol_ingre.bin_qtde)                                   AS qtde_ven,
       SUM(Case tb_bol_ingre.bin_dev
              WHEN 'S' THEN tb_bol_ingre.bin_qtde
              ELSE 0
           END)                                                      AS qtde_dev,
       SUM(tb_bol_ingre.bin_qtde - CASE tb_bol_ingre.bin_dev
                                      WHEN 'S' THEN tb_bol_ingre.bin_qtde
                                      ELSE 0
                                   END)                              AS qtde_utl,
       SUM((tb_bol_ingre.bin_qtde - CASE tb_bol_ingre.bin_dev
                                       WHEN 'S' THEN tb_bol_ingre.bin_qtde
                                       ELSE 0
                                    END) * tb_bol_ingre.ing_valor)   AS vlr_tot
FROM  tb_bol_ingre,
      tb_bol_tp_ingr
WHERE tb_bol_ingre.igt_cd     = tb_bol_tp_ingr.igt_cd
AND   tb_bol_ingre.sre_data   = tb_bol_ingre.bol_dt_mov
GROUP BY tb_bol_ingre.bol_dt_mov,
         tb_bol_ingre.emp_cd,
         tb_bol_ingre.cin_cd,
         tb_bol_ingre.sal_cd,
         tb_bol_ingre.fil_cd,
         tb_bol_ingre.igt_cd,
         tb_bol_tp_ingr.igt_desc,
         tb_bol_ingre.ing_valor
GO

         

CREATE VIEW vw_bordero_aux AS
SELECT convert(varchar(10),vw_bordero.bol_dt_mov,112)+
       convert(varchar(3),vw_bordero.emp_cd)+
       convert(varchar(3),vw_bordero.cin_cd)+
       convert(varchar(3),vw_bordero.sal_cd)+
       convert(varchar(9),vw_bordero.fil_cd) AS grupo,
       vw_bordero.bol_dt_mov,
       vw_bordero.emp_nm,
       vw_bordero.cin_nm,
       vw_bordero.cinCnpj,
       vw_bordero.cinEnd,
       vw_bordero.cin_cid_end,
       vw_bordero.sal_desc,
       vw_bordero.sal_lugares,
       vw_bordero.fil_nm,
       vw_bordero.fil_distribuidora,
       vw_bordero.Nacional,
       CASE
          WHEN vw_bordero.Sessoes IS NOT NULL THEN vw_bordero.Sessoes
          ELSE 'EXCLUIDO'
       END AS Sessoes,
       vw_bordero.Semana,
       vw_bor_vlrs.igt_cd,
       vw_bor_vlrs.tipo_Ingr                AS tipo_Ingr,
       vw_bor_vlrs.vlr_ingr                 AS vlr_ingr,
       ISNULL(vw_bor_vlrs.qtde_ven, 0)      AS qtde_ven,
       ISNULL(vw_bor_vlrs.qtde_dev, 0)      AS qtde_dev,
       ISNULL(vw_bor_vlrs.qtde_utl, 0)      AS qtde_utl,
       ISNULL(vw_bor_vlrs.vlr_tot, 0)       AS vlr_tot,
       ISNULL(vw_bordero.custo_ingresso, 0) AS custo_ingresso,
       ISNULL(vw_bordero.imposto_mun, 0)    AS imposto_mun,
       ISNULL(vw_bordero.direitos_aut, 0)   AS direitos_aut,
       ISNULL(vw_bordero.outros, 0)         AS outros,
       ISNULL(vw_bordero.custo_ingresso+
              vw_bordero.imposto_mun+
              vw_bordero.direitos_aut+
              vw_bordero.outros, 0)         AS vlr_custo,
       ISNULL(vw_bordero.vlr_liquido, 0)    AS vlr_liquido,
       ISNULL(vw_bordero.qtdeSUso, 0)       AS qtdeSUso,
       ISNULL(vw_bordero.qtdeVend, 0)       AS qtdeVend,
       vw_bordero.emp_cd,
       vw_bordero.cin_cd,
       vw_bordero.sal_cd,
       vw_bordero.fil_cd,
       CONVERT(char(10), CASE 
                           WHEN vw_bordero.PreEstreia = 'N' THEN vw_bordero.bol_dt_mov
                           ELSE DATEADD(d, 1, vw_bordero.bol_dt_mov)
                         END, 103) AS boldtmov1,
       ISNULL(vw_bordero.qtdeInt, 0)        AS qtdeInt,
       ISNULL(vw_bordero.qtdeMeia, 0)       AS qtdeMeia,
       vw_bordero.dis_cd
FROM vw_bordero LEFT OUTER JOIN vw_bor_vlrs
ON  vw_bordero.bol_dt_mov = vw_bor_vlrs.bol_dt_mov
AND vw_bordero.emp_cd     = vw_bor_vlrs.emp_cd    
AND vw_bordero.cin_cd     = vw_bor_vlrs.cin_cd    
AND vw_bordero.sal_cd     = vw_bor_vlrs.sal_cd    
AND vw_bordero.fil_cd     = vw_bor_vlrs.fil_cd    
GO

DROP VIEW vw_rel_tot_publico
go

DROP VIEW vw_ocupacao_aux1
go

DROP VIEW vw_ocupacao_aux2
go

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ufHORARIOS]')  and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ufHORARIOS]
GO

--*********************************************************************************
--                               TOTAL PUBLICO
--*********************************************************************************
CREATE VIEW vw_rel_tot_publico AS
SELECT tb_bol_empr.emp_cd          AS emp_cd,
       tb_bol_empr.emp_nm          AS emp_nm,
       tb_bol_cin.cin_cd           AS cin_cd,
       tb_bol_cin.cin_nm           AS cin_nm,
       tb_bol_ingre.sre_data       AS sre_data,
       CONVERT(VARCHAR(10),tb_bol_ingre.sre_data, 103)  AS sre_data_str,
       SUM(tb_bol_ingre.bin_qtde)  AS publico,
       SUM(tb_bol_ingre.ing_valor * tb_bol_ingre.bin_qtde) AS renda,
       SUM(CASE
              WHEN tb_bol_ingre.igt_cd = 9 THEN tb_bol_ingre.bin_qtde
              ELSE 0
           END)                    AS cortesias,
       SUM(CASE
              WHEN tb_bol_ingre.igt_cd = 1 
              OR   tb_bol_ingre.igt_cd = 3 THEN tb_bol_ingre.bin_qtde
              ELSE 0
           END)                    AS interias,
       SUM(CASE
              WHEN tb_bol_ingre.igt_cd = 2 
              OR   tb_bol_ingre.igt_cd = 4 THEN tb_bol_ingre.bin_qtde
              ELSE 0
           END)                    AS meias,
       CASE 
          WHEN DATEPART (weekday, tb_bol_ingre.sre_data) = 1 THEN 'DOM'
          WHEN DATEPART (weekday, tb_bol_ingre.sre_data) = 2 THEN 'SEG'
          WHEN DATEPART (weekday, tb_bol_ingre.sre_data) = 3 THEN 'TER'
          WHEN DATEPART (weekday, tb_bol_ingre.sre_data) = 4 THEN 'QUA'
          WHEN DATEPART (weekday, tb_bol_ingre.sre_data) = 5 THEN 'QUI'
          WHEN DATEPART (weekday, tb_bol_ingre.sre_data) = 6 THEN 'SEX'
          WHEN DATEPART (weekday, tb_bol_ingre.sre_data) = 7 THEN 'SAB'
       END                        AS diaSemana
FROM tb_boletim,
     tb_bol_empr, 
     tb_bol_cin,
     tb_bol_ingre
WHERE tb_boletim.bol_dt_mov = tb_bol_empr.bol_dt_mov
AND   tb_boletim.emp_cd     = tb_bol_empr.emp_cd
AND   tb_boletim.bol_dt_mov = tb_bol_cin.bol_dt_mov
AND   tb_boletim.emp_cd     = tb_bol_cin.emp_cd
AND   tb_boletim.cin_cd     = tb_bol_cin.cin_cd
AND   tb_boletim.bol_dt_mov = tb_bol_ingre.bol_dt_mov
AND   tb_boletim.emp_cd     = tb_bol_ingre.emp_cd
AND   tb_boletim.cin_cd     = tb_bol_ingre.cin_cd
AND   tb_boletim.sal_cd     = tb_bol_ingre.sal_cd
AND   tb_boletim.fil_cd     = tb_bol_ingre.fil_cd
AND   tb_bol_ingre.sre_data = tb_bol_ingre.bol_dt_mov ---******
AND   tb_bol_ingre.bin_dev  = 'N'
GROUP BY tb_bol_empr.emp_cd,
         tb_bol_empr.emp_nm,
         tb_bol_cin.cin_cd,
         tb_bol_cin.cin_nm ,
         tb_bol_ingre.sre_data,
         CASE 
            WHEN DATEPART (weekday, tb_bol_ingre.sre_data) = 1 THEN 'DOM'
            WHEN DATEPART (weekday, tb_bol_ingre.sre_data) = 2 THEN 'SEG'
            WHEN DATEPART (weekday, tb_bol_ingre.sre_data) = 3 THEN 'TER'
            WHEN DATEPART (weekday, tb_bol_ingre.sre_data) = 4 THEN 'QUA'
            WHEN DATEPART (weekday, tb_bol_ingre.sre_data) = 5 THEN 'QUI'
            WHEN DATEPART (weekday, tb_bol_ingre.sre_data) = 6 THEN 'SEX'
            WHEN DATEPART (weekday, tb_bol_ingre.sre_data) = 7 THEN 'SAB'
         END
GO

--*********************************************************************************
--                               OCUPACAO
--*********************************************************************************
CREATE VIEW vw_ocupacao_aux1 AS
SELECT tb_bol_empr.emp_cd          AS emp_cd,
       tb_bol_empr.emp_nm          AS emp_nm,
       tb_bol_cin.cin_cd           AS cin_cd,
       tb_bol_cin.cin_nm           AS cin_nm,
       tb_bol_sala.sal_cd          AS sal_cd,
       tb_bol_sala.sal_desc        AS sal_desc,
       tb_bol_sala.sal_lugares     AS sal_lugares,
       tb_bol_filme.fil_cd         AS fil_cd,
       tb_bol_filme.fil_nm         AS fil_nm,
       tb_bol_ingre.sre_data       AS sre_data,
       tb_bol_ingre.sre_horario    AS sre_horario,
       SUM(tb_bol_ingre.bin_qtde)  AS publico
FROM tb_boletim,
     tb_bol_empr, 
     tb_bol_cin,
     tb_bol_sala,
     tb_bol_filme,
     tb_bol_ingre
WHERE tb_boletim.bol_dt_mov   = tb_bol_empr.bol_dt_mov
AND   tb_boletim.emp_cd       = tb_bol_empr.emp_cd
AND   tb_boletim.bol_dt_mov   = tb_bol_cin.bol_dt_mov
AND   tb_boletim.emp_cd       = tb_bol_cin.emp_cd
AND   tb_boletim.cin_cd       = tb_bol_cin.cin_cd
AND   tb_boletim.bol_dt_mov   = tb_bol_sala.bol_dt_mov
AND   tb_boletim.emp_cd       = tb_bol_sala.emp_cd
AND   tb_boletim.cin_cd       = tb_bol_sala.cin_cd
AND   tb_boletim.sal_cd       = tb_bol_sala.sal_cd
AND   tb_boletim.bol_dt_mov   = tb_bol_filme.bol_dt_mov
AND   tb_boletim.emp_cd       = tb_bol_filme.emp_cd
AND   tb_boletim.cin_cd       = tb_bol_filme.cin_cd
AND   tb_boletim.fil_cd       = tb_bol_filme.fil_cd
AND   tb_boletim.bol_dt_mov   = tb_bol_ingre.bol_dt_mov
AND   tb_boletim.emp_cd       = tb_bol_ingre.emp_cd
AND   tb_boletim.cin_cd       = tb_bol_ingre.cin_cd
AND   tb_boletim.sal_cd       = tb_bol_ingre.sal_cd
AND   tb_boletim.fil_cd       = tb_bol_ingre.fil_cd
AND   tb_bol_ingre.sre_data   = tb_bol_ingre.bol_dt_mov ---******
AND   tb_bol_ingre.bin_dev    = 'N'
GROUP BY tb_bol_empr.emp_cd,
         tb_bol_empr.emp_nm,
         tb_bol_cin.cin_cd,
         tb_bol_cin.cin_nm,
         tb_bol_sala.sal_cd,
         tb_bol_sala.sal_desc,
         tb_bol_sala.sal_lugares,
         tb_bol_filme.fil_cd,
         tb_bol_filme.fil_nm,
         tb_bol_ingre.sre_data,
         tb_bol_ingre.sre_horario
         
GO

CREATE FUNCTION ufHORARIOS(@emp_cd     int,
                           @cin_cd     int,
                           @sal_cd     int,
                           @fil_cd     int,
                           @sre_data   datetime,
                           @publico    int)
RETURNS VARCHAR(255)
AS
   BEGIN	
      DECLARE @ret         VARCHAR(255),
              @sre_horario datetime
      
      DECLARE curSessoes CURSOR
      FOR
         SELECT sre_horario
         FROM vw_ocupacao_aux1
         WHERE emp_cd     = @emp_cd
         AND   cin_cd     = @cin_cd
         AND   sal_cd     = @sal_cd
         AND   fil_cd     = @fil_cd
         AND   sre_data   = @sre_data
         AND   publico    = @publico
         

      OPEN curSessoes
      
      FETCH NEXT FROM curSessoes INTO @sre_horario

      IF (@@FETCH_STATUS <> -1)
         BEGIN
            IF (@@FETCH_STATUS <> -2)
               SELECT @ret =RTRIM(convert(char(5), @sre_horario, 108))
       
            FETCH NEXT FROM curSessoes INTO @sre_horario
            WHILE (@@FETCH_STATUS <> -1)
               BEGIN
                  IF (@@FETCH_STATUS <> -2)
                     SELECT @ret = @ret + '/' + RTRIM(convert(char(5), @sre_horario, 108))
           
              FETCH NEXT FROM curSessoes INTO @sre_horario
            END
         END
      
      CLOSE curSessoes
      DEALLOCATE curSessoes

      RETURN @ret
   END	

GO

CREATE VIEW vw_ocupacao_aux2 AS
SELECT emp_cd,
       emp_nm,
       cin_cd,
       cin_nm,
       sal_cd,
       sal_desc,
       sal_lugares,
       fil_cd,
       fil_nm,
       sre_data,
       CONVERT(VARCHAR(10),sre_data, 103)  AS sre_data_str,
       COUNT(1)                     AS TotSessoes,
       SUM(publico)                 AS TotPublico,
       MAX(publico)                 AS MaxPublico,
       dbo.ufHORARIOS(emp_cd,
                      cin_cd,
                      sal_cd,
                      fil_cd,
                      sre_data,
                      MAX(publico)) AS MaxSessao,
       MIN(publico)                 AS MinPublico,
       dbo.ufHORARIOS(emp_cd,
                      cin_cd,
                      sal_cd,
                      fil_cd,
                      sre_data,
                      MIN(publico)) AS MinSessao
FROM vw_ocupacao_aux1
GROUP BY emp_cd,
         emp_nm,
         cin_cd,
         cin_nm,
         sal_cd,
         sal_desc,
         sal_lugares,
         fil_cd,
         fil_nm,
         sre_data,
         CONVERT(VARCHAR(10),sre_data, 103)
GO       
  