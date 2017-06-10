VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExportBol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Movtos"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private oConexaoADO As ADODB.Connection

'local variable(s) to hold property value(s)
Private mvarCodigoErro  As Long 'local copy
Private mvarMensagemErro As String 'local copy
Private mvarDirExoprt As String 'local copy

Public Property Let DirExoprt(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DirExoprt = 5
    mvarDirExoprt = vData
End Property

Public Property Get DirExoprt() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DirExoprt
    DirExoprt = mvarDirExoprt
End Property

Public Property Get MensagemErro() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.msgSerro
    MensagemErro = mvarMensagemErro
End Property

Public Property Get CodigoErro() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.codErro
    CodigoErro = mvarCodigoErro
End Property

Public Property Set ConexaoADO(ByRef vData As ADODB.Connection)
    Set oConexaoADO = vData
End Property

Public Function TransfMovtos(movtos As clsMovtos) As Boolean
    Dim i As Integer
    'Dim gSQLServer As SQLDMO.SQLServer
    'Dim dbCine     As SQLDMO.Database
    'Dim vwTbAux    As SQLDMO.View
    
    
    'Dim bcopy      As New SQLDMO.BulkCopy


    'On Error GoTo TransfMovtos_Error

    TransfMovtos = False

    If limpaAuxBoletim() Then
        For i = 0 To movtos.Count - 1
            If Not incluiDtMovtoAux(movtos.Item(i + 1).dtMovto) Then
                Exit Function
            End If
        Next i
        
        
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        
        
        'Set gSQLServer = New SQLDMO.SQLServer
        
        'gSQLServer.LoginTimeout = 15
        'gSQLServer.Connect oConexaoADO.Properties("Data Source Name").Value, oConexaoADO.Properties("User ID").Value, oConexaoADO.Properties("Password").Value
        'Set dbCine = gSQLServer.Databases(oConexaoADO.Properties("Current Catalog").Value)

        'Exporta Boletins
        'Set vwTbAux = New SQLDMO.View
        'vwTbAux.Text = "CREATE VIEW vwTbAux as SELECT * FROM tb_boletim WHERE bol_dt_mov IN (SELECT bol_dt_mov FROM tb_aux_boletim)"
        'dbCine.Views.Add vwTbAux
        
        rs.Open "CREATE VIEW vwTbAux as SELECT * FROM tb_boletim WHERE bol_dt_mov IN (SELECT bol_dt_mov FROM tb_aux_boletim)", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic

        'bcopy.DataFilePath = mvarDirExoprt & "\bckBoletins.bak"
        'bcopy.DataFileType = SQLDMODataFile_Default
        'bcopy.ErrorFilePath = mvarDirExoprt & "\bckBoletins.err"
        'bcopy.MaximumErrorsBeforeAbort = 60000

        'Call vwTbAux.ExportData(bcopy)
        'Call dbCine.Views.Remove("vwTbAux")
        
        
        rs.Open "Select *,CONVERT(time,[bol_dt_emissao],103) AS emissao from vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        
        
        Dim wSql As String
        Dim wFile As Double
        Dim X As Integer
        wFile = FreeFile
        Open mvarDirExoprt + "\bckBoletins.bak" For Output As wFile
            Do While Not rs.EOF
                    Print #wFile, Format(rs!bol_dt_mov, "YYYY-MM-DD HH:MM:SS.000") + "," + Trim(str(rs!emp_cd)) + "," + Trim(str(rs!cin_cd)) + "," + Trim(str(rs!sal_cd)) + "," + Trim(str(rs!fil_cd)) + "," + Format(rs!bol_dt_abertura, "YYYY-MM-DD HH:MM:SS.000") + "," + Format(rs!bol_dt_emissao, "YYYY-MM-DD") + " " + Left(rs!emissao, 12) + "," + Format(rs!bol_dt_ini_per, "YYYY-MM-DD HH:MM:SS.000") + "," + Format(rs!bol_dt_fim_per, "YYYY-MM-DD HH:MM:SS.000") + "," + rs!bol_status
                rs.MoveNext
            Loop
        Close #wFile
        
        
        If rs.State = 1 Then rs.Close
        rs.Open "DROP VIEW vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        rs.Open "CREATE VIEW vwTbAux as SELECT * FROM tb_bol_param WHERE bol_dt_mov IN (SELECT bol_dt_mov FROM tb_aux_boletim)", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        rs.Open "Select * from vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        wFile = FreeFile
        Open mvarDirExoprt + "\bckParametros.bak" For Output As wFile
            Do While Not rs.EOF
                    Print #wFile, Format(rs!bol_dt_mov, "YYYY-MM-DD HH:MM:SS.000") + ",";
                    Print #wFile, Trim(str(rs!emp_cd)) + ",";
                    Print #wFile, Trim(str(rs!cin_cd)) + ",";
                    Print #wFile, Format(rs!par_hora_max_ses, "YYYY-MM-DD HH:MM:SS.000") + ",";
                    Print #wFile, Format(rs!par_hora_limite, "YYYY-MM-DD HH:MM:SS.000") + ",";
                    Print #wFile, Replace(FormatNumber(rs!par_custo_ingresso, 4, False), ",", ".") + ",";
                    Print #wFile, Replace(FormatNumber(rs!par_imposto_mun, 4, False), ",", ".") + ",";
                    Print #wFile, Replace(FormatNumber(rs!par_direitos_aut, 4, False), ",", ".") + ",";
                    Print #wFile, Replace(FormatNumber(rs!par_outros, 4, False), ",", ".") + ",";
                    Print #wFile, Format(rs!par_hora_limite23, "YYYY-MM-DD HH:MM:SS.000") + ",";
                    Print #wFile, Format(rs!par_hora_limite34, "YYYY-MM-DD HH:MM:SS.000") + ",";
                    Print #wFile, Format(rs!par_hora_limite45, "YYYY-MM-DD HH:MM:SS.000") + ",";
                    Print #wFile, Format(rs!par_hora_limite56, "YYYY-MM-DD HH:MM:SS.000") + ",";
                    Print #wFile, Format(rs!par_hora_limite12, "YYYY-MM-DD HH:MM:SS.000") + ",";
                    Print #wFile, Replace(FormatNumber(rs!par_perc_meias, 4, False), ",", ".") + ",";
                    Print #wFile, Replace(FormatNumber(rs!par_perc_cortesias, 4, False), ",", ".")
                rs.MoveNext
            Loop
        Close #wFile
        If rs.State = 1 Then rs.Close
        rs.Open "DROP VIEW vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic

        'Exporta Distribuidoras*/
        rs.Open "CREATE VIEW vwTbAux as SELECT * FROM tb_bol_distrib WHERE bol_dt_mov IN (SELECT bol_dt_mov FROM tb_aux_boletim)", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        rs.Open "Select * from vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        wFile = FreeFile
        Open mvarDirExoprt + "\bckDistrib.bak" For Output As wFile
            Do While Not rs.EOF
                Print #wFile, Format(rs!bol_dt_mov, "YYYY-MM-DD HH:MM:SS.000") + ",";
                Print #wFile, Trim(str(rs!dis_cd)) + ",";
                Print #wFile, Trim(rs!dis_nm)
                rs.MoveNext
            Loop
        Close #wFile
        If rs.State = 1 Then rs.Close
        rs.Open "DROP VIEW vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic

        'Exporta Filmes*/
        rs.Open "CREATE VIEW vwTbAux as SELECT * FROM tb_bol_filme WHERE bol_dt_mov IN (SELECT bol_dt_mov FROM tb_aux_boletim)", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        rs.Open "Select * from vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        wFile = FreeFile
        Open mvarDirExoprt + "\bckFilmes.bak" For Output As wFile
            Do While Not rs.EOF
                Print #wFile, Format(rs!bol_dt_mov, "YYYY-MM-DD HH:MM:SS.000") + ",";
                Print #wFile, Trim(str(rs!emp_cd)) + ",";
                Print #wFile, Trim(str(rs!cin_cd)) + ",";
                Print #wFile, Trim(str(rs!fil_cd)) + ",";
                Print #wFile, Trim(str(rs!dis_cd)) + ",";
                Print #wFile, Trim(rs!fil_nm) + ",";
                Print #wFile, Format(rs!fil_dt_ini, "YYYY-MM-DD HH:MM:SS.000") + ",";
                Print #wFile, Format(rs!fil_dt_fim, "YYYY-MM-DD HH:MM:SS.000") + ",";
                Print #wFile, Trim(str(rs!fil_censura)) + ",";
                Print #wFile, Trim(rs!fil_id_nacio)
                rs.MoveNext
            Loop
        Close #wFile
        If rs.State = 1 Then rs.Close
        rs.Open "DROP VIEW vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        
        'Exporta Cinemas*/
        rs.Open "CREATE VIEW vwTbAux as SELECT * FROM tb_bol_cin WHERE bol_dt_mov IN (SELECT bol_dt_mov FROM tb_aux_boletim)", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        rs.Open "Select * from vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        wFile = FreeFile
        Open mvarDirExoprt + "\bckCinemas.bak" For Output As wFile
            Do While Not rs.EOF
                Print #wFile, Format(rs!bol_dt_mov, "YYYY-MM-DD HH:MM:SS.000") + ",";
                Print #wFile, Trim(str(rs!emp_cd)) + ",";
                Print #wFile, Trim(str(rs!cin_cd)) + ",";
                Print #wFile, Trim(rs!cin_nm) + ",";
                Print #wFile, Trim(str(rs!cin_cnpj)) + ",";
                Print #wFile, Trim(rs!cin_end) + ",";
                Print #wFile, Trim(str(rs!cin_num_end)) + ",";
                Print #wFile, Trim(rs!cin_cmp_end) + ",";
                Print #wFile, Trim(rs!cin_brr_end) + ",";
                Print #wFile, Trim(rs!cin_cid_end) + ",";
                Print #wFile, Trim(rs!cin_uf_end) + ",";
                Print #wFile, Trim(rs!cin_cep_end)
                rs.MoveNext
            Loop
        Close #wFile
        If rs.State = 1 Then rs.Close
        rs.Open "DROP VIEW vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
'
        'Exporta Catracas*/
        rs.Open "CREATE VIEW vwTbAux as SELECT * FROM tb_bol_catraca WHERE bol_dt_mov IN (SELECT bol_dt_mov FROM tb_aux_boletim)", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        rs.Open "Select * from vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        wFile = FreeFile
        Open mvarDirExoprt + "\bckCatracas.bak" For Output As wFile
            Do While Not rs.EOF
                Print #wFile, Format(rs!bol_dt_mov, "YYYY-MM-DD HH:MM:SS.000") + ",";
                Print #wFile, Trim(str(rs!emp_cd)) + ",";
                Print #wFile, Trim(str(rs!cin_cd)) + ",";
                Print #wFile, Trim(str(rs!cat_cd)) + ",";
                Print #wFile, Trim(rs!cat_nm) + ",";
                Print #wFile, Trim(str(rs!cat_cd)) + ",";
                Print #wFile, Trim(str(rs!ctc_ini_cont)) + ",";
                Print #wFile, Trim(str(rs!ctc_fim_cont))
                rs.MoveNext
            Loop
        Close #wFile
        If rs.State = 1 Then rs.Close
        rs.Open "DROP VIEW vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic

        'Exporta Salas*/
        rs.Open "CREATE VIEW vwTbAux as SELECT * FROM tb_bol_sala WHERE bol_dt_mov IN (SELECT bol_dt_mov FROM tb_aux_boletim)", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        rs.Open "Select * from vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        wFile = FreeFile
        Open mvarDirExoprt + "\bckSalas.bak" For Output As wFile
            Do While Not rs.EOF
                Print #wFile, Format(rs!bol_dt_mov, "YYYY-MM-DD HH:MM:SS.000") + ",";
                Print #wFile, Trim(str(rs!emp_cd)) + ",";
                Print #wFile, Trim(str(rs!cin_cd)) + ",";
                Print #wFile, Trim(str(rs!sal_cd)) + ",";
                Print #wFile, Trim(rs!sal_desc) + ",";
                Print #wFile, Trim(str(rs!sal_lugares))
                rs.MoveNext
            Loop
        Close #wFile
        If rs.State = 1 Then rs.Close
        rs.Open "DROP VIEW vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic

        'Exporta Catracas_Salas*/
        rs.Open "CREATE VIEW vwTbAux as SELECT * FROM tb_bol_catraca_sala WHERE bol_dt_mov IN (SELECT bol_dt_mov FROM tb_aux_boletim)", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        rs.Open "Select * from vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        wFile = FreeFile
        Open mvarDirExoprt + "\bckCatracasSalas.bak" For Output As wFile
            Do While Not rs.EOF
                Print #wFile, Format(rs!bol_dt_mov, "YYYY-MM-DD HH:MM:SS.000") + ",";
                Print #wFile, Trim(str(rs!emp_cd)) + ",";
                Print #wFile, Trim(str(rs!cin_cd)) + ",";
                Print #wFile, Trim(str(rs!sal_cd)) + ",";
                Print #wFile, Trim(str(rs!cat_cd))
                rs.MoveNext
            Loop
        Close #wFile
        If rs.State = 1 Then rs.Close
        rs.Open "DROP VIEW vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic

        'Exporta Empresa*/
        rs.Open "CREATE VIEW vwTbAux as SELECT * FROM tb_bol_empr WHERE bol_dt_mov IN (SELECT bol_dt_mov FROM tb_aux_boletim)", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        rs.Open "Select * from vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        wFile = FreeFile
        Open mvarDirExoprt + "\bckEmpresa.bak" For Output As wFile
            Do While Not rs.EOF
                Print #wFile, Format(rs!bol_dt_mov, "YYYY-MM-DD HH:MM:SS.000") + ",";
                Print #wFile, Trim(str(rs!emp_cd)) + ",";
                Print #wFile, Trim(rs!emp_nm)
                rs.MoveNext
            Loop
        Close #wFile
        If rs.State = 1 Then rs.Close
        rs.Open "DROP VIEW vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic

        'Exporta Sessoes*/
        rs.Open "CREATE VIEW vwTbAux as SELECT * FROM tb_bol_sessao WHERE bol_dt_mov IN (SELECT bol_dt_mov FROM tb_aux_boletim)", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        rs.Open "Select * from vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        wFile = FreeFile
        Open mvarDirExoprt + "\bckSessoes.bak" For Output As wFile
            Do While Not rs.EOF
                Print #wFile, Format(rs!bol_dt_mov, "YYYY-MM-DD HH:MM:SS.000") + ",";
                Print #wFile, Trim(str(rs!emp_cd)) + ",";
                Print #wFile, Trim(str(rs!cin_cd)) + ",";
                Print #wFile, Trim(str(rs!sal_cd)) + ",";
                Print #wFile, Trim(str(rs!fil_cd)) + ",";
                Print #wFile, Format(rs!sre_data, "YYYY-MM-DD HH:MM:SS.000") + ",";
                Print #wFile, Format(rs!ses_horario, "YYYY-MM-DD HH:MM:SS.000") + ",";
                Print #wFile, Trim(rs!ses_pre_estreia)
                rs.MoveNext
            Loop
        Close #wFile
        If rs.State = 1 Then rs.Close
        rs.Open "DROP VIEW vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        
        'Exporta Ingressos*/
        rs.Open "CREATE VIEW vwTbAux as SELECT * FROM tb_bol_ingre WHERE bol_dt_mov IN (SELECT bol_dt_mov FROM tb_aux_boletim)", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        rs.Open "Select *,convert(time,ing_dt_venda,103) as Venda from vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        wFile = FreeFile
        Open mvarDirExoprt + "\bckIngressos.bak" For Output As wFile
            Do While Not rs.EOF
                Print #wFile, Format(rs!bol_dt_mov, "YYYY-MM-DD HH:MM:SS.000") + ",";
                Print #wFile, Trim(str(rs!emp_cd)) + ",";
                Print #wFile, Trim(str(rs!cin_cd)) + ",";
                Print #wFile, Trim(str(rs!sal_cd)) + ",";
                Print #wFile, Trim(str(rs!fil_cd)) + ",";
                Print #wFile, Format(rs!sre_data, "YYYY-MM-DD HH:MM:SS.000") + ",";
                Print #wFile, Format(rs!sre_horario, "YYYY-MM-DD HH:MM:SS.000") + ",";
                Print #wFile, Trim(str(rs!igt_cd)) + ",";
                Print #wFile, Trim(str(rs!ing_status)) + ",";
                Print #wFile, Format(rs!ing_dt_venda, "YYYY-MM-DD") + " " + Left(rs!Venda, 12) + ",";
                Print #wFile, Trim(str(rs!opt_cd)) + ",";
                Print #wFile, Trim(rs!bin_dev) + ",";
                If rs!cxp_talao = False Then
                    Print #wFile, "0,";
                Else
                    Print #wFile, "1,";
                End If
                
                Print #wFile, Trim(str(rs!ppr_cd)) + ",";
                Print #wFile, Trim(str(rs!pgt_cd)) + ",";
                Print #wFile, Trim(str(rs!bin_qtde)) + ",";
                Print #wFile, Replace(FormatNumber(rs!ing_valor, 4, False), ",", ".")
                rs.MoveNext
            Loop
        Close #wFile
        If rs.State = 1 Then rs.Close
        rs.Open "DROP VIEW vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic

        'Exporta talao*/"
        rs.Open "CREATE VIEW vwTbAux as SELECT * FROM tb_bol_talao WHERE bol_dt_mov IN (SELECT bol_dt_mov FROM tb_aux_boletim)", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        rs.Open "Select * from vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        wFile = FreeFile
        Open mvarDirExoprt + "\bcktalao.bak" For Output As wFile
            Do While Not rs.EOF
                Print #wFile, Format(rs!bol_dt_mov, "YYYY-MM-DD HH:MM:SS.000") + ",";
                Print #wFile, Trim(str(rs!emp_cd)) + ",";
                Print #wFile, Trim(str(rs!cin_cd)) + ",";
                Print #wFile, Trim(str(rs!sal_cd)) + ",";
                Print #wFile, Trim(str(rs!fil_cd)) + ",";
                Print #wFile, Trim(str(rs!igt_cd)) + ",";
                Print #wFile, Trim(str(rs!num_talao_ini)) + ",";
                Print #wFile, Trim(str(rs!num_talao_fim))
                rs.MoveNext
            Loop
        Close #wFile
        If rs.State = 1 Then rs.Close
        rs.Open "DROP VIEW vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic

        'Exporta log*/"
        rs.Open "CREATE VIEW vwTbAux as SELECT * FROM tb_sis_log WHERE convert(datetime, convert(char(10), slg_data, 103), 103) IN (SELECT bol_dt_mov FROM tb_aux_boletim)", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        rs.Open "Select *,Convert(Time,slg_Data,103) as Data from vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic
        
        wFile = FreeFile
        Open mvarDirExoprt + "\bcksislog.bak" For Output As wFile
            Do While Not rs.EOF
                Print #wFile, Trim(str(rs!emp_cd)) + ",";
                Print #wFile, Trim(str(rs!cin_cd)) + ",";
                Print #wFile, Format(rs!slg_data, "YYYY-MM-DD") + " " + Left(rs!data, 12) + ",";
                Print #wFile, Trim(rs!usu_nm) + ",";
                Print #wFile, Trim(rs!slg_descricao)
                rs.MoveNext
            Loop
        Close #wFile
        If rs.State = 1 Then rs.Close
        rs.Open "DROP VIEW vwTbAux", "file Name=" + Trim(App.Path) + "\Cinema.udl", adOpenDynamic, adLockPessimistic

        TransfMovtos = True
    End If

    Exit Function

TransfMovtos_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description
End Function

Public Function movtosTransf() As clsMovtos
    Dim movtos As New clsMovtos
    Dim Cmd    As New ADODB.Command
    Dim oRs    As ADODB.Recordset
    
    On Error GoTo movtosTransf_Error
    
    Set Cmd.ActiveConnection = oConexaoADO
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "upMOVTOS_TRANSF"
    Cmd.CommandTimeout = 30

    Cmd.Parameters.Append Cmd.CreateParameter("@Erro", adInteger, adParamOutput)
    Cmd.Parameters.Append Cmd.CreateParameter("@MsgErr", adVarChar, adParamOutput, 255)
             
    Set oRs = Cmd.Execute

    mvarCodigoErro = IIf(IsNull(Cmd.Parameters.Item("@Erro").Value), 0, Cmd.Parameters.Item("@Erro").Value)
    mvarMensagemErro = IIf(IsNull(Cmd.Parameters.Item("@MsgErr").Value), "", Cmd.Parameters.Item("@MsgErr").Value)
    
    Do While Not oRs.EOF
        
        movtos.Add oRs.Fields("bol_dt_mov").Value
    
        oRs.MoveNext
    Loop

    Set movtosTransf = movtos
    
    Exit Function

movtosTransf_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description
End Function

Public Function movtosPTransf() As clsMovtos
    Dim movtos As New clsMovtos
    Dim Cmd    As New ADODB.Command
    Dim oRs    As ADODB.Recordset
    
    On Error GoTo movtosPTransf_Error
    
    Set Cmd.ActiveConnection = oConexaoADO
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "upMOVTOS_PARA_TRANSF"
    Cmd.CommandTimeout = 30

    Cmd.Parameters.Append Cmd.CreateParameter("@Erro", adInteger, adParamOutput)
    Cmd.Parameters.Append Cmd.CreateParameter("@MsgErr", adVarChar, adParamOutput, 255)
             
    Set oRs = Cmd.Execute

    mvarCodigoErro = IIf(IsNull(Cmd.Parameters.Item("@Erro").Value), 0, Cmd.Parameters.Item("@Erro").Value)
    mvarMensagemErro = IIf(IsNull(Cmd.Parameters.Item("@MsgErr").Value), "", Cmd.Parameters.Item("@MsgErr").Value)
    
    Do While Not oRs.EOF
        
        movtos.Add oRs.Fields("bol_dt_mov").Value
    
        oRs.MoveNext
    Loop
    
    Set movtosPTransf = movtos
    
    Exit Function

movtosPTransf_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description
End Function

Private Function limpaAuxBoletim() As Boolean
    Dim Cmd As New ADODB.Command
    Set Cmd.ActiveConnection = oConexaoADO
    
    On Error GoTo limpaAuxBoletim_Error
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "upLIMPA_AUX_BOLETIM"
    Cmd.CommandTimeout = 30

    Cmd.Parameters.Append Cmd.CreateParameter("@Erro", adInteger, adParamOutput)
    Cmd.Parameters.Append Cmd.CreateParameter("@MsgErr", adVarChar, adParamOutput, 255)
             
    Cmd.Execute

    mvarCodigoErro = IIf(IsNull(Cmd.Parameters.Item("@Erro").Value), 0, Cmd.Parameters.Item("@Erro").Value)
    mvarMensagemErro = IIf(IsNull(Cmd.Parameters.Item("@MsgErr").Value), "", Cmd.Parameters.Item("@MsgErr").Value)
    
    limpaAuxBoletim = (mvarCodigoErro = 0)
    
    Exit Function

limpaAuxBoletim_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description
End Function

Private Function incluiDtMovtoAux(dtMovto As Date) As Boolean
    Dim Cmd As New ADODB.Command
    Set Cmd.ActiveConnection = oConexaoADO
    
    On Error GoTo incluiDtMovtoAux_Error
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "upINCLUI_DTMOVTO_AUX"
    Cmd.CommandTimeout = 30

    Cmd.Parameters.Append Cmd.CreateParameter("@DataMov", adDate, adParamInput, , dtMovto)
    Cmd.Parameters.Append Cmd.CreateParameter("@Erro", adInteger, adParamOutput)
    Cmd.Parameters.Append Cmd.CreateParameter("@MsgErr", adVarChar, adParamOutput, 255)
             
    Cmd.Execute

    mvarCodigoErro = IIf(IsNull(Cmd.Parameters.Item("@Erro").Value), 0, Cmd.Parameters.Item("@Erro").Value)
    mvarMensagemErro = IIf(IsNull(Cmd.Parameters.Item("@MsgErr").Value), "", Cmd.Parameters.Item("@MsgErr").Value)
    
    incluiDtMovtoAux = (mvarCodigoErro = 0)
    
    Exit Function

incluiDtMovtoAux_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description
End Function

Public Function ConfTransfMovtos(movtos As clsMovtos) As Boolean
    Dim i As Integer
    Dim Cmd As New ADODB.Command
    Set Cmd.ActiveConnection = oConexaoADO
    
    On Error GoTo TransfMovtos_Error
    
    ConfTransfMovtos = False
    
    If limpaAuxBoletim() Then
        For i = 0 To movtos.Count - 1
            If Not incluiDtMovtoAux(movtos.Item(i + 1).dtMovto) Then
                Exit Function
            End If
        Next i
        
        Call incluiDtMovto
        
        ConfTransfMovtos = True
    End If
    
    Exit Function

TransfMovtos_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description
End Function

Private Function incluiDtMovto() As Boolean
    Dim Cmd As New ADODB.Command
    Set Cmd.ActiveConnection = oConexaoADO
    
    On Error GoTo incluiDtMovto_Error
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "upINCLUI_DTMOVTO"
    Cmd.CommandTimeout = 30

    Cmd.Parameters.Append Cmd.CreateParameter("@Erro", adInteger, adParamOutput)
    Cmd.Parameters.Append Cmd.CreateParameter("@MsgErr", adVarChar, adParamOutput, 255)
             
    Cmd.Execute

    mvarCodigoErro = IIf(IsNull(Cmd.Parameters.Item("@Erro").Value), 0, Cmd.Parameters.Item("@Erro").Value)
    mvarMensagemErro = IIf(IsNull(Cmd.Parameters.Item("@MsgErr").Value), "", Cmd.Parameters.Item("@MsgErr").Value)
    
    incluiDtMovto = (mvarCodigoErro = 0)
    
    Exit Function

incluiDtMovto_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description
End Function

