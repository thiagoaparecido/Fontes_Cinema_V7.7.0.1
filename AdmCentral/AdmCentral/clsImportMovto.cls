VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportMovto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private oConexaoADO As ADODB.Connection

'local variable(s) to hold property value(s)
Private mvarCodigoErro  As Long 'local copy
Private mvarMensagemErro As String 'local copy
Private mvarDirTrab As String 'local copy

Public Property Let DirTrab(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DirTrab = 5
    mvarDirTrab = vData
End Property

Public Property Get DirTrab() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DirTrab
    DirTrab = mvarDirTrab
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

'Public Function ImportMovtoOld(arqInport As String) As Boolean
'    Dim gSQLServer As SQLDMO.SQLServer
'    Dim dbCine     As SQLDMO.Database
'    Dim tbAux      As SQLDMO.Table
'    Dim bcopy      As New SQLDMO.BulkCopy
'
'    On Error GoTo ImportMovto_Error
'
'    ImportMovto = False
'
'    If expandeArquivos(gGetShortPathName(arqInport), gGetShortPathName(mvarDirTrab)) Then
'        Set gSQLServer = New SQLDMO.SQLServer
'        gSQLServer.LoginTimeout = 15
'        gSQLServer.Connect oConexaoADO.Properties("Data Source Name").Value, oConexaoADO.Properties("User ID").Value, oConexaoADO.Properties("Password").Value
'        Set dbCine = gSQLServer.Databases(oConexaoADO.Properties("Current Catalog").Value)
'
'        'Importa Parametros
'        Set tbAux = dbCine.Tables("tb_bol_param")
'
'        bcopy.DataFilePath = mvarDirTrab & "\bckParametros.bak"
'        bcopy.DATAFILETYPE = SQLDMODataFile_Default
'        bcopy.ErrorFilePath = mvarDirTrab & "\bckParametros.err"
'        bcopy.MaximumErrorsBeforeAbort = 60000
'
'        Call tbAux.ImportData(bcopy)
'
'        'Importa Distribuidoras
'        On Error Resume Next
'
'        Set tbAux = dbCine.Tables("tb_bol_distrib")
'
'        bcopy.DataFilePath = mvarDirTrab & "\bckDistrib.bak"
'        bcopy.DATAFILETYPE = SQLDMODataFile_Default
'        bcopy.ErrorFilePath = mvarDirTrab & "\bckDistrib.err"
'        bcopy.MaximumErrorsBeforeAbort = 60000
'
'        Call tbAux.ImportData(bcopy)
'
'        On Error GoTo ImportMovto_Error
'
'        'Importa Filmes
'        Set tbAux = dbCine.Tables("tb_bol_filme")
'
'        bcopy.DataFilePath = mvarDirTrab & "\bckFilmes.bak"
'        bcopy.DATAFILETYPE = SQLDMODataFile_Default
'        bcopy.ErrorFilePath = mvarDirTrab & "\bckFilmes.err"
'        bcopy.MaximumErrorsBeforeAbort = 60000
'
'        Call tbAux.ImportData(bcopy)
'
'        'Importa Cinemas
'        Set tbAux = dbCine.Tables("tb_bol_cin")
'
'        bcopy.DataFilePath = mvarDirTrab & "\bckCinemas.bak"
'        bcopy.DATAFILETYPE = SQLDMODataFile_Default
'        bcopy.ErrorFilePath = mvarDirTrab & "\bckCinemas.err"
'        bcopy.MaximumErrorsBeforeAbort = 60000
'
'        Call tbAux.ImportData(bcopy)
'
'        'Importa Catraca
'        Set tbAux = dbCine.Tables("tb_bol_catraca")
'
'        bcopy.DataFilePath = mvarDirTrab & "\bckCatracas.bak"
'        bcopy.DATAFILETYPE = SQLDMODataFile_Default
'        bcopy.ErrorFilePath = mvarDirTrab & "\bckCatracas.err"
'        bcopy.MaximumErrorsBeforeAbort = 60000
'
'        Call tbAux.ImportData(bcopy)
'
'        'Importa Salas
'        Set tbAux = dbCine.Tables("tb_bol_sala")
'
'        bcopy.DataFilePath = mvarDirTrab & "\bckSalas.bak"
'        bcopy.DATAFILETYPE = SQLDMODataFile_Default
'        bcopy.ErrorFilePath = mvarDirTrab & "\bckSalas.err"
'        bcopy.MaximumErrorsBeforeAbort = 60000
'
'        Call tbAux.ImportData(bcopy)
'
'        'Importa Catracas_Salas
'        Set tbAux = dbCine.Tables("tb_bol_catraca_sala")
'
'        bcopy.DataFilePath = mvarDirTrab & "\bckCatracasSalas.bak"
'        bcopy.DATAFILETYPE = SQLDMODataFile_Default
'        bcopy.ErrorFilePath = mvarDirTrab & "\bckCatracasSalas.err"
'        bcopy.MaximumErrorsBeforeAbort = 60000
'
'        Call tbAux.ImportData(bcopy)
'
'        'Importa Empresa
'        On Error Resume Next
'        Set tbAux = dbCine.Tables("tb_bol_empr")
'
'        bcopy.DataFilePath = mvarDirTrab & "\bckEmpresa.bak"
'        bcopy.DATAFILETYPE = SQLDMODataFile_Default
'        bcopy.ErrorFilePath = mvarDirTrab & "\bckEmpresa.err"
'        bcopy.MaximumErrorsBeforeAbort = 60000
'
'        Call tbAux.ImportData(bcopy)
'
'        On Error GoTo ImportMovto_Error
'
'        'Importa Boletins
'        Set tbAux = dbCine.Tables("tb_boletim")
'
'        bcopy.DataFilePath = mvarDirTrab & "\bckBoletins.bak"
'        bcopy.DATAFILETYPE = SQLDMODataFile_Default
'        bcopy.ErrorFilePath = mvarDirTrab & "\bckBoletins.err"
'        bcopy.MaximumErrorsBeforeAbort = 60000
'
'        Call tbAux.ImportData(bcopy)
'
'        'Importa Sessoes
'        Set tbAux = dbCine.Tables("tb_bol_sessao")
'
'        bcopy.DataFilePath = mvarDirTrab & "\bckSessoes.bak"
'        bcopy.DATAFILETYPE = SQLDMODataFile_Default
'        bcopy.ErrorFilePath = mvarDirTrab & "\bckSessoes.err"
'        bcopy.MaximumErrorsBeforeAbort = 60000
'
'        Call tbAux.ImportData(bcopy)
'
'        'Importa Ingressos
'        Set tbAux = dbCine.Tables("tb_bol_ingre")
'
'        bcopy.DataFilePath = mvarDirTrab & "\bckIngressos.bak"
'        bcopy.DATAFILETYPE = SQLDMODataFile_Default
'        bcopy.ErrorFilePath = mvarDirTrab & "\bckIngressos.err"
'        bcopy.MaximumErrorsBeforeAbort = 60000
'
'        Call tbAux.ImportData(bcopy)
'
'        'Importa talao
'        Set tbAux = dbCine.Tables("tb_bol_talao")
'
'        bcopy.DataFilePath = mvarDirTrab & "\bcktalao.bak"
'        bcopy.DATAFILETYPE = SQLDMODataFile_Default
'        bcopy.ErrorFilePath = mvarDirTrab & "\bcktalao.err"
'        bcopy.MaximumErrorsBeforeAbort = 60000
'
'        Call tbAux.ImportData(bcopy)
'
'        'Importa sisLog
'        Set tbAux = dbCine.Tables("tb_sis_log")
'
'        bcopy.DataFilePath = mvarDirTrab & "\bcksislog.bak"
'        bcopy.DATAFILETYPE = SQLDMODataFile_Default
'        bcopy.ErrorFilePath = mvarDirTrab & "\bcksislog.err"
'        bcopy.MaximumErrorsBeforeAbort = 60000
'
'        Call tbAux.ImportData(bcopy)
'
'        Call gSQLServer.Disconnect
'        Set tbAux = Nothing
'        Set dbCine = Nothing
'        Set bcopy = Nothing
'
'        ImportMovto = True
'    End If
'
'    Kill mvarDirTrab & "\*.bak"
'    Kill mvarDirTrab & "\*.err"
'
'    Exit Function
'
'ImportMovto_Error:
'    If Err.Number = 53 Then
'        Resume Next
'    Else
'        mvarCodigoErro = Err.Number
'        mvarMensagemErro = Err.Description
'    End If
'End Function
'
Public Function ImportMovto(arqInport As String) As Boolean
    
    'On Error GoTo ImportMovto_Error
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim wSql As String
    
    ImportMovto = False
    
    If expandeArquivos(gGetShortPathName(arqInport), gGetShortPathName(mvarDirTrab)) Then
            
        'Importa Parametros
        wSql = "BULK INSERT tb_bol_param From '" + mvarDirTrab + "\bckParametros.bak' " + _
               " WITH  (    DATAFILETYPE = 'char',     FIELDTERMINATOR = ','  ); "
        rs.Open wSql, "file Name=" + Trim(App.Path) + "\AdmCentral.udl", adOpenDynamic, adLockPessimistic

                         
        'Importa Distribuidoras
        wSql = "BULK INSERT tb_bol_distrib From '" + mvarDirTrab + "\bckDistrib.bak' " + _
               " WITH  (    DATAFILETYPE = 'char',     FIELDTERMINATOR = ','  ); "
        rs.Open wSql, "file Name=" + Trim(App.Path) + "\AdmCentral.udl", adOpenDynamic, adLockPessimistic
                         
                         
        'Importa Filmes
        wSql = "BULK INSERT tb_bol_filme From '" + mvarDirTrab + "\bckFilmes.bak' " + _
               " WITH  (    DATAFILETYPE = 'char',     FIELDTERMINATOR = ','  ); "
        rs.Open wSql, "file Name=" + Trim(App.Path) + "\AdmCentral.udl", adOpenDynamic, adLockPessimistic
        
    
        'Importa Cinemas
        wSql = "BULK INSERT tb_bol_cin From '" + mvarDirTrab + "\bckCinemas.bak' " + _
               " WITH  (    DATAFILETYPE = 'char',     FIELDTERMINATOR = ','  ); "
        rs.Open wSql, "file Name=" + Trim(App.Path) + "\AdmCentral.udl", adOpenDynamic, adLockPessimistic
        
        'Importa Catraca
        wSql = "BULK INSERT tb_bol_catraca From '" + mvarDirTrab + "\bckCatracas.bak' " + _
        " WITH  (    DATAFILETYPE = 'char',     FIELDTERMINATOR = ','  ); "
        rs.Open wSql, "file Name=" + Trim(App.Path) + "\AdmCentral.udl", adOpenDynamic, adLockPessimistic

        'Importa Salas
        wSql = "BULK INSERT tb_bol_sala From '" + mvarDirTrab + "\bckSalas.bak' " + _
        " WITH  (    DATAFILETYPE = 'char',     FIELDTERMINATOR = ','  ); "
        rs.Open wSql, "file Name=" + Trim(App.Path) + "\AdmCentral.udl", adOpenDynamic, adLockPessimistic

        'Importa Catracas_Salas
        wSql = "BULK INSERT tb_bol_catraca_sala From '" + mvarDirTrab + "\bckCatracasSalas.bak' " + _
        " WITH  (    DATAFILETYPE = 'char',     FIELDTERMINATOR = ','  ); "
        rs.Open wSql, "file Name=" + Trim(App.Path) + "\AdmCentral.udl", adOpenDynamic, adLockPessimistic

        'Importa Empresa
        wSql = "BULK INSERT tb_bol_empr From '" + mvarDirTrab + "\bckEmpresa.bak' " + _
        " WITH  (    DATAFILETYPE = 'char',     FIELDTERMINATOR = ','  ); "
        rs.Open wSql, "file Name=" + Trim(App.Path) + "\AdmCentral.udl", adOpenDynamic, adLockPessimistic

'        'Importa Boletins
        wSql = "BULK INSERT tb_boletim From '" + mvarDirTrab + "\bckBoletins.bak' " + _
        " WITH  (    DATAFILETYPE = 'char',     FIELDTERMINATOR = ','  ); "
        rs.Open wSql, "file Name=" + Trim(App.Path) + "\AdmCentral.udl", adOpenDynamic, adLockPessimistic

        'Importa Sessoes
        wSql = "BULK INSERT tb_bol_sessao From '" + mvarDirTrab + "\bckSessoes.bak' " + _
        " WITH  (    DATAFILETYPE = 'char',     FIELDTERMINATOR = ','  ); "
        rs.Open wSql, "file Name=" + Trim(App.Path) + "\AdmCentral.udl", adOpenDynamic, adLockPessimistic


        'Importa Ingressos
        wSql = "BULK INSERT tb_bol_ingre From '" + mvarDirTrab + "\bckIngressos.bak' " + _
        " WITH  (    DATAFILETYPE = 'char',     FIELDTERMINATOR = ','  ); "
        rs.Open wSql, "file Name=" + Trim(App.Path) + "\AdmCentral.udl", adOpenDynamic, adLockPessimistic

        'Importa talao
        wSql = "BULK INSERT tb_bol_talao From '" + mvarDirTrab + "\bcktalao.bak' " + _
        " WITH  (    DATAFILETYPE = 'char',     FIELDTERMINATOR = ','  ); "
        rs.Open wSql, "file Name=" + Trim(App.Path) + "\AdmCentral.udl", adOpenDynamic, adLockPessimistic
        
        'Importa sisLog
        wSql = "BULK INSERT tb_sis_log From '" + mvarDirTrab + "\bcksislog.bak' " + _
        " WITH  (    DATAFILETYPE = 'char',     FIELDTERMINATOR = ','  ); "
        rs.Open wSql, "file Name=" + Trim(App.Path) + "\AdmCentral.udl", adOpenDynamic, adLockPessimistic

        If rs.State = 1 Then
            rs.Close
        End If
        Set rs = Nothing
        
        ImportMovto = True
    End If
    
    Kill mvarDirTrab & "\*.bak"
    'Kill mvarDirTrab & "\*.err"
    
    Exit Function

ImportMovto_Error:
    If Err.Number = 53 Then
        Resume Next
    Else
        mvarCodigoErro = Err.Number
        mvarMensagemErro = Err.Description
    End If
End Function


Public Function existeMovto(dtMovto As Date, emp_cd As Integer, cin_cd As Integer) As Boolean
    On Error GoTo existeMovto_Error
    
    Dim Cmd As New ADODB.Command
    Set Cmd.ActiveConnection = oConexaoADO
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "upExisteMovto"
    Cmd.CommandTimeout = 30

    Cmd.Parameters.Append Cmd.CreateParameter("@Data", adDate, adParamInput, , dtMovto)
    Cmd.Parameters.Append Cmd.CreateParameter("@emp_cd", adInteger, adParamInput, , emp_cd)
    Cmd.Parameters.Append Cmd.CreateParameter("@cin_cd", adInteger, adParamInput, , cin_cd)
    Cmd.Parameters.Append Cmd.CreateParameter("@Erro", adInteger, adParamOutput)
    Cmd.Parameters.Append Cmd.CreateParameter("@MsgErr", adVarChar, adParamOutput, 255)
    
    Call Cmd.Execute

    mvarCodigoErro = IIf(IsNull(Cmd.Parameters.Item("@Erro").Value), 0, Cmd.Parameters.Item("@Erro").Value)
    mvarMensagemErro = IIf(IsNull(Cmd.Parameters.Item("@MsgErr").Value), "", Cmd.Parameters.Item("@MsgErr").Value)
    
    existeMovto = (mvarCodigoErro = 0)

    Exit Function

existeMovto_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description
End Function

Public Function excluiMovto(dtMovto As Date, emp_cd As Integer, cin_cd As Integer) As Boolean
    On Error GoTo excluiMovto_Error
    
    Dim Cmd As New ADODB.Command
    Set Cmd.ActiveConnection = oConexaoADO
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "dbo.upExcluiMovto"
    Cmd.CommandTimeout = 30

    Cmd.Parameters.Append Cmd.CreateParameter("@Data", adDate, adParamInput, , dtMovto)
    Cmd.Parameters.Append Cmd.CreateParameter("@emp_cd", adInteger, adParamInput, , emp_cd)
    Cmd.Parameters.Append Cmd.CreateParameter("@cin_cd", adInteger, adParamInput, , cin_cd)
    Cmd.Parameters.Append Cmd.CreateParameter("@Erro", adInteger, adParamOutput)
    Cmd.Parameters.Append Cmd.CreateParameter("@MsgErr", adVarChar, adParamOutput, 255)
    
    Call Cmd.Execute

    mvarCodigoErro = IIf(IsNull(Cmd.Parameters.Item("@Erro").Value), 0, Cmd.Parameters.Item("@Erro").Value)
    mvarMensagemErro = IIf(IsNull(Cmd.Parameters.Item("@MsgErr").Value), "", Cmd.Parameters.Item("@MsgErr").Value)
    
    excluiMovto = (mvarCodigoErro = 0)

    Exit Function

excluiMovto_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description
End Function

