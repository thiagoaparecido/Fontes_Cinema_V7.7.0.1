VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private oConexaoADO As ADODB.Connection

'local variable(s) to hold property value(s)
Private mvarusu_nm        As String  'local copy
Private mvartpl_cd        As Integer 'local copy
Private mvarslg_descricao As String  'local copy
Private mvardt_ini        As Date    'local copy
Private mvardt_fim        As Date    'local copy
Private mvaremp_cd        As Integer 'local copy
Private mvarcin_cd        As Integer 'local copy
Private mvarMensagemErro  As String  'local copy
Private mvarCodigoErro    As Long    'local copy

Public Function consLog(ByRef oRs As ADODB.Recordset) As Boolean
    Dim Cmd As New ADODB.Command

    On Error GoTo consLog_Error
    
    Set Cmd.ActiveConnection = oConexaoADO
    
    consLog = False
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "upTB_LOG_S"
    Cmd.CommandTimeout = 30

    Cmd.Parameters.Append Cmd.CreateParameter("@dt_ini datetime", adDate, adParamInput, , Me.dt_ini)
    Cmd.Parameters.Append Cmd.CreateParameter("@dt_fim datetime", adDate, adParamInput, , Me.dt_fim)
    Cmd.Parameters.Append Cmd.CreateParameter("@emp_cd int", adInteger, adParamInput, , Me.emp_cd)
    Cmd.Parameters.Append Cmd.CreateParameter("@cin_cd int", adInteger, adParamInput, , Me.cin_cd)
             
    Set oRs = Cmd.Execute
    
    consLog = True
    
    Exit Function

consLog_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description
End Function

Public Function consEmpresa(ByRef oRs As ADODB.Recordset) As Boolean
    Dim Cmd As New ADODB.Command

    On Error GoTo consEmpresa_Error
    
    Set Cmd.ActiveConnection = oConexaoADO
    
    consEmpresa = False
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "upEmpresa_S"
    Cmd.CommandTimeout = 30

    Set oRs = Cmd.Execute
    
    consEmpresa = True
    
    Exit Function

consEmpresa_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description
End Function

Public Function consCinema(ByRef oRs As ADODB.Recordset) As Boolean
    Dim Cmd As New ADODB.Command

    On Error GoTo consCinema_Error
    
    Set Cmd.ActiveConnection = oConexaoADO
    
    consCinema = False
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "upCinema_S"
    Cmd.CommandTimeout = 30

    Cmd.Parameters.Append Cmd.CreateParameter("@emp_cd int", adInteger, adParamInput, , Me.emp_cd)
             
    Set oRs = Cmd.Execute
    
    consCinema = True
    
    Exit Function

consCinema_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description
End Function

Public Function insereLog() As Boolean
    Dim Cmd As New ADODB.Command

    On Error GoTo insereLog_Error
    
    Set Cmd.ActiveConnection = oConexaoADO
    
    insereLog = False
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "upTB_SIS_LOG_I"
    Cmd.CommandTimeout = 30

    Cmd.Parameters.Append Cmd.CreateParameter("@usu_nm", adVarChar, adParamInput, 50, Me.usu_nm)
    Cmd.Parameters.Append Cmd.CreateParameter("@slg_descricao", adVarChar, adParamInput, 4000, Me.slg_descricao)
    Cmd.Parameters.Append Cmd.CreateParameter("@Erro", adInteger, adParamOutput)
    Cmd.Parameters.Append Cmd.CreateParameter("@MsgErr", adVarChar, adParamOutput, 255)
             
    Cmd.Execute

    mvarCodigoErro = IIf(IsNull(Cmd.Parameters.Item("@Erro").Value), 0, Cmd.Parameters.Item("@Erro").Value)
    mvarMensagemErro = IIf(IsNull(Cmd.Parameters.Item("@MsgErr").Value), "", Cmd.Parameters.Item("@MsgErr").Value)
    
    insereLog = (mvarCodigoErro = 0)
    
    Exit Function

insereLog_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description
End Function

Public Property Set ConexaoADO(ByRef vData As ADODB.Connection)
    Set oConexaoADO = vData
End Property

Public Property Get CodigoErro() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CodigoErro
    CodigoErro = mvarCodigoErro
End Property

Public Property Get MensagemErro() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MensagemErro
    MensagemErro = mvarMensagemErro
End Property

Public Property Let dt_fim(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dt_fim = 5
    mvardt_fim = vData
End Property

Public Property Get dt_fim() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dt_fim
    dt_fim = mvardt_fim
End Property

Public Property Let dt_ini(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dt_ini = 5
    mvardt_ini = vData
End Property

Public Property Get dt_ini() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dt_ini
    dt_ini = mvardt_ini
End Property

Public Property Let slg_descricao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.slg_descricao = 5
    mvarslg_descricao = vData
End Property

Public Property Get slg_descricao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.slg_descricao
    slg_descricao = mvarslg_descricao
End Property

Public Property Let usu_nm(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.usu_nm = 5
    mvarusu_nm = vData
End Property

Public Property Get usu_nm() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.usu_nm
    usu_nm = mvarusu_nm
End Property

Public Property Let cin_cd(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.cin_cd = 5
    mvarcin_cd = vData
End Property

Public Property Get cin_cd() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cin_cd
    cin_cd = mvarcin_cd
End Property

Public Property Let emp_cd(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.emp_cd = 5
    mvaremp_cd = vData
End Property

Public Property Get emp_cd() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.emp_cd
    emp_cd = mvaremp_cd
End Property


