VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExpurgo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarCodigoErro As Long             'local copy
Private mvarConexaoADO As ADODB.Connection 'local copy
Private mvarMensagemErro As String         'local copy
Private mvardias As Integer                'local copy

Public Property Let dias(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dias = 5
    mvardias = vData
End Property

Public Property Get dias() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dias
    dias = mvardias
End Property

Public Property Get MensagemErro() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MensagemErro
    MensagemErro = mvarMensagemErro
End Property

Public Property Set ConexaoADO(ByVal vData As ADODB.Connection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ConexaoADO = Form1
    Set mvarConexaoADO = vData
End Property

Public Property Get CodigoErro() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CodigoErro
    CodigoErro = mvarCodigoErro
End Property

Public Function expurgo() As Boolean
    Dim Cmd As New ADODB.Command

    On Error GoTo expurgo_Error
    
    Set Cmd.ActiveConnection = mvarConexaoADO
    
    expurgo = False
    
    If Me.dias < 30 Then
       mvarCodigoErro = 99
       mvarMensagemErro = "Numero de dias deve ser maior que 30"
    End If
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "upExpurgo"
    Cmd.CommandTimeout = 120

    Cmd.Parameters.Append Cmd.CreateParameter("@dias", adInteger, adParamInput, , Me.dias)
    Cmd.Parameters.Append Cmd.CreateParameter("@Erro", adInteger, adParamOutput)
    Cmd.Parameters.Append Cmd.CreateParameter("@MsgErr", adVarChar, adParamOutput, 255)
             
    Cmd.Execute

    mvarCodigoErro = IIf(IsNull(Cmd.Parameters.Item("@Erro").Value), 0, Cmd.Parameters.Item("@Erro").Value)
    mvarMensagemErro = IIf(IsNull(Cmd.Parameters.Item("@MsgErr").Value), "", Cmd.Parameters.Item("@MsgErr").Value)
    
    expurgo = (mvarCodigoErro = 0)
    
    Exit Function

expurgo_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description
End Function

Public Function expurgoCentral() As Boolean
    Dim Cmd As New ADODB.Command

    On Error GoTo expurgoCentral_Error
    
    Set Cmd.ActiveConnection = mvarConexaoADO
    
    expurgoCentral = False
    
    If Me.dias < 30 Then
       mvarCodigoErro = 99
       mvarMensagemErro = "Numero de dias deve ser maior que 30"
    End If
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "upExpurgoCentral"
    Cmd.CommandTimeout = 30

    Cmd.Parameters.Append Cmd.CreateParameter("@dias", adInteger, adParamInput, , Me.dias)
    Cmd.Parameters.Append Cmd.CreateParameter("@Erro", adInteger, adParamOutput)
    Cmd.Parameters.Append Cmd.CreateParameter("@MsgErr", adVarChar, adParamOutput, 255)
             
    Cmd.Execute

    mvarCodigoErro = IIf(IsNull(Cmd.Parameters.Item("@Erro").Value), 0, Cmd.Parameters.Item("@Erro").Value)
    mvarMensagemErro = IIf(IsNull(Cmd.Parameters.Item("@MsgErr").Value), "", Cmd.Parameters.Item("@MsgErr").Value)
    
    expurgoCentral = (mvarCodigoErro = 0)
    
    Exit Function

expurgoCentral_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description
End Function


