VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFechamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private oConexaoADO As ADODB.Connection

'local variable(s) to hold property value(s)
Private mvarcxp_dt_abertura As Date 'local copy
Private mvarcxa_cd As Integer 'local copy
Private mvarMensagemErro As String 'local copy
Private mvarCodigoErro As Long 'local copy

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

Public Property Set ConexaoADO(ByRef vData As ADODB.Connection)
    Set oConexaoADO = vData
End Property

Public Property Let cxa_cd(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.cxa_cd = 5
    mvarcxa_cd = vData
End Property

Public Property Get cxa_cd() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cxa_cd
    cxa_cd = mvarcxa_cd
End Property
    
Public Property Let cxp_dt_abertura(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.cxp_dt_abertura = 5
    mvarcxp_dt_abertura = vData
End Property

Public Property Get cxp_dt_abertura() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cxp_dt_abertura
    cxp_dt_abertura = mvarcxp_dt_abertura
End Property

Public Function CaixaValores(ByRef oRs As ADODB.Recordset) As Boolean
    CaixaValores = ExecutaProcedure("CAIXA_VALOR", oRs)
End Function

Public Function CaixaBilhetes(ByRef oRs As ADODB.Recordset) As Boolean
    CaixaBilhetes = ExecutaProcedure("CAIXA_BILHETE", oRs)
End Function

Public Function CaixaCombos(ByRef oRs As ADODB.Recordset) As Boolean
    CaixaCombos = ExecutaProcedure("CAIXA_COMBO", oRs)
End Function

Public Function Caixa(ByRef oRs As ADODB.Recordset) As Boolean
    Caixa = ExecutaProcedure("CAIXA", oRs)
End Function

Public Function Caixa1(ByRef oRs As ADODB.Recordset) As Boolean
    Caixa1 = ExecutaProcedure("CAIXA1", oRs)
End Function


Public Function CaixaSelecao(ByRef oRs As ADODB.Recordset) As Boolean
    CaixaSelecao = ExecutaProcedure("CAIXA_SELECAO", oRs)
End Function

Private Function ExecutaProcedure(ByVal sProc As String, Optional ByRef oRs As ADODB.Recordset) As Boolean

    On Error GoTo ExecutaProcedure_Error
    
    Dim Cmd As New ADODB.Command
    Set Cmd.ActiveConnection = oConexaoADO
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "upFECHAMENTO_" & sProc
    Cmd.CommandTimeout = 30

    Select Case sProc
        Case "CAIXA_VALOR", "CAIXA_BILHETE", "CAIXA_COMBO", "CAIXA", "CAIXA1"
            Cmd.Parameters.Append Cmd.CreateParameter("@cxa_cd", adInteger, adParamInput, , Me.cxa_cd)
    End Select
    
    Select Case sProc
        Case "CAIXA_VALOR", "CAIXA_BILHETE", "CAIXA_COMBO", "CAIXA", "CAIXA_SELECAO", "CAIXA1"
            Cmd.Parameters.Append Cmd.CreateParameter("@cxp_dt_abertura", adDate, adParamInput, , Me.cxp_dt_abertura)
    End Select

    Cmd.Parameters.Append Cmd.CreateParameter("@Erro", adInteger, adParamOutput)
    Cmd.Parameters.Append Cmd.CreateParameter("@MsgErr", adVarChar, adParamOutput, 255)
             
    Set oRs = Cmd.Execute

    mvarCodigoErro = IIf(IsNull(Cmd.Parameters.Item("@Erro").Value), 0, Cmd.Parameters.Item("@Erro").Value)
    mvarMensagemErro = IIf(IsNull(Cmd.Parameters.Item("@MsgErr").Value), "", Cmd.Parameters.Item("@MsgErr").Value)
    
    ExecutaProcedure = (mvarCodigoErro = 0)
    Exit Function

ExecutaProcedure_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description

End Function

Private Sub Class_Initialize()
    Me.cxp_dt_abertura = Empty
    Me.cxa_cd = 0
End Sub

