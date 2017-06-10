VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBoletim2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsCaixa"
Option Explicit

Private oConexaoADO As ADODB.Connection

'local variable(s) to hold property value(s)
Private mvarsal_cd As Integer 'local copy
Private mvarfil_cd As Long 'local copy
Private mvardata As Date 'local copy
Private mvarMensagemErro As String 'local copy
Private mvarCodigoErro As Long 'local copy
Private mvarclsCaixa As clsCaixa

Public Property Get clsCaixa() As clsCaixa
    If mvarclsCaixa Is Nothing Then
        Set mvarclsCaixa = New clsCaixa
    End If


    Set clsCaixa = mvarclsCaixa
End Property

Public Property Set clsCaixa(vData As clsCaixa)
    Set mvarclsCaixa = vData
End Property

Private Sub Class_Terminate()
    Set mvarclsCaixa = Nothing
End Sub

Public Function numeracaoTalao(ByRef oRs As ADODB.Recordset) As Boolean
    numeracaoTalao = ExecutaProcedure("NUMERACAO_TALAO2", oRs)
End Function

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

Public Property Let data(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.data = 5
    mvardata = vData
End Property

Public Property Get data() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.data
    data = mvardata
End Property

Public Property Let fil_cd(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fil_cd = 5
    mvarfil_cd = vData
End Property

Public Property Get fil_cd() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fil_cd
    fil_cd = mvarfil_cd
End Property

Public Property Let sal_cd(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sal_cd = 5
    mvarsal_cd = vData
End Property

Public Property Get sal_cd() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sal_cd
    sal_cd = mvarsal_cd
End Property

Public Function Capa(ByRef oRs As ADODB.Recordset) As Boolean
    Capa = ExecutaProcedure("CAPA2", oRs)
End Function

Public Function SessoesFilme(ByRef oRs As ADODB.Recordset) As Boolean
    SessoesFilme = ExecutaProcedure("SESSOES_FILME2", oRs)
End Function

Public Function VendasDia(ByRef oRs As ADODB.Recordset) As Boolean
    VendasDia = ExecutaProcedure("VENDAS_DIA2", oRs)
End Function

Public Function VendasTotal(ByRef oRs As ADODB.Recordset) As Boolean
    VendasTotal = ExecutaProcedure("VENDAS_TOTAL2", oRs)
End Function

Public Function PreVenda(ByRef oRs As ADODB.Recordset) As Boolean
    PreVenda = ExecutaProcedure("PRE_VENDA2", oRs)
End Function

Public Function Cortesia(ByRef oRs As ADODB.Recordset) As Boolean
    Cortesia = ExecutaProcedure("CORTESIA2", oRs)
End Function

Public Function Devolucao(ByRef oRs As ADODB.Recordset) As Boolean
    Devolucao = ExecutaProcedure("DEVOLUCAO2", oRs)
End Function

Public Function TotalSessao(ByRef oRs As ADODB.Recordset) As Boolean
    TotalSessao = ExecutaProcedure("TOTAL_SESSAO2", oRs)
End Function

Public Function VendaAntecipada(ByRef oRs As ADODB.Recordset) As Boolean
    VendaAntecipada = ExecutaProcedure("VENDA_ANTECIPADA2", oRs)
End Function

Public Function FormaPagto(ByRef oRs As ADODB.Recordset) As Boolean
    FormaPagto = ExecutaProcedure("FORMA_PAGTO2", oRs)
End Function

Public Function Catraca(ByRef oRs As ADODB.Recordset) As Boolean
    Catraca = ExecutaProcedure("CATRACA2", oRs)
End Function

Public Function ingressosSemUso(ByRef oRs As ADODB.Recordset) As Boolean
    ingressosSemUso = ExecutaProcedure("INGRESSO_S_USO2", oRs)
End Function

Public Function FilmesBol(ByRef oRs As ADODB.Recordset) As Boolean

    On Error GoTo FilmesBol_Error
    
    Dim Cmd As New ADODB.Command
    Set Cmd.ActiveConnection = oConexaoADO
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "upTB_BOL_FILME_CARTAZ"
    Cmd.CommandTimeout = 30

    Cmd.Parameters.Append Cmd.CreateParameter("@data", adDate, adParamInput, , Me.data)
    Cmd.Parameters.Append Cmd.CreateParameter("@Erro", adInteger, adParamOutput)
    Cmd.Parameters.Append Cmd.CreateParameter("@MsgErr", adVarChar, adParamOutput, 255)
             
    Set oRs = Cmd.Execute

    mvarCodigoErro = IIf(IsNull(Cmd.Parameters.Item("@Erro").Value), 0, Cmd.Parameters.Item("@Erro").Value)
    mvarMensagemErro = IIf(IsNull(Cmd.Parameters.Item("@MsgErr").Value), "", Cmd.Parameters.Item("@MsgErr").Value)
    
    FilmesBol = (mvarCodigoErro = 0)
    
    Exit Function

FilmesBol_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description
End Function

Public Function VerificaBol() As Boolean

    On Error GoTo FilmesBol_Error
    
    Dim Cmd As New ADODB.Command
    Set Cmd.ActiveConnection = oConexaoADO
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "upTB_BOLETIM_VERIFICA"
    Cmd.CommandTimeout = 30

    Cmd.Parameters.Append Cmd.CreateParameter("@DataMov", adDate, adParamInput, , Me.data)
    Cmd.Parameters.Append Cmd.CreateParameter("@Erro", adInteger, adParamOutput)
    Cmd.Parameters.Append Cmd.CreateParameter("@MsgErr", adVarChar, adParamOutput, 255)
             
    Cmd.Execute

    mvarCodigoErro = IIf(IsNull(Cmd.Parameters.Item("@Erro").Value), 0, Cmd.Parameters.Item("@Erro").Value)
    mvarMensagemErro = IIf(IsNull(Cmd.Parameters.Item("@MsgErr").Value), "", Cmd.Parameters.Item("@MsgErr").Value)
    
    VerificaBol = (mvarCodigoErro = 0)
    
    Exit Function

FilmesBol_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description
End Function

Public Function GeraBol() As Boolean

    On Error GoTo GeraBol_Error
    
    Dim Cmd As New ADODB.Command
    Set Cmd.ActiveConnection = oConexaoADO
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "upBOLETIM2_GERA"
    Cmd.CommandTimeout = 30

    Cmd.Parameters.Append Cmd.CreateParameter("@data", adDate, adParamInput, , Me.data)
    Cmd.Parameters.Append Cmd.CreateParameter("@Erro", adInteger, adParamOutput)
    Cmd.Parameters.Append Cmd.CreateParameter("@MsgErr", adVarChar, adParamOutput, 255)
             
    Cmd.Execute

    mvarCodigoErro = IIf(IsNull(Cmd.Parameters.Item("@Erro").Value), 0, Cmd.Parameters.Item("@Erro").Value)
    mvarMensagemErro = IIf(IsNull(Cmd.Parameters.Item("@MsgErr").Value), "", Cmd.Parameters.Item("@MsgErr").Value)
    
    GeraBol = (mvarCodigoErro = 0)
    
    Exit Function

GeraBol_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description
End Function

Private Function ExecutaProcedure(ByVal sProc As String, Optional ByRef oRs As ADODB.Recordset) As Boolean

    On Error GoTo ExecutaProcedure_Error
    
    Dim Cmd As New ADODB.Command
    Set Cmd.ActiveConnection = oConexaoADO
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "upBOLETIM_" & sProc
    Cmd.CommandTimeout = 30

    Cmd.Parameters.Append Cmd.CreateParameter("@data", adDate, adParamInput, , Me.data)
    Cmd.Parameters.Append Cmd.CreateParameter("@sal_cd", adInteger, adParamInput, , Me.sal_cd)
    If sProc <> "CATRACA2" Then
        Cmd.Parameters.Append Cmd.CreateParameter("@fil_cd", adInteger, adParamInput, , Me.fil_cd)
    End If
    
    'Select Case sProc
    '    Case "CAPA"
    '    Case Else
    '        Cmd.Parameters.Append Cmd.CreateParameter("@fil_cd", adInteger, adParamInput, , Me.fil_cd)
    'End Select

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
    Me.data = Empty
    Me.sal_cd = 0
    Me.fil_cd = 0
End Sub



