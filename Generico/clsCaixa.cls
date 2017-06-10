VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private oConexaoADO As ADODB.Connection

'local variable(s) to hold property value(s)
Private mvarDataExibicao As Date 'local copy
Private mvarfil_cd As Long 'local copy
Private mvarsal_cd As Integer 'local copy
Private mvarses_horario As Date 'local copy
Private mvarope_cd As Long 'local copy
Private mvaring_cd As Double 'local copy
Private mvarvcb_cd As Double 'local copy
Private mvarMensagemErro As String 'local copy
Private mvarCodigoErro As Long 'local copy
Private mvarcxa_cd       As Integer 'local copy

Public Property Get DiaSemana() As Integer
    
    Dim oRs As New ADODB.Recordset
    Dim bRet As Boolean
    
    bRet = ExecutaProcedure("upDIA_SEMANA", oRs)

    If Not oRs.EOF() Then
        DiaSemana = oRs.Fields("diaSemana")
    End If
    
    oRs.Close
    Set oRs = Nothing
    
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

Public Property Set ConexaoADO(ByRef vData As ADODB.Connection)
    Set oConexaoADO = vData
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
    
Public Property Let ses_horario(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ses_horario = 5
    mvarses_horario = vData
End Property

Public Property Get ses_horario() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ses_horario
    ses_horario = mvarses_horario
End Property
    
Public Property Let ope_cd(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ope_cd = 5
    mvarope_cd = vData
End Property

Public Property Get ope_cd() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ope_cd
    ope_cd = mvarope_cd
End Property

Public Property Let ing_cd(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ing_cd = 5
    mvaring_cd = vData
End Property

Public Property Get ing_cd() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ing_cd
    ing_cd = mvaring_cd
End Property

Public Property Let vcb_cd(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.vcb_cd = 5
    mvarvcb_cd = vData
End Property

Public Property Get vcb_cd() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.vcb_cd
    vcb_cd = mvarvcb_cd
End Property

Public Property Let DataExibicao(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DataExibicao = 5
    mvarDataExibicao = vData
End Property

Public Property Get DataExibicao() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DataExibicao
    DataExibicao = mvarDataExibicao
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


Public Function Combos(ByRef oRs As ADODB.Recordset) As Boolean
    Combos = ExecutaProcedure("upTB_PROG_COMBO_DATA", oRs)
End Function

Public Function FilmesCartaz(ByRef oRs As ADODB.Recordset) As Boolean
    FilmesCartaz = ExecutaProcedure("upTB_FILME_CARTAZ", oRs)
End Function

Public Function FilmesCartaz2(ByRef oRs As ADODB.Recordset) As Boolean
    FilmesCartaz2 = ExecutaProcedure("upTB_FILME_CARTAZ2", oRs)
End Function

Public Function IngressoOperacao(ByRef oRs As ADODB.Recordset) As Boolean
    IngressoOperacao = ExecutaProcedure("upINGRESSO_OPERACAO", oRs)
End Function

Public Function ComboOperacao(ByRef oRs As ADODB.Recordset) As Boolean
    ComboOperacao = ExecutaProcedure("upCOMBO_OPERACAO", oRs)
End Function

Public Function FechamentoCaixa(ByRef oRs As ADODB.Recordset) As Boolean
    FechamentoCaixa = ExecutaProcedure("upFECHA_CAIXA", oRs)
End Function

Public Function PreencheGrid(ByRef oRs As ADODB.Recordset) As Boolean
    PreencheGrid = ExecutaProcedure("upSESSOES_DIA", oRs)
End Function

Public Function PreencheGridPromocao(ByRef oRs As ADODB.Recordset) As Boolean
    PreencheGridPromocao = ExecutaProcedure("upSESSOES_DIA_PROMOCAO", oRs)
End Function

Public Function DataSistema() As Date
    Dim oRs As New ADODB.Recordset
    Call ExecutaProcedure("upDATA_SISTEMA", oRs)
    DataSistema = oRs.Fields(0)
    oRs.Close
End Function

Private Function ExecutaProcedure(ByVal sProc As String, Optional ByRef oRs As ADODB.Recordset) As Boolean

    On Error GoTo ExecutaProcedure_Error
    
    Dim Cmd As New ADODB.Command
    Set Cmd.ActiveConnection = oConexaoADO
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = sProc
    Cmd.CommandTimeout = 30

    Select Case sProc
        Case "upTB_PROG_COMBO_DATA", "upTB_FILME_CARTAZ", "upDIA_SEMANA", "upSESSOES_DIA", "upSESSOES_DIA_PROMOCAO", "upTB_FILME_CARTAZ2"
            Cmd.Parameters.Append Cmd.CreateParameter("@DataExibicao", adDate, adParamInput, , Me.DataExibicao)
    End Select
    
    Select Case sProc
        Case "upSESSOES_DIA", "upSESSOES_DIA_PROMOCAO"
            Cmd.Parameters.Append Cmd.CreateParameter("@fil_cd", adInteger, adParamInput, , Me.fil_cd)
    End Select

    Select Case sProc
        Case "upSESSOES_DIA", "upSESSOES_DIA_PROMOCAO"
            Cmd.Parameters.Append Cmd.CreateParameter("@sal_cd", adInteger, adParamInput, , Me.sal_cd)
    End Select

    Select Case sProc
        Case "upSESSOES_DIA_PROMOCAO"
            Cmd.Parameters.Append Cmd.CreateParameter("@ses_horario", adDate, adParamInput, , Me.ses_horario)
    End Select

    Select Case sProc
        
        Case "upINGRESSO_OPERACAO"
            Cmd.Parameters.Append Cmd.CreateParameter("@ope_cd", adInteger, adParamInput, , IIf(Me.ope_cd = 0, Null, Me.ope_cd))
            Cmd.Parameters.Append Cmd.CreateParameter("@ing_cd", adBigInt, adParamInput, , IIf(Me.ing_cd = 0, Null, Me.ing_cd))
        
        Case "upCOMBO_OPERACAO"
            Cmd.Parameters.Append Cmd.CreateParameter("@ope_cd", adInteger, adParamInput, , IIf(Me.ope_cd = 0, Null, Me.ope_cd))
            Cmd.Parameters.Append Cmd.CreateParameter("@vcb_cd", adBigInt, adParamInput, , IIf(Me.vcb_cd = 0, Null, Me.vcb_cd))
    
    End Select

    Cmd.Parameters.Append Cmd.CreateParameter("@Erro", adInteger, adParamOutput)
    Cmd.Parameters.Append Cmd.CreateParameter("@MsgErr", adVarChar, adParamOutput, 255)

    Select Case sProc
    
        Case "upTB_PROG_COMBO_DATA", "upTB_FILME_CARTAZ", _
             "upDIA_SEMANA", "upSESSOES_DIA", _
             "upSESSOES_DIA_PROMOCAO", "upINGRESSO_OPERACAO", _
             "upCOMBO_OPERACAO", "upFECHA_CAIXA", "upDATA_SISTEMA", "upTB_FILME_CARTAZ2"
             
            Set oRs = Cmd.Execute
        
        Case Else
            Call Cmd.Execute
            
    End Select
    
    mvarCodigoErro = IIf(IsNull(Cmd.Parameters.Item("@Erro").Value), 0, Cmd.Parameters.Item("@Erro").Value)
    mvarMensagemErro = IIf(IsNull(Cmd.Parameters.Item("@MsgErr").Value), "", Cmd.Parameters.Item("@MsgErr").Value)
    
    ExecutaProcedure = (mvarCodigoErro = 0)
    Exit Function

ExecutaProcedure_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description

End Function

Private Sub Class_Initialize()
    Me.DataExibicao = Empty
    Me.fil_cd = 0
End Sub

Public Function valorCaixa() As Double
    On Error GoTo ExecutaProcedure_Error
    
    Dim Cmd As New ADODB.Command
    
    valorCaixa = 0
    
    Set Cmd.ActiveConnection = oConexaoADO
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "upVALOR_EM_CAIXA"
    Cmd.CommandTimeout = 30

    Cmd.Parameters.Append Cmd.CreateParameter("@cxa_cd", adInteger, adParamInput, , IIf(Me.cxa_cd = 0, Null, Me.cxa_cd))
    Cmd.Parameters.Append Cmd.CreateParameter("valorCaixa", adDouble, adParamOutput, , 0)
    Cmd.Parameters.Append Cmd.CreateParameter("@Erro", adInteger, adParamOutput)
    Cmd.Parameters.Append Cmd.CreateParameter("@MsgErr", adVarChar, adParamOutput, 255)
    
    Call Cmd.Execute

    valorCaixa = IIf(IsNull(Cmd.Parameters.Item("valorCaixa").Value), 0, Cmd.Parameters.Item("valorCaixa").Value)
    
    mvarCodigoErro = IIf(IsNull(Cmd.Parameters.Item("@Erro").Value), 0, Cmd.Parameters.Item("@Erro").Value)
    mvarMensagemErro = IIf(IsNull(Cmd.Parameters.Item("@MsgErr").Value), "", Cmd.Parameters.Item("@MsgErr").Value)

    Exit Function

ExecutaProcedure_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description
End Function

Public Function posicaoCaixas(ByRef oRs As ADODB.Recordset) As Boolean
    On Error GoTo ExecutaProcedure_Error
    
    Dim Cmd As New ADODB.Command
    
    Set Cmd.ActiveConnection = oConexaoADO
    
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "upPOSICAO_CAIXAS"
    Cmd.CommandTimeout = 30

    Cmd.Parameters.Append Cmd.CreateParameter("@Erro", adInteger, adParamOutput)
    Cmd.Parameters.Append Cmd.CreateParameter("@MsgErr", adVarChar, adParamOutput, 255)
    
    Set oRs = Cmd.Execute

    mvarCodigoErro = IIf(IsNull(Cmd.Parameters.Item("@Erro").Value), 0, Cmd.Parameters.Item("@Erro").Value)
    mvarMensagemErro = IIf(IsNull(Cmd.Parameters.Item("@MsgErr").Value), "", Cmd.Parameters.Item("@MsgErr").Value)

    posicaoCaixas = (mvarCodigoErro = 0)
    Exit Function

ExecutaProcedure_Error:
    mvarCodigoErro = Err.Number
    mvarMensagemErro = Err.Description
End Function

    
