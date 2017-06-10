VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C9E1AFB0-1172-11D7-83AD-0050DA238ADA}#1.0#0"; "Coptr19.ocx"
Begin VB.Form frmBoletimAdm1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Boletim Administrativo - Parcial"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   544
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   492
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7LCtl.VSFlexGrid vsfBoletim 
      Height          =   6615
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1365
      Width           =   5535
      _cx             =   9763
      _cy             =   11668
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmBoletimAdm1.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   2
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.CommandButton cmdImprime 
      Cancel          =   -1  'True
      Caption         =   "&Imprime"
      Height          =   390
      Left            =   5790
      TabIndex        =   5
      Top             =   1080
      Width           =   1438
   End
   Begin VB.CommandButton cmdCancela 
      Caption         =   "&Cancela"
      Height          =   390
      Left            =   5790
      TabIndex        =   4
      Top             =   600
      Width           =   1438
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   5790
      TabIndex        =   3
      Top             =   120
      Width           =   1438
   End
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   195
      Width           =   5535
      Begin VB.TextBox txtDtRef 
         Height          =   315
         Left            =   555
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   1410
      End
      Begin VB.ComboBox cbo_sala 
         Height          =   315
         ItemData        =   "frmBoletimAdm1.frx":004C
         Left            =   555
         List            =   "frmBoletimAdm1.frx":004E
         TabIndex        =   8
         Top             =   630
         Width           =   4545
      End
      Begin VB.CommandButton cmdProcura 
         Height          =   315
         Left            =   5100
         Picture         =   "frmBoletimAdm1.frx":0050
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Pesquisa"
         Top             =   630
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sala:"
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   660
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         Height          =   195
         Left            =   90
         TabIndex        =   1
         Top             =   300
         Width           =   390
      End
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   6330
      Top             =   1845
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   0   'False
      Handshaking     =   1
      OutBufferSize   =   1311
      ParityReplace   =   32
      RThreshold      =   1
      BaudRate        =   19200
   End
   Begin VB.Frame fraAviso 
      Height          =   1890
      Left            =   375
      TabIndex        =   10
      Top             =   2760
      Visible         =   0   'False
      Width           =   6630
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Aguarde Carregando Boletins..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1095
         TabIndex        =   11
         Top             =   735
         Width           =   4425
      End
   End
   Begin COPTRLib.OPOSPOSPrinter OPOSPOSPrinter1 
      Left            =   6570
      Top             =   2985
      _Version        =   65536
      _ExtentX        =   609
      _ExtentY        =   582
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBoletimAdm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LinhaNormal
Private LinhaTitulo
Private LinhaDupla

Dim sala    As Collection
Dim filme   As Collection
Dim sesExcl As Collection

'Public linha()   As String
'Dim tpLinha() As String

Dim impriminto As Boolean
Dim dataRef    As Date
Dim inicio     As Boolean

Private Sub cbo_sala_Click()
    If cbo_sala.ListCount > 0 Then
        If cbo_sala.ListIndex <> -1 Then
            If Mid(cbo_sala.Text, 1, 1) = "#" Then
                MsgBox "Utilize boletim final para impressão deste item.", vbExclamation, App.ProductName
            Else
                Call cmdOK_Click
            End If
        End If
    End If
End Sub

Private Sub cmdProcura_Click()
    Dim clsCaixa As New clsCaixa
    Dim oRsAux   As ADODB.Recordset
    
    cbo_sala.Clear
    Set sala = New Collection
    Set filme = New Collection
    Set sesExcl = New Collection
    
    Set clsCaixa.ConexaoADO = dbConnect
    clsCaixa.DataExibicao = dataRef
    
    DoEvents
    fraAviso.ZOrder 0
    fraAviso.Visible = True
    DoEvents
    
    If clsCaixa.FilmesCartaz2(oRsAux) Then
        If Not (oRsAux.BOF And oRsAux.EOF) Then
            Do While Not oRsAux.EOF
                If oRsAux.Fields("ses_excl").Value = "N" Then
                    cbo_sala.AddItem oRsAux.Fields("Sala - Filme").Value
                ElseIf oRsAux.Fields("ses_excl").Value = "S" Then
                    cbo_sala.AddItem "* " & oRsAux.Fields("Sala - Filme").Value
                ElseIf oRsAux.Fields("ses_excl").Value = "P" Then
                    cbo_sala.AddItem "# " & oRsAux.Fields("Sala - Filme").Value
                ElseIf oRsAux.Fields("ses_excl").Value = "Q" Then
                    cbo_sala.AddItem "#* " & oRsAux.Fields("Sala - Filme").Value
                End If
                sala.Add oRsAux.Fields("sal_cd").Value
                filme.Add oRsAux.Fields("fil_cd").Value
                sesExcl.Add oRsAux.Fields("ses_excl").Value
                
                oRsAux.MoveNext
            Loop
            
            cbo_sala.ListIndex = 0
        End If
    End If

    DoEvents
    fraAviso.Visible = False
    vsfBoletim.ZOrder 0
    DoEvents
End Sub

Private Sub Form_Activate()
    If inicio Then
        txtDtRef.Text = Format(dataRef, "dd/mm/yyyy")
        Call cmdProcura_Click
        inicio = False
    End If
End Sub

Private Sub Form_Load()
    Dim gerais As New clsGerais
    
    'Call AjustaCOM(MSComm)
    'SLS 18/01/2012
    'If abreImp(OPOSPOSPrinter1) Then
    'End If
    
    Set gerais.ConexaoADO = dbConnect
    
    dataRef = gerais.DataRefAtu
    
    impriminto = False
    inicio = True
End Sub

Private Sub cmdCancela_Click()
    If Not impriminto Then
        Unload Me
    End If
End Sub

Private Sub cmdOK_Click()
    If cbo_sala.ListCount = 0 Or cbo_sala.ListIndex = -1 Then
        MsgBox "Sala deve ser informada!", vbCritical, App.ProductName
        Exit Sub
    End If
    
    impriminto = True
    Call ImprimeBoletim(dataRef, sala.Item(cbo_sala.ListIndex + 1), filme.Item(cbo_sala.ListIndex + 1), sesExcl.Item(cbo_sala.ListIndex + 1))
    impriminto = False
End Sub

Private Sub ImprimeBoletim(ByVal dData As Date, ByVal iSala As Integer, ByVal iFilme As Long, ByVal sSesExcl As String)

    On Error GoTo ImprimeBoletim_erro
    
    Dim clsBoletim  As New clsBoletim
    Dim oRsCapa     As New ADODB.Recordset
    Dim oRsAux      As New ADODB.Recordset
    Dim sSessoes    As String
    Dim dValorTotal As Double
    Dim iQtdeTotal  As Integer
    Dim dValorIng   As Double
    Dim dValorGeral As Double
    Dim iQtdeGeral  As Double
    Dim dTotalInt   As Double
    Dim dTotalMeia  As Double
    Dim dTotalCort  As Double
        
    DoEvents
    fraAviso.ZOrder 0
    fraAviso.Visible = True
    DoEvents
        
    LinhaNormal = Chr$(27) & Chr$(33) & Chr$(10)
    LinhaTitulo = Chr$(27) & Chr$(33) & Chr$(60)
    LinhaDupla = Chr$(27) & Chr$(33) & Chr$(40)
    
    vsfBoletim.Clear
    vsfBoletim.Rows = 0
    'vsfBoletim.Cols = 2
    vsfBoletim.FontName = "Courier"
    vsfBoletim.FontSize = 6
    vsfBoletim.GridLines = flexGridNone
    
    Set clsBoletim.ConexaoADO = dbConnect
    
    clsBoletim.data = dData
    clsBoletim.sal_cd = iSala
    clsBoletim.fil_cd = iFilme
    clsBoletim.ses_excl = sSesExcl
    
    If Not clsBoletim.Capa(oRsCapa) Then
        MsgBox clsBoletim.MensagemErro, vbCritical, App.ProductName
        GoTo ImprimeBoletim_fim
    End If
    
    ReDim linha(1 To 1) As String
    ReDim tplinha(1 To 1) As String
    linha(UBound(linha)) = ""
    tplinha(UBound(tplinha)) = "N"
    
    Do While Not oRsCapa.EOF()
        If IsNull(oRsCapa.Fields("dt_abertura")) Then
            MsgBox "Não houve movimento na data para a sala selecionada!", vbInformation, App.ProductName
            oRsCapa.Close
            GoTo ImprimeBoletim_fim
        End If

        vsfBoletim.AddItem (LinhaTitulo & Chr(9) & "BOLETIM PARCIAL" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = "BOLETIM PARCIAL"
        tplinha(UBound(tplinha)) = "T"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = ""
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("Exibidora : " & EliminaAcentos(oRsCapa.Fields("emp_nm")), COLUNAS_IMP, esquerda) & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = Alinha("Exibidora : " & EliminaAcentos(oRsCapa.Fields("emp_nm")), COLUNAS_IMP, esquerda)
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = ""
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("Cine .....: " & EliminaAcentos(oRsCapa.Fields("cin_nm")), COLUNAS_IMP, esquerda) & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = Alinha("Cine .....: " & EliminaAcentos(oRsCapa.Fields("cin_nm")), COLUNAS_IMP, esquerda)
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("Sala .....: " & EliminaAcentos(oRsCapa.Fields("sal_desc")), COLUNAS_IMP, esquerda) & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = Alinha("Sala .....: " & EliminaAcentos(oRsCapa.Fields("sal_desc")), COLUNAS_IMP, esquerda)
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("Lotacao ..: " & oRsCapa.Fields("sal_lugares"), COLUNAS_IMP, esquerda) & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = Alinha("Lotacao ..: " & oRsCapa.Fields("sal_lugares"), COLUNAS_IMP, esquerda)
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = ""
        tplinha(UBound(tplinha)) = "N"
        
        If oRsCapa.Fields("pre_estreia").Value = "N" Then
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("Movimento : " & Format(dData, "dd/mm/yyyy"), COLUNAS_IMP, esquerda) & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = Alinha("Movimento : " & Format(dData, "dd/mm/yyyy"), COLUNAS_IMP, esquerda)
            tplinha(UBound(tplinha)) = "N"
        Else
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("Movimento : " & Format(DateAdd("d", 1, dData), "dd/mm/yyyy"), COLUNAS_IMP, esquerda) & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = Alinha("Movimento : " & Format(DateAdd("d", 1, dData), "dd/mm/yyyy"), COLUNAS_IMP, esquerda)
            tplinha(UBound(tplinha)) = "N"
        End If
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("Abertura .: " & Format(oRsCapa.Fields("dt_abertura"), "dd/mm/yyyy hh:mm:ss"), COLUNAS_IMP, esquerda) & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = Alinha("Abertura .: " & Format(oRsCapa.Fields("dt_abertura"), "dd/mm/yyyy hh:mm:ss"), COLUNAS_IMP, esquerda)
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("Emissao ..: " & Format(oRsCapa.Fields("dt_atual"), "dd/mm/yyyy hh:mm:ss"), COLUNAS_IMP, esquerda) & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = Alinha("Emissao ..: " & Format(oRsCapa.Fields("dt_atual"), "dd/mm/yyyy hh:mm:ss"), COLUNAS_IMP, esquerda)
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = ""
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("Filme ....: " & EliminaAcentos(Mid(oRsCapa.Fields("fil_nm"), 1, 26)), COLUNAS_IMP, esquerda) & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = Alinha("Filme ....: " & EliminaAcentos(Mid(oRsCapa.Fields("fil_nm"), 1, 26)), COLUNAS_IMP, esquerda)
        tplinha(UBound(tplinha)) = "N"
        
        If Trim(Mid(oRsCapa.Fields("fil_nm"), 28, 27)) <> "" Then
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("            " & EliminaAcentos(Mid(oRsCapa.Fields("fil_nm"), 27, 26)), COLUNAS_IMP, esquerda) & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = Alinha("            " & EliminaAcentos(Mid(oRsCapa.Fields("fil_nm"), 27, 26)), COLUNAS_IMP, esquerda)
            tplinha(UBound(tplinha)) = "N"
        End If
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("Distrib ..: " & EliminaAcentos(oRsCapa.Fields("fil_distribuidora")), COLUNAS_IMP, esquerda) & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = Alinha("Distrib ..: " & EliminaAcentos(oRsCapa.Fields("fil_distribuidora")), COLUNAS_IMP, esquerda)
        tplinha(UBound(tplinha)) = "N"
        
        clsBoletim.fil_cd = oRsCapa.Fields("fil_cd")
        
        'Set oRsAux = New ADODB.Recordset
                
        'If Not clsBoletim.SessoesFilme(oRsAux) Then
        '    MsgBox clsBoletim.MensagemErro, vbCritical, App.ProductName
        '    GoTo ImprimeBoletim_fim
        'End If
        
        'sSessoes = ""
        
        'Do While Not oRsAux.EOF()
        '    sSessoes = sSessoes & Format(oRsAux.Fields("ses_horario"), "hh:mm") & " "
        '    oRsAux.MoveNext
        'Loop
                
        'oRsAux.Close
        
        'vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("Sessoes ..: " & Mid(sSessoes, 1, 24), COLUNAS_IMP, Esquerda) & Chr(9) & Chr$(10))
        '
        'If Trim(Mid(sSessoes, 25, 24)) <> "" Then
        '    vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("            " & Mid(sSessoes, 25, 24), COLUNAS_IMP, Esquerda) & Chr(9) & Chr$(10))
        'End If
        
        'If Trim(Mid(sSessoes, 49, 24)) <> "" Then
        '    vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("            " & Mid(sSessoes, 49, 24), COLUNAS_IMP, Esquerda) & Chr(9) & Chr$(10))
        'End If
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("Periodo ..: " & Format(oRsCapa.Fields("prg_dt_ini"), "dd/mm/yyyy") & " a " & Format(oRsCapa.Fields("prg_dt_fim"), "dd/mm/yyyy"), COLUNAS_IMP, esquerda) & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = Alinha("Periodo ..: " & Format(oRsCapa.Fields("prg_dt_ini"), "dd/mm/yyyy") & " a " & Format(oRsCapa.Fields("prg_dt_fim"), "dd/mm/yyyy"), COLUNAS_IMP, esquerda)
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = ""
        tplinha(UBound(tplinha)) = "N"

        ' ****************** Vendas do Dia
        
        Set oRsAux = New ADODB.Recordset
                
        If Not clsBoletim.VendasDia(oRsAux) Then
            MsgBox clsBoletim.MensagemErro, vbCritical, App.ProductName
            GoTo ImprimeBoletim_fim
        End If

        iQtdeTotal = 0
        dValorTotal = 0
        dValorGeral = 0
        iQtdeGeral = 0
        
        vsfBoletim.AddItem (LinhaDupla & Chr(9) & "VENDAS DO DIA" & Chr(9))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = "VENDAS DO DIA"
        tplinha(UBound(tplinha)) = "S"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = ""
        tplinha(UBound(tplinha)) = "N"
        
        Do While Not oRsAux.EOF()
        
            dValorIng = 0
            
            If Val(oRsAux.Fields("qtde")) > 0 Then
                dValorIng = oRsAux.Fields("ing_valor") / oRsAux.Fields("qtde")
            End If
            
            dValorTotal = dValorTotal + oRsAux.Fields("ing_valor")
            iQtdeTotal = iQtdeTotal + oRsAux.Fields("qtde")
            
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha(TrataTipo(oRsAux.Fields("igt_desc")), 16, esquerda, ".") & " " & _
                                Alinha(Format(dValorIng, "0.00"), 6, Direita, " ") & " " & _
                                Alinha(oRsAux.Fields("qtde"), 4, Direita, " ") & " " & _
                                Alinha(Format(oRsAux.Fields("ing_valor"), "0.00"), 9, Direita, " ") & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = Alinha(TrataTipo(oRsAux.Fields("igt_desc")), 16, esquerda, ".") & " " & _
                                   Alinha(Format(dValorIng, "0.00"), 6, Direita, " ") & " " & _
                                   Alinha(oRsAux.Fields("qtde"), 4, Direita, " ") & " " & _
                                   Alinha(Format(oRsAux.Fields("ing_valor"), "0.00"), 9, Direita, " ")
            tplinha(UBound(tplinha)) = "N"
            
            oRsAux.MoveNext
            
        Loop
        
        oRsAux.Close
                
        vsfBoletim.AddItem (String(COLUNAS_IMP, Chr$(196)) & Chr(9) & "" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = String(COLUNAS_IMP, "_")
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("TOTAL VENDAS DO DIA", 23, esquerda, " ") & " " & _
                            Alinha(iQtdeTotal, 4, Direita, " ") & " " & _
                            Alinha(Format(dValorTotal, "0.00"), 9, Direita, " ") & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = Alinha("TOTAL VENDAS DO DIA", 23, esquerda, " ") & " " & _
                               Alinha(iQtdeTotal, 4, Direita, " ") & " " & _
                               Alinha(Format(dValorTotal, "0.00"), 9, Direita, " ")
        tplinha(UBound(tplinha)) = "N"
        
        dValorGeral = dValorGeral + dValorTotal
        iQtdeGeral = iQtdeGeral + iQtdeTotal
                
        dValorTotal = 0
        iQtdeTotal = 0
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = ""
        tplinha(UBound(tplinha)) = "N"
        
        ' ****************** Pré-Venda
        
        Set oRsAux = New ADODB.Recordset
                
        If Not clsBoletim.PreVenda(oRsAux) Then
            MsgBox clsBoletim.MensagemErro, vbCritical, App.ProductName
            GoTo ImprimeBoletim_fim
        End If
         
        If Not oRsAux.EOF() Then
            vsfBoletim.AddItem (LinhaDupla & Chr(9) & "PRE-VENDA" & Chr(9))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = "PRE-VENDA"
            tplinha(UBound(tplinha)) = "S"
            
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = ""
            tplinha(UBound(tplinha)) = "N"
        
            Do While Not oRsAux.EOF()
        
                dValorIng = 0
            
                If Val(oRsAux.Fields("qtde")) > 0 Then
                   dValorIng = oRsAux.Fields("ing_valor") / oRsAux.Fields("qtde")
                End If
        
                dValorTotal = dValorTotal + oRsAux.Fields("ing_valor")
                iQtdeTotal = iQtdeTotal + oRsAux.Fields("qtde")
            
                vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha(TrataTipo(oRsAux.Fields("igt_desc")), 16, esquerda, ".") & " " & _
                                    Alinha(Format(dValorIng, "0.00"), 6, Direita, " ") & " " & _
                                    Alinha(oRsAux.Fields("qtde"), 4, Direita, " ") & " " & _
                                    Alinha(Format(oRsAux.Fields("ing_valor"), "0.00"), 9, Direita, " ") & Chr(9) & Chr$(10))
                ReDim Preserve linha(1 To UBound(linha) + 1) As String
                ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
                linha(UBound(linha)) = Alinha(TrataTipo(oRsAux.Fields("igt_desc")), 16, esquerda, ".") & " " & _
                                       Alinha(Format(dValorIng, "0.00"), 6, Direita, " ") & " " & _
                                       Alinha(oRsAux.Fields("qtde"), 4, Direita, " ") & " " & _
                                       Alinha(Format(oRsAux.Fields("ing_valor"), "0.00"), 9, Direita, " ")
                tplinha(UBound(tplinha)) = "N"
                
                oRsAux.MoveNext
            
            Loop
        
            vsfBoletim.AddItem (String(COLUNAS_IMP, Chr$(196)) & Chr(9) & "" & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = String(COLUNAS_IMP, "_")
            tplinha(UBound(tplinha)) = "N"
            
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("TOTAL PRE-VENDA", 23, esquerda, " ") & " " & _
                                Alinha(iQtdeTotal, 4, Direita, " ") & " " & _
                                Alinha(Format(dValorTotal, "0.00"), 9, Direita, " ") & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = Alinha("TOTAL PRE-VENDA", 23, esquerda, " ") & " " & _
                                   Alinha(iQtdeTotal, 4, Direita, " ") & " " & _
                                   Alinha(Format(dValorTotal, "0.00"), 9, Direita, " ")
            tplinha(UBound(tplinha)) = "N"
            
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = ""
            tplinha(UBound(tplinha)) = "N"
        End If
        
        oRsAux.Close
        
        dValorGeral = dValorGeral + dValorTotal
        iQtdeGeral = iQtdeGeral + iQtdeTotal
                
        dValorTotal = 0
        iQtdeTotal = 0
        
        
        ' ****************** Devolução
        
        Set oRsAux = New ADODB.Recordset
                
        If Not clsBoletim.Devolucao(oRsAux) Then
            MsgBox clsBoletim.MensagemErro, vbCritical, App.ProductName
            GoTo ImprimeBoletim_fim
        End If

        If Not oRsAux.EOF() Then
            vsfBoletim.AddItem (LinhaDupla & Chr(9) & "DEVOLUCAO" & Chr(9))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = "DEVOLUCAO"
            tplinha(UBound(tplinha)) = "S"
            
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = ""
            tplinha(UBound(tplinha)) = "N"
        
            Do While Not oRsAux.EOF()
        
                dValorIng = 0
            
                If Val(oRsAux.Fields("qtde")) > 0 Then
                    dValorIng = oRsAux.Fields("ing_valor") / oRsAux.Fields("qtde")
                End If
        
                dValorTotal = dValorTotal + oRsAux.Fields("ing_valor")
                iQtdeTotal = iQtdeTotal + oRsAux.Fields("qtde")
            
                vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha(TrataTipo(oRsAux.Fields("igt_desc")), 16, esquerda, ".") & " " & _
                                    Alinha(Format(dValorIng, "0.00"), 6, Direita, " ") & " " & _
                                    Alinha(oRsAux.Fields("qtde"), 4, Direita, " ") & " " & _
                                    Alinha(Format(oRsAux.Fields("ing_valor"), "0.00"), 9, Direita, " ") & Chr(9) & Chr$(10))
                ReDim Preserve linha(1 To UBound(linha) + 1) As String
                ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
                linha(UBound(linha)) = Alinha(TrataTipo(oRsAux.Fields("igt_desc")), 16, esquerda, ".") & " " & _
                                       Alinha(Format(dValorIng, "0.00"), 6, Direita, " ") & " " & _
                                       Alinha(oRsAux.Fields("qtde"), 4, Direita, " ") & " " & _
                                       Alinha(Format(oRsAux.Fields("ing_valor"), "0.00"), 9, Direita, " ")
                tplinha(UBound(tplinha)) = "N"
                
                oRsAux.MoveNext
            
            Loop
        
            vsfBoletim.AddItem (String(COLUNAS_IMP, Chr$(196)) & Chr(9) & "" & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = String(COLUNAS_IMP, "_")
            tplinha(UBound(tplinha)) = "N"
            
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("TOTAL DEVOLUCAO", 23, esquerda, " ") & " " & _
                                Alinha(iQtdeTotal, 4, Direita, " ") & " " & _
                                Alinha(Format(dValorTotal, "0.00"), 9, Direita, " ") & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = Alinha("TOTAL DEVOLUCAO", 23, esquerda, " ") & " " & _
                                   Alinha(iQtdeTotal, 4, Direita, " ") & " " & _
                                   Alinha(Format(dValorTotal, "0.00"), 9, Direita, " ")
            tplinha(UBound(tplinha)) = "N"
            
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = ""
            tplinha(UBound(tplinha)) = "N"
        End If
        
        oRsAux.Close
                
        
        dValorGeral = dValorGeral - dValorTotal
        iQtdeGeral = iQtdeGeral - iQtdeTotal
                
        dValorTotal = 0
        iQtdeTotal = 0
        
                
        vsfBoletim.AddItem (String(COLUNAS_IMP, Chr$(196)) & Chr(9) & "" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = String(COLUNAS_IMP, "_")
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("TOTAL GERAL", 23, esquerda, " ") & " " & _
                            Alinha(iQtdeGeral, 4, Direita, " ") & " " & _
                            Alinha(Format(dValorGeral, "0.00"), 9, Direita, " ") & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = Alinha("TOTAL GERAL", 23, esquerda, " ") & " " & _
                               Alinha(iQtdeGeral, 4, Direita, " ") & " " & _
                               Alinha(Format(dValorGeral, "0.00"), 9, Direita, " ")
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (String(COLUNAS_IMP, Chr$(196)) & Chr(9) & "" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = String(COLUNAS_IMP, "_")
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = ""
        tplinha(UBound(tplinha)) = "N"
        
        '****Cortesia
        Set oRsAux = New ADODB.Recordset
                
        If Not clsBoletim.Cortesia(oRsAux) Then
            MsgBox clsBoletim.MensagemErro, vbCritical, App.ProductName
            GoTo ImprimeBoletim_fim
        End If

        Do While Not oRsAux.EOF()
            
            iQtdeGeral = iQtdeGeral + oRsAux.Fields("qtde")
            
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha(TrataTipo(oRsAux.Fields("igt_desc")) & oRsAux.Fields("desc_dia"), 16, esquerda, " ") & " " & _
                                Alinha("", 6, Direita, " ") & " " & _
                                Alinha(oRsAux.Fields("qtde"), 4, Direita, " ") & " " & _
                                Alinha("", 9, Direita, " ") & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = Alinha(TrataTipo(oRsAux.Fields("igt_desc")) & oRsAux.Fields("desc_dia"), 16, esquerda, " ") & " " & _
                                   Alinha("", 6, Direita, " ") & " " & _
                                   Alinha(oRsAux.Fields("qtde"), 4, Direita, " ") & " " & _
                                   Alinha("", 9, Direita, " ")
            tplinha(UBound(tplinha)) = "N"
            
            oRsAux.MoveNext
            
        Loop
        
        oRsAux.Close
        
        vsfBoletim.AddItem (String(COLUNAS_IMP, Chr$(196)) & Chr(9) & "" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = String(COLUNAS_IMP, "_")
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("TOTAL PUBLICO", 23, esquerda, " ") & " " & _
                            Alinha(iQtdeGeral, 4, Direita, " ") & " " & _
                            Alinha(" ", 9, Direita, " ") & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = Alinha("TOTAL PUBLICO", 23, esquerda, " ") & " " & _
                               Alinha(iQtdeGeral, 4, Direita, " ") & " " & _
                               Alinha(" ", 9, Direita, " ")
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = ""
        tplinha(UBound(tplinha)) = "N"
        
        
        ' ****************** Total por Sessão
        
        Set oRsAux = New ADODB.Recordset
                
        If Not clsBoletim.TotalSessao(oRsAux) Then
            MsgBox clsBoletim.MensagemErro, vbCritical, App.ProductName
            GoTo ImprimeBoletim_fim
        End If

        vsfBoletim.AddItem (LinhaDupla & Chr(9) & "TOTAL POR SESSAO" & Chr(9))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = "TOTAL POR SESSAO"
        tplinha(UBound(tplinha)) = "S"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = ""
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & "SESSAO" & Alinha("INTEIRA", 8, Direita) & Alinha("MEIA", 9, Direita) & Alinha("CORTESIA", 9, Direita) & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = "SESSAO" & Alinha("INTEIRA", 8, Direita) & Alinha("MEIA", 9, Direita) & Alinha("CORTESIA", 9, Direita)
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = ""
        tplinha(UBound(tplinha)) = "N"
        
        dTotalInt = 0
        dTotalMeia = 0
        dTotalCort = 0
        
        Do While Not oRsAux.EOF()
        
            dTotalInt = dTotalInt + oRsAux.Fields("qtde_int")
            dTotalMeia = dTotalMeia + oRsAux.Fields("qtde_meia")
            dTotalCort = dTotalCort + oRsAux.Fields("qtde_cort")
            
            If oRsAux.Fields("hora_excl") = "N" Then
                vsfBoletim.AddItem (LinhaNormal & Chr(9) & " " & Format(oRsAux.Fields("ses_horario"), "hh:mm") & _
                                    Alinha(Format(oRsAux.Fields("qtde_int"), "###,##0"), 8, Direita, " ") & " " & _
                                    Alinha(Format(oRsAux.Fields("qtde_meia"), "###,##0"), 8, Direita, " ") & " " & _
                                    Alinha(Format(oRsAux.Fields("qtde_cort"), "###,##0"), 8, Direita, " ") & Chr(9) & Chr$(10))
            Else
                vsfBoletim.AddItem (LinhaNormal & Chr(9) & "*" & Format(oRsAux.Fields("ses_horario"), "hh:mm") & _
                                    Alinha(Format(oRsAux.Fields("qtde_int"), "###,##0"), 8, Direita, " ") & " " & _
                                    Alinha(Format(oRsAux.Fields("qtde_meia"), "###,##0"), 8, Direita, " ") & " " & _
                                    Alinha(Format(oRsAux.Fields("qtde_cort"), "###,##0"), 8, Direita, " ") & Chr(9) & Chr$(10))
            End If
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            
            If oRsAux.Fields("hora_excl") = "N" Then
                linha(UBound(linha)) = " " & Format(oRsAux.Fields("ses_horario"), "hh:mm") & _
                                       Alinha(Format(oRsAux.Fields("qtde_int"), "###,##0"), 8, Direita, " ") & " " & _
                                       Alinha(Format(oRsAux.Fields("qtde_meia"), "###,##0"), 8, Direita, " ") & " " & _
                                       Alinha(Format(oRsAux.Fields("qtde_cort"), "###,##0"), 8, Direita, " ")
            Else
                linha(UBound(linha)) = "*" & Format(oRsAux.Fields("ses_horario"), "hh:mm") & _
                                       Alinha(Format(oRsAux.Fields("qtde_int"), "###,##0"), 8, Direita, " ") & " " & _
                                       Alinha(Format(oRsAux.Fields("qtde_meia"), "###,##0"), 8, Direita, " ") & " " & _
                                       Alinha(Format(oRsAux.Fields("qtde_cort"), "###,##0"), 8, Direita, " ")
            End If
            tplinha(UBound(tplinha)) = "N"
            
            oRsAux.MoveNext
            
        Loop
    
        oRsAux.Close
                
        vsfBoletim.AddItem (String(COLUNAS_IMP, Chr$(196)) & Chr(9) & "" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = String(COLUNAS_IMP, "_")
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & " TOTAL" & _
                            Alinha(Format(dTotalInt, "###,##0"), 8, Direita, " ") & " " & _
                            Alinha(Format(dTotalMeia, "###,##0"), 8, Direita, " ") & " " & _
                            Alinha(Format(dTotalCort, "###,##0"), 8, Direita, " ") & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = " TOTAL" & _
                               Alinha(Format(dTotalInt, "###,##0"), 8, Direita, " ") & " " & _
                               Alinha(Format(dTotalMeia, "###,##0"), 8, Direita, " ") & " " & _
                               Alinha(Format(dTotalCort, "###,##0"), 8, Direita, " ")
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = ""
        tplinha(UBound(tplinha)) = "N"
        
        Dim dReceitaLiq As Double
        Dim dReceitaBruta As Double
        Dim dTotalDesc As Double
        
        dReceitaLiq = 0
        dTotalDesc = 0
        
        dReceitaBruta = Format(dValorGeral, "0.00")
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("RECEITA BRUTA", 22, esquerda, ".")) & Alinha(Format(dReceitaBruta, "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10)
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = Alinha("RECEITA BRUTA", 22, esquerda, ".") & Alinha(Format(dReceitaBruta, "#,##0.00"), 10, Direita)
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = ""
        tplinha(UBound(tplinha)) = "N"
        
        If dCustoIngresso > 0 Then
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("CUSTO INGRESSO (" & Format(dCustoIngresso, "0.00") & ")", 22, esquerda, ".")) & Alinha(Format(iQtdeGeral * dCustoIngresso, "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10)
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = Alinha("CUSTO INGRESSO (" & Format(dCustoIngresso, "0.00") & ")", 22, esquerda, ".") & Alinha(Format(iQtdeGeral * dCustoIngresso, "#,##0.00"), 10, Direita)
            tplinha(UBound(tplinha)) = "N"
        End If
        
        dReceitaLiq = dReceitaLiq - Format(iQtdeGeral * dCustoIngresso, "0.00")
        dTotalDesc = dTotalDesc + Format(iQtdeGeral * dCustoIngresso, "0.00")
        
        If dImpostoMun > 0 Then
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("IMP MUNICIPAL (" & Format(dImpostoMun, "0.00") & "%)", 22, esquerda, ".")) & Alinha(Format(dReceitaBruta * (dImpostoMun / 100), "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10)
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = Alinha("IMP MUNICIPAL (" & Format(dImpostoMun, "0.00") & "%)", 22, esquerda, ".") & Alinha(Format(dReceitaBruta * (dImpostoMun / 100), "#,##0.00"), 10, Direita)
            tplinha(UBound(tplinha)) = "N"
        End If
        
        dReceitaLiq = dReceitaLiq - Format(dReceitaBruta * (dImpostoMun / 100), "0.00")
        dTotalDesc = dTotalDesc + Format(dReceitaBruta * (dImpostoMun / 100), "0.00")
        
        If dDireitosAut > 0 Then
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("DIR AUTORAIS (" & Format(dDireitosAut, "0.00") & "%)", 22, esquerda, ".")) & Alinha(Format(dReceitaBruta * (dDireitosAut / 100), "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10)
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = Alinha("DIR AUTORAIS (" & Format(dDireitosAut, "0.00") & "%)", 22, esquerda, ".") & Alinha(Format(dReceitaBruta * (dDireitosAut / 100), "#,##0.00"), 10, Direita)
            tplinha(UBound(tplinha)) = "N"
        End If
        
        dReceitaLiq = dReceitaLiq - Format(dReceitaBruta * (dDireitosAut / 100), "0.00")
        dTotalDesc = dTotalDesc + Format(dReceitaBruta * (dDireitosAut / 100), "0.00")
        
        If dOutros > 0 Then
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("OUTROS (" & Format(dOutros, "0.00") & "%)", 22, esquerda, ".")) & Alinha(Format(dReceitaBruta * (dOutros / 100), "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10)
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = Alinha("OUTROS (" & Format(dOutros, "0.00") & "%)", 22, esquerda, ".") & Alinha(Format(dReceitaBruta * (dOutros / 100), "#,##0.00"), 10, Direita)
            tplinha(UBound(tplinha)) = "N"
        End If
        
        dReceitaLiq = dReceitaLiq - Format(dReceitaBruta * (dOutros / 100), "0.00")
        dTotalDesc = dTotalDesc + Format(dReceitaBruta * (dOutros / 100), "0.00")
        
        If dTotalDesc > 0 Then
            vsfBoletim.AddItem (String(COLUNAS_IMP, Chr$(196)) & Chr(9) & "" & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = String(COLUNAS_IMP, "_")
            tplinha(UBound(tplinha)) = "N"
            
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("TOTAL DE DESCONTOS", 22, esquerda, ".")) & Alinha(Format(dTotalDesc, "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10)
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = Alinha("TOTAL DE DESCONTOS", 22, esquerda, ".") & Alinha(Format(dTotalDesc, "#,##0.00"), 10, Direita)
            tplinha(UBound(tplinha)) = "N"
        
            vsfBoletim.AddItem (String(COLUNAS_IMP, Chr$(196)) & Chr(9) & "" & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = String(COLUNAS_IMP, "_")
            tplinha(UBound(tplinha)) = "N"
            
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("RECEITA LIQUIDA", 22, esquerda, ".")) & Alinha(Format(dReceitaBruta - dTotalDesc, "#,##0.00"), 10, Direita) & Chr(9) & Chr$(10)
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = Alinha("RECEITA LIQUIDA", 22, esquerda, ".") & Alinha(Format(dReceitaBruta - dTotalDesc, "#,##0.00"), 10, Direita)
            tplinha(UBound(tplinha)) = "N"
        End If
        
        ' ****************** Venda Antecipada

        Dim sData As String
        
        Set oRsAux = New ADODB.Recordset
                
        If Not clsBoletim.VendaAntecipada(oRsAux) Then
            MsgBox clsBoletim.MensagemErro, vbCritical, App.ProductName
            GoTo ImprimeBoletim_fim
        End If

        If Not oRsAux.EOF() Then
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = ""
            tplinha(UBound(tplinha)) = "N"
            
            vsfBoletim.AddItem (LinhaDupla & Chr(9) & "VENDA ANTECIPADA" & Chr(9))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = "VENDA ANTECIPADA"
            tplinha(UBound(tplinha)) = "S"
        
            dValorTotal = 0
            iQtdeTotal = 0
        
            Do While Not oRsAux.EOF()
                If sData <> Format(oRsAux.Fields("sre_data"), "dd/mm/yyyy") Then
                    vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
                    ReDim Preserve linha(1 To UBound(linha) + 1) As String
                    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
                    linha(UBound(linha)) = ""
                    tplinha(UBound(tplinha)) = "N"
                    
                    sData = Format(oRsAux.Fields("sre_data"), "dd/mm/yyyy")
                    vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("PARA DIA", 16, esquerda, ".")) & " " & sData & Space(10) & Chr(9) & Chr$(10)
                    ReDim Preserve linha(1 To UBound(linha) + 1) As String
                    ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
                    linha(UBound(linha)) = Alinha("PARA DIA", 18, esquerda, ".") & " " & sData '& Space(8)
                    tplinha(UBound(tplinha)) = "N"
                End If
            
                dValorIng = 0
            
                If Val(oRsAux.Fields("qtde")) > 0 Then
                    dValorIng = oRsAux.Fields("ing_valor") / oRsAux.Fields("qtde")
                End If
        
                dValorTotal = dValorTotal + oRsAux.Fields("ing_valor")
                iQtdeTotal = iQtdeTotal + oRsAux.Fields("qtde")
            
                vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha(TrataTipo(oRsAux.Fields("igt_desc")), 16, esquerda, ".") & " " & _
                                    Alinha(Format(dValorIng, "0.00"), 6, Direita, " ") & " " & _
                                    Alinha(oRsAux.Fields("qtde"), 4, Direita, " ") & " " & _
                                    Alinha(Format(oRsAux.Fields("ing_valor"), "0.00"), 9, Direita, " ") & Chr(9) & Chr$(10))
                ReDim Preserve linha(1 To UBound(linha) + 1) As String
                ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
                linha(UBound(linha)) = Alinha(TrataTipo(oRsAux.Fields("igt_desc")), 16, esquerda, ".") & " " & _
                                       Alinha(Format(dValorIng, "0.00"), 6, Direita, " ") & " " & _
                                       Alinha(oRsAux.Fields("qtde"), 4, Direita, " ") & " " & _
                                       Alinha(Format(oRsAux.Fields("ing_valor"), "0.00"), 9, Direita, " ")
                tplinha(UBound(tplinha)) = "N"
                
                oRsAux.MoveNext
            Loop
        
            vsfBoletim.AddItem (String(COLUNAS_IMP, Chr$(196)) & Chr(9) & "" & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = String(COLUNAS_IMP, "_")
            tplinha(UBound(tplinha)) = "N"
            
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("TOTAL VENDA ANTECIPADA", 23, esquerda, " ") & " " & _
                                Alinha(iQtdeTotal, 4, Direita, " ") & " " & _
                                Alinha(Format(dValorTotal, "0.00"), 9, Direita, " ") & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = Alinha("TOTAL VENDA ANTECIPADA", 23, esquerda, " ") & " " & _
                                   Alinha(iQtdeTotal, 4, Direita, " ") & " " & _
                                   Alinha(Format(dValorTotal, "0.00"), 9, Direita, " ")
            tplinha(UBound(tplinha)) = "N"
            
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = ""
            tplinha(UBound(tplinha)) = "N"
        End If
        
        oRsAux.Close
        
        '******************* Talão
        
        Set oRsAux = New ADODB.Recordset
                
        If Not clsBoletim.numeracaoTalao(oRsAux) Then
            MsgBox clsBoletim.MensagemErro, vbCritical, App.ProductName
            GoTo ImprimeBoletim_fim
        End If
        
        If Not oRsAux.EOF() Then
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = ""
            tplinha(UBound(tplinha)) = "N"
            
            vsfBoletim.AddItem (LinhaDupla & Chr(9) & "NUMERACAO TALAO" & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = "NUMERACAO TALAO"
            tplinha(UBound(tplinha)) = "S"
        
            Do While Not oRsAux.EOF()
                vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha(TrataTipo(oRsAux.Fields("igt_desc")), 18, esquerda, ".") & " " & _
                                    Alinha(Format(oRsAux.Fields("numIni"), "000000"), 6, Direita, " ") & " " & _
                                    Alinha(Format(oRsAux.Fields("numFim"), "000000"), 6, Direita, " ") & Chr(9) & Chr$(10))
                ReDim Preserve linha(1 To UBound(linha) + 1) As String
                ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
                linha(UBound(linha)) = Alinha(TrataTipo(oRsAux.Fields("igt_desc")), 18, esquerda, ".") & " " & _
                                       Alinha(Format(oRsAux.Fields("numIni"), "000000"), 6, Direita, " ") & " " & _
                                       Alinha(Format(oRsAux.Fields("numFim"), "000000"), 6, Direita, " ")
                tplinha(UBound(tplinha)) = "N"
                
                oRsAux.MoveNext
            Loop
        End If
        
        oRsAux.Close
        
        
        '******************* Catraca
        Set oRsAux = New ADODB.Recordset
                
        If Not clsBoletim.Catraca(oRsAux) Then
            MsgBox clsBoletim.MensagemErro, vbCritical, App.ProductName
            GoTo ImprimeBoletim_fim
        End If
        
'        vsfBoletim.AddItem (Chr$(10))
        
        If Not oRsAux.EOF() Then
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = ""
            tplinha(UBound(tplinha)) = "N"
            
            Do While Not oRsAux.EOF()
                vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha(Trim(oRsAux.Fields("cat_nm").Value) & " - Iniciante : " & oRsAux.Fields("ctc_ini_cont").Value, COLUNAS_IMP, esquerda) & Chr(9) & Chr$(10))
                ReDim Preserve linha(1 To UBound(linha) + 1) As String
                ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
                linha(UBound(linha)) = Trim(oRsAux.Fields("cat_nm")) & " - Iniciante : " & oRsAux.Fields("ctc_ini_cont").Value
                tplinha(UBound(tplinha)) = "N"
            
'                vsfBoletim.AddItem (LinhaNormal & Chr(9) & "Iniciante : " & oRsAux.Fields("ctc_ini_cont").Value & Chr(9) & Chr$(10))
'                ReDim Preserve linha(1 To UBound(linha) + 1) As String
'                ReDim Preserve tpLinha(1 To UBound(tpLinha) + 1) As String
'                linha(UBound(linha)) = ""
'                tpLinha(UBound(tpLinha)) = "N"
            
                vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha(Space(Len(Trim(oRsAux.Fields("cat_nm").Value))) & " - Encerrante: " & oRsAux.Fields("ctc_fim_cont").Value, COLUNAS_IMP, esquerda) & Chr(9) & Chr$(10))
                ReDim Preserve linha(1 To UBound(linha) + 1) As String
                ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
                linha(UBound(linha)) = Space(Len(Trim(oRsAux.Fields("cat_nm").Value))) & " - Encerrante: " & oRsAux.Fields("ctc_fim_cont").Value
                tplinha(UBound(tplinha)) = "N"
                
                vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
                ReDim Preserve linha(1 To UBound(linha) + 1) As String
                ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
                linha(UBound(linha)) = ""
                tplinha(UBound(tplinha)) = "N"

                
                oRsAux.MoveNext
            Loop
            
            oRsAux.Close
        
            '*** ingressos sem uso
            Set oRsAux = New ADODB.Recordset
                
            If Not clsBoletim.ingressosSemUso(oRsAux) Then
                MsgBox clsBoletim.MensagemErro, vbCritical, App.ProductName
                GoTo ImprimeBoletim_fim
            End If
    
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("Ingres. sem uso: " & oRsAux.Fields("qtdeSUso").Value, COLUNAS_IMP, esquerda) & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = "Ingres. sem uso: " & oRsAux.Fields("qtdeSUso").Value
            tplinha(UBound(tplinha)) = "N"
            
            oRsAux.Close
        Else
            oRsAux.Close
        End If
        
        
        ' ****************** Forma de Recebimento
        
        Set oRsAux = New ADODB.Recordset
                
        If Not clsBoletim.FormaPagto(oRsAux) Then
            MsgBox clsBoletim.MensagemErro, vbCritical, App.ProductName
            GoTo ImprimeBoletim_fim
        End If

        vsfBoletim.AddItem (Chr$(10))
        vsfBoletim.AddItem (LinhaDupla & Chr(9) & "CAIXA" & Chr(9))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = "CAIXA"
        tplinha(UBound(tplinha)) = "S"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = ""
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & "FORMA DE RECEBIMENTO DO DIA/FILME" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = "FORMA DE RECEBIMENTO DO DIA/FILME"
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = ""
        tplinha(UBound(tplinha)) = "N"
        
        dValorTotal = 0
        
        Do While Not oRsAux.EOF()
        
            dValorTotal = dValorTotal + Format(oRsAux.Fields("valor"), "0.00")
            
            vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha(EliminaAcentos(oRsAux.Fields("pgt_desc")), 22, esquerda, ".") & _
                                Alinha(Format(oRsAux.Fields("valor"), "#,##0.00"), 10, Direita, " ") & Chr(9) & Chr$(10))
            ReDim Preserve linha(1 To UBound(linha) + 1) As String
            ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
            linha(UBound(linha)) = Alinha(EliminaAcentos(oRsAux.Fields("pgt_desc")), 22, esquerda, ".") & _
                                   Alinha(Format(oRsAux.Fields("valor"), "#,##0.00"), 10, Direita, " ")
            tplinha(UBound(tplinha)) = "N"
            
            oRsAux.MoveNext
            
        Loop
        
        oRsAux.Close
                
        vsfBoletim.AddItem (String(COLUNAS_IMP, Chr$(196)) & Chr(9) & "" & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = String(COLUNAS_IMP, "_")
        tplinha(UBound(tplinha)) = "N"
        
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & Alinha("TOTAL GERAL", 22, esquerda, ".") & _
                            Alinha(Format(dValorTotal, "#,##0.00"), 10, Direita, " ") & Chr(9) & Chr$(10))
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = Alinha("TOTAL GERAL", 22, esquerda, ".") & _
                               Alinha(Format(dValorTotal, "#,##0.00"), 10, Direita, " ")
        tplinha(UBound(tplinha)) = "N"
                
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
        vsfBoletim.AddItem (LinhaNormal & Chr(9) & "" & Chr(9) & Chr$(10))
        vsfBoletim.AddItem (Chr$(27) & "i" & Chr(9) & "" & Chr(9) & Chr$(10))
        
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = ""
        tplinha(UBound(tplinha)) = "N"
        
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = ""
        tplinha(UBound(tplinha)) = "N"
        
        ReDim Preserve linha(1 To UBound(linha) + 1) As String
        ReDim Preserve tplinha(1 To UBound(tplinha) + 1) As String
        linha(UBound(linha)) = ""
        tplinha(UBound(tplinha)) = "N"
        
        oRsCapa.MoveNext
    Loop
     
ImprimeBoletim_fim:
    If oRsCapa.State = 1 Then oRsCapa.Close
    
ImprimeBoletim_erro:
    DoEvents
    fraAviso.Visible = False
    vsfBoletim.ZOrder 0
    DoEvents
    
End Sub

Private Function TrataTipo(ByVal Tipo As String) As String
    TrataTipo = EliminaAcentos(Replace(Replace(Tipo, "PROMOÇÃO", "PROM"), "INTEIRA", "INT"))
End Function

Private Sub cmdImprime_Click()

Dim Row             As Long
Dim fonteNormal     As StdFont
Dim fonteTitulo     As StdFont
Dim fonteSubTtitulo As StdFont
Dim tpImp           As Integer
Dim hf              As Single
Dim Y               As Single
Dim tamPag          As Single
Dim strBoletim      As String
Dim lInicio         As Long

Imprime_Boletim:

    On Error GoTo erro_ImprimeBoletim
    
    tpImp = frmTipoImp.tipoImpo
    
    DoEvents
    
    'SLS 18/01/2012
    Call abreImp(OPOSPOSPrinter1)

    
    If tpImp = 1 Then
        Set fonteNormal = New StdFont
        fonteNormal.Name = "Courier New"
        fonteNormal.Size = 10
        fonteNormal.Bold = False
        fonteNormal.Italic = False
          
        Set fonteTitulo = New StdFont
        fonteTitulo.Name = "Courier New"
        fonteTitulo.Size = 14
        fonteTitulo.Bold = True
        fonteTitulo.Italic = False
        
        Set fonteSubTtitulo = New StdFont
        fonteSubTtitulo.Name = "Courier New"
        fonteSubTtitulo.Size = 12
        fonteSubTtitulo.Bold = True
        fonteSubTtitulo.Italic = False
        
        Y = 10
    
        'Printer.PaperSize = vbPRPSA4
        Printer.ScaleMode = vbMillimeters
        Printer.Orientation = vbPRORPortrait
        tamPag = Printer.ScaleHeight
        
        For Row = 1 To UBound(linha)
            If tplinha(Row) = "T" Then
                Set Printer.Font = fonteTitulo
            ElseIf tplinha(Row) = "S" Then
                Set Printer.Font = fonteSubTtitulo
            Else
                Set Printer.Font = fonteNormal
            End If
            
            hf = Printer.TextHeight("X")
            
            Printer.CurrentX = 10
            Printer.CurrentY = Y
            Printer.Print linha(Row)
            
            Y = Y + hf
            If Y + 10 > tamPag Then
                Printer.NewPage
                Y = 10
            End If
        Next
    
        Printer.EndDoc
        
    ElseIf tpImp = 2 Then
        'Inicializa impressora
        If Not iniImpressora(OPOSPOSPrinter1) Then
            GoTo erro_ImprimeBoletim
        End If
        
        If verificaImp(OPOSPOSPrinter1) Then
            strBoletim = ""
            
            'AlinhaCentralizado
            'If Not AlinhaCentralizado(OPOSPOSPrinter1) Then
            '    GoTo erro_ImprimeBoletim
            'End If
            strBoletim = strBoletim & strAlinhaCentralizado
    
            For Row = LBound(linha) To UBound(linha)
                Select Case tplinha(Row)
                    Case "N"
                        'If Not imprimeNormal(OPOSPOSPrinter1, linha(Row) & Chr$(10)) Then
                        '    GoTo erro_ImprimeBoletim
                        'End If
                        strBoletim = strBoletim & strNormal(linha(Row) & Chr$(10))
                    Case "T"
                        'If Not imprimeTitulo(OPOSPOSPrinter1, linha(Row) & Chr$(10)) Then
                        '    GoTo erro_ImprimeBoletim
                        'End If
                        strBoletim = strBoletim & strTitulo(linha(Row) & Chr$(10))
                    Case "S"
                        'If Not imprimeDupla(OPOSPOSPrinter1, linha(Row) & Chr$(10)) Then
                        '    GoTo erro_ImprimeBoletim
                        'End If
                        strBoletim = strBoletim & strDupla(linha(Row) & Chr$(10))
                End Select
            Next
            
            'If Not imprimeNormal(OPOSPOSPrinter1, Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10)) Then
            '    GoTo erro_ImprimeBoletim
            'End If
            strBoletim = strBoletim & strNormal(Chr$(10) & Chr$(10) & Chr$(10) & Chr$(10))
            
            'If Not cortaPapel(OPOSPOSPrinter1) Then
            '    GoTo erro_ImprimeBoletim
            'End If
            strBoletim = strBoletim & strCortaPapel
            
            If Not imprime(OPOSPOSPrinter1, strBoletim) Then
                GoTo erro_ImprimeBoletim
            End If
            
            lInicio = timeGetTime
            
            Do While lInicio + CInt(pTempImp2) > timeGetTime
                DoEvents
            Loop
        Else
            GoTo erro_ImprimeBoletim
        End If
    End If

    'sls 18/01/2012
    Call fechaImp(OPOSPOSPrinter1)
    
    Exit Sub

erro_ImprimeBoletim:
    Beep
    
    If MsgBox("Impressora não está pronta! Ajuste-a e clique em SIM para tentar imprimir novamente ou NÃO para desistir.", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        GoTo Imprime_Boletim
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If fechaImp(OPOSPOSPrinter1) Then
    'End If
End Sub

Private Sub MSComm_OnComm()
    Dim EventComm
    
    EventComm = MSComm.CommEvent
    
    Select Case EventComm
        Case 2
            'Recebeu um caractere, provavelmente o retorno
            'do status da impressora
            pbErroImp = False
        Case Is > 1000
            'Erro
            pbErroImp = True
    End Select
    
End Sub
