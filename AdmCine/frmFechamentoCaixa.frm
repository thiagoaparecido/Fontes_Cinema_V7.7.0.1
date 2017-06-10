VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C9E1AFB0-1172-11D7-83AD-0050DA238ADA}#1.0#0"; "Coptr19.ocx"
Begin VB.Form frmFechamentoCaixa 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fechamento de Caixa"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancela 
      Cancel          =   -1  'True
      Caption         =   "&Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4995
      TabIndex        =   5
      Top             =   3900
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprime 
      Caption         =   "&Imprime"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3675
      TabIndex        =   4
      Top             =   3900
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3675
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   9735
      Begin VB.CommandButton cmdMarca 
         Caption         =   "Desmarca Todos"
         Height          =   435
         Index           =   1
         Left            =   8340
         TabIndex        =   2
         Top             =   240
         Width           =   1275
      End
      Begin VB.CommandButton cmdMarca 
         Caption         =   "Marca Todos"
         Height          =   435
         Index           =   0
         Left            =   7020
         TabIndex        =   1
         Top             =   240
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker dtpAbertura 
         Height          =   315
         Left            =   1260
         TabIndex        =   0
         Top             =   300
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58851329
         CurrentDate     =   38606
      End
      Begin VSFlex7LCtl.VSFlexGrid vsfFechamento 
         Height          =   2775
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   780
         Width           =   9495
         _cx             =   16748
         _cy             =   4895
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
         BackColorBkg    =   -2147483636
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   0
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmFechamentoCaixa.frx":0000
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
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Abertura:"
         Height          =   195
         Left            =   165
         TabIndex        =   7
         Top             =   360
         Width           =   1035
      End
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   7770
      Top             =   3810
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
   Begin COPTRLib.OPOSPOSPrinter OPOSPOSPrinter1 
      Left            =   525
      Top             =   3930
      _Version        =   65536
      _ExtentX        =   609
      _ExtentY        =   582
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmFechamentoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim impriminto As Boolean

Private Sub cmdCancela_Click()
    If Not impriminto Then
        Unload Me
    End If
End Sub

Private Sub cmdImprime_Click()

    Dim Row   As Long
    Dim tpImp As Integer
    
    tpImp = frmTipoImp.tipoImpo
    
    
    'SLS 18/01/2012
    Call abreImp(OPOSPOSPrinter1)
    
    If tpImp <> 3 Then
'        If tpImp = 2 Then
'            MSComm.PortOpen = True
'        End If

        impriminto = True
        For Row = 1 To vsfFechamento.Rows - 1
        
            If vsfFechamento.TextMatrix(Row, 0) Then
                Dim iCaixa As Integer
                Dim dtAbertura As Date
                
                iCaixa = vsfFechamento.TextMatrix(Row, vsfFechamento.ColIndex("Caixa"))
                dtAbertura = vsfFechamento.TextMatrix(Row, vsfFechamento.ColIndex("Data Abertura"))
                
                Call ImprimeFechamentoCaixa("A", OPOSPOSPrinter1, pbErroImp, iCaixa, tpImp, dtAbertura)
            End If
        Next
        
'        If tpImp = 2 Then
'            MSComm.PortOpen = False
'        End If
        
        'sls 18/01/2012
        Call fechaImp(OPOSPOSPrinter1)
        
        impriminto = False
    End If
End Sub

Private Sub cmdMarca_Click(Index As Integer)

    Dim Row As Long
    
    If Not impriminto Then
        For Row = 1 To vsfFechamento.Rows - 1
            If Index = 0 Then
                vsfFechamento.TextMatrix(Row, 0) = -1
            Else
                vsfFechamento.TextMatrix(Row, 0) = 0
            End If
        Next
        
        vsfFechamento.SetFocus
    End If
End Sub

Private Sub dtpAbertura_Change()
    Call Carrega
End Sub

Private Sub Form_Load()
    dtpAbertura.Value = Now
    Call Carrega
    'Call AjustaCOM(MSComm)
    'SLS 18/01/2011
    'If abreImp(OPOSPOSPrinter1) Then
    'End If
    
    impriminto = False
    
    Dim log As New clsLog
    
    Set log.ConexaoADO = dbConnect
    log.usu_nm = strUsuario
    log.slg_descricao = "Entrou na tela de " & Me.Caption & " do módulo " & App.ProductName
    
    log.insereLog

End Sub

Private Sub Carrega()

    Dim oRs As New ADODB.Recordset
    Dim clsFechamento As New clsFechamento
    
    Set clsFechamento.ConexaoADO = dbConnect
            
    clsFechamento.cxp_dt_abertura = dtpAbertura.Value
    
    If Not clsFechamento.CaixaSelecao(oRs) Then
        MsgBox clsFechamento.MensagemErro, vbCritical, App.ProductName
        Exit Sub
    End If
    
    Call CarregaGridAutomatico(vsfFechamento, oRs)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim log As New clsLog
    
    Set log.ConexaoADO = dbConnect
    log.usu_nm = strUsuario
    log.slg_descricao = "Saiu da tela de " & Me.Caption & " do módulo " & App.ProductName
    
    log.insereLog
    
    'If fechaImp(OPOSPOSPrinter1) Then
    'End If
End Sub

Private Sub vsfFechamento_DblClick()
   If vsfFechamento.RowSel > 0 And vsfFechamento.Rows > 1 Then
      vsfFechamento.TextMatrix(vsfFechamento.RowSel, 0) = Not vsfFechamento.TextMatrix(vsfFechamento.RowSel, 0)
   End If
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

