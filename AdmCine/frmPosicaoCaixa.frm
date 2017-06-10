VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmPosicaoCaixa 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Posição Caixas"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
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
      Left            =   4991
      TabIndex        =   3
      Top             =   3270
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   435
      Top             =   3360
   End
   Begin VB.CommandButton cmdAtualiza 
      Caption         =   "&Atualiza"
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
      Left            =   3678
      TabIndex        =   1
      Top             =   3270
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3105
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   9735
      Begin VSFlex7LCtl.VSFlexGrid vsfPosicao 
         Height          =   2775
         Left            =   120
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   210
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
         FormatString    =   $"frmPosicaoCaixa.frx":0000
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
   End
End
Attribute VB_Name = "frmPosicaoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAtualiza_Click()
    Call Carrega
End Sub

Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call Carrega
End Sub

Private Sub Carrega()
    Dim oRs As New ADODB.Recordset
    Dim clsCaixa As New clsCaixa
    
    Set clsCaixa.ConexaoADO = dbConnect
            
    If Not clsCaixa.posicaoCaixas(oRs) Then
        MsgBox clsCaixa.MensagemErro, vbCritical, App.ProductName
        Set clsCaixa = Nothing
        Exit Sub
    End If
    
    Call CarregaGridAutomatico(vsfPosicao, oRs)
    Set clsCaixa = Nothing
End Sub


