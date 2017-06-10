VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log Sistema"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   11505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7575
      TabIndex        =   6
      Top             =   90
      Width           =   1590
   End
   Begin VB.CommandButton cmdCancela 
      Cancel          =   -1  'True
      Caption         =   "Cancela"
      Height          =   420
      Left            =   9495
      TabIndex        =   5
      Top             =   90
      Width           =   1590
   End
   Begin VSFlex7DAOCtl.VSFlexGrid vfgLog 
      Height          =   6060
      Left            =   90
      TabIndex        =   0
      Top             =   585
      Width           =   11310
      _cx             =   19950
      _cy             =   10689
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmLog.frx":0000
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
      DataMode        =   0
      VirtualData     =   -1  'True
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSComCtl2.DTPicker dtpDataDe 
      Height          =   315
      Left            =   1065
      TabIndex        =   1
      Top             =   90
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   58851329
      CurrentDate     =   38606
   End
   Begin MSComCtl2.DTPicker dtpDataAte 
      Height          =   315
      Left            =   2925
      TabIndex        =   2
      Top             =   90
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   58851329
      CurrentDate     =   38606
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ate"
      Height          =   195
      Left            =   2580
      TabIndex        =   4
      Top             =   150
      Width           =   225
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Período de:"
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   150
      Width           =   840
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim oRs   As ADODB.Recordset
    Dim clLog As New clsLog
    
    On Error GoTo cmdOK_Click_Erro
    
    vfgLog.Rows = 1
    
    If dtpDataDe.Value > dtpDataAte.Value Then
        MsgBox "Data final não pode ser anterior a data inicial!", vbInformation, App.ProductName
        Exit Sub
    End If

    Set clLog.ConexaoADO = dbConnect
    
    clLog.emp_cd = iEmp
    clLog.cin_cd = iCin
    clLog.dt_ini = dtpDataDe.Value
    clLog.dt_fim = dtpDataAte.Value
    
    Call clLog.consLog(oRs)

    Do While Not oRs.EOF
        vfgLog.Rows = vfgLog.Rows + 1
        
        vfgLog.TextMatrix(vfgLog.Rows - 1, vfgLog.ColIndex("colData")) = oRs.Fields("slg_data").Value
        vfgLog.TextMatrix(vfgLog.Rows - 1, vfgLog.ColIndex("colUsuario")) = oRs.Fields("usu_nm").Value
        vfgLog.TextMatrix(vfgLog.Rows - 1, vfgLog.ColIndex("colLog")) = oRs.Fields("slg_descricao").Value
        
        oRs.MoveNext
    Loop

    oRs.Close
    
    Set oRs = Nothing
    Set clLog = Nothing

    Exit Sub
    
cmdOK_Click_Erro:
    Set oRs = Nothing
    Set clLog = Nothing
    
    MsgBox "Erro na consulta do log! " & Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Form_Load()
    dtpDataDe.Value = Now
    dtpDataAte.Value = Now
    
    iEmp = codigoEmpresa()
    iCin = codigoCinema()
    
    vfgLog.Rows = 1

End Sub

Private Function codigoEmpresa() As Integer
    Dim oRs           As New ADODB.Recordset
    Dim clsTB_EMPRESA As New Cine2005.clsTB_EMPRESA
    
    On Error GoTo codigoEmpresa_Erro
    
    codigoEmpresa = 0
    
    Set clsTB_EMPRESA.ConexaoADO = dbConnect
    
    If Not clsTB_EMPRESA.Selecionar(oRs) Then
        Exit Function
    End If
    
    If Not oRs.EOF() Then
        codigoEmpresa = oRs.Fields("emp_cd")
    End If
    
codigoEmpresa_Erro:
    If oRs.State = 1 Then oRs.Close
    
    Set oRs = Nothing
    Set clsTB_EMPRESA = Nothing
End Function

Private Function codigoCinema() As Integer
    Dim oRs          As New ADODB.Recordset
    Dim clsTB_CINEMA As New Cine2005.clsTB_CINEMA
    
    On Error GoTo codigoCinema_Erro
    
    codigoCinema = 0
    
    Set clsTB_CINEMA.ConexaoADO = dbConnect
    
    If Not clsTB_CINEMA.Selecionar(oRs) Then
        Exit Function
    End If
    
    If Not oRs.EOF() Then
        codigoCinema = oRs.Fields("cin_cd")
    End If
    
codigoCinema_Erro:
    If oRs.State = 1 Then oRs.Close
    
    Set oRs = Nothing
    Set clsTB_CINEMA = Nothing
End Function

