VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{234FCAFF-8D53-4DC2-9CD1-F90F4F6CB524}#32.0#0"; "Comandos.ocx"
Object = "{1753B334-0119-4B34-9134-D2B3CD181550}#36.0#0"; "Combo.ocx"
Begin VB.Form frmCaixas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Caixas"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   Icon            =   "frmCaixa.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6615
   Begin VB.Frame fraGrid 
      Caption         =   "Caixas Cadastrados"
      Height          =   3315
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   6495
      Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid 
         Height          =   2775
         Left            =   180
         TabIndex        =   0
         Top             =   360
         Width           =   6135
         _cx             =   10821
         _cy             =   4895
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
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
   Begin VB.Frame fraManut 
      Caption         =   "Manutenção"
      Enabled         =   0   'False
      Height          =   1185
      Left            =   60
      TabIndex        =   4
      Top             =   3420
      Width           =   6495
      Begin VB.TextBox txt_cxa_desc 
         Height          =   315
         Left            =   945
         MaxLength       =   12
         TabIndex        =   2
         Top             =   675
         Width           =   5400
      End
      Begin Combo.cboCodDesc ccd_cin_cd 
         Height          =   315
         Left            =   945
         TabIndex        =   1
         Top             =   300
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         NomeTabela      =   "tb_cinema"
         NomeCampoCodigo =   "cin_cd"
         NomeCampoDescricao=   "cin_nm"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoCampoCodigo =   2
         MostraBotaoNovo =   0   'False
         CodigoVisible   =   0   'False
         Filtro          =   "cin_dt_des is null"
         MostraBotaoAtualiza=   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cinema:"
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   720
         Width           =   765
      End
   End
   Begin Comandos.cmdComandos cmdComandos 
      Align           =   2  'Align Bottom
      Height          =   765
      Left            =   0
      TabIndex        =   3
      Top             =   4665
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1349
      EnabledAltera   =   -1  'True
      EnabledExclui   =   -1  'True
   End
End
Attribute VB_Name = "frmCaixas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sOperacao As String

Private Sub Form_Load()

    Set ccd_cin_cd.ConexaoADO = dbConnect
    
    Call HabilitaManut(False)
    Call PreencheGrid
    
    Dim CineAx As New CineAux.CineAx
    Call CineAx.CentralizaTela(MDIFrmAdmCine, Me)
    
    Set CineAx = Nothing
    
End Sub
Private Sub cmdComandos_Click(ByVal iButtonClicked As Comandos.eButtonClicked, Cancel As Boolean)
    
    Select Case iButtonClicked
    
        Case ButtonAltera
            sOperacao = "A"
        
            If VSFlexGrid.RowSel > 0 And VSFlexGrid.Rows > 1 Then
                Call CarregaControles
                Call HabilitaManut(True)
            Else
                Cancel = True
            End If
            
        Case ButtonNovo
            sOperacao = "I"
            Call LimpaControles
            Call HabilitaManut(True)

        Case ButtonExclui
            If VSFlexGrid.RowSel > 0 And VSFlexGrid.Rows > 1 Then
                If MsgBox("Confirma exclusão do Caixa selecionado?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
                    If Exclui() Then
                        Call VSFlexGrid.RemoveItem(VSFlexGrid.RowSel)
                        Call CarregaControles
                    End If
                End If
            End If
    
        Case ButtonGrava
            If Grava() Then
                Call HabilitaManut(False)
            Else
                Cancel = True
            End If
    
        Case ButtonFecha
            Unload Me
            
        Case ButtonCancela
            Call HabilitaManut(False)
            Call CarregaControles
            
    End Select

End Sub
Private Sub HabilitaManut(ByVal bHabilita As Boolean)
    fraGrid.Enabled = Not bHabilita
    fraManut.Enabled = bHabilita
End Sub

Private Sub CarregaControles()
    Call LimpaControles
    If VSFlexGrid.RowSel > 0 And VSFlexGrid.Rows > 1 Then
        ccd_cin_cd.codigo = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("cin_cd"))
        ccd_cin_cd.Refresh
        txt_cxa_desc.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Caixa"))
    End If
End Sub
Private Sub LimpaControles()
    ccd_cin_cd.codigo = ""
    txt_cxa_desc.Text = ""
End Sub
Private Function Grava() As Boolean

    On Error GoTo Grava_Erro
    
    If Not Consiste() Then
        Exit Function
    End If
    
    Dim clsTB_CAIXA As New Cine2005.clsTB_CAIXA
    
    Set clsTB_CAIXA.ConexaoADO = dbConnect
    
    clsTB_CAIXA.cin_cd = ccd_cin_cd.codigo
    clsTB_CAIXA.cxa_desc = txt_cxa_desc.Text
    
    If sOperacao = "I" Then
        If Not clsTB_CAIXA.Incluir() Then
            MsgBox "Não foi possível incluir o Caixa!" & vbCrLf & clsTB_CAIXA.MensagemErro, vbInformation, App.ProductName
            GoTo Grava_Fim
        End If
    Else
        clsTB_CAIXA.cxa_cd = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("cxa_cd"))
        If Not clsTB_CAIXA.Alterar() Then
            MsgBox "Não foi possível alterar o Caixa Selecionado!" & vbCrLf & clsTB_CAIXA.MensagemErro, vbInformation, App.ProductName
            GoTo Grava_Fim
        End If
    End If
    
    Call PreencheGrid

    Grava = True
    GoTo Grava_Fim

Grava_Erro:
    MsgBox "Erro de execução! 'Grava/frmCaixas'", vbCritical, App.ProductName
    
Grava_Fim:
    Set clsTB_CAIXA = Nothing
    
End Function
Private Function Exclui() As Boolean

    On Error GoTo Exclui_Erro
    
    Dim clsTB_CAIXA As New Cine2005.clsTB_CAIXA

    Set clsTB_CAIXA.ConexaoADO = dbConnect
    
    clsTB_CAIXA.cxa_cd = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("cxa_cd"))
    
    If Not clsTB_CAIXA.Excluir() Then
        MsgBox "Não foi possível excluir o Caixa Selecionado!" & vbCrLf & clsTB_CAIXA.MensagemErro, vbInformation, App.ProductName
        GoTo Exclui_Fim
    End If
            
    Exclui = True
    GoTo Exclui_Fim
    
Exclui_Erro:
    MsgBox "Erro de execução! 'Exclui/frmCaixas'", vbCritical, App.ProductName
    
Exclui_Fim:
    Set clsTB_CAIXA = Nothing
    
End Function
Private Sub PreencheGrid()

    On Error GoTo PreencheGrid_Erro
    
    Dim oRs As New ADODB.Recordset
    Dim clsTB_CAIXA As New Cine2005.clsTB_CAIXA
    
    Set clsTB_CAIXA.ConexaoADO = dbConnect
    
    If Not clsTB_CAIXA.PreencheGrid(oRs) Then
        MsgBox clsTB_CAIXA.MensagemErro, vbCritical, App.ProductName
        GoTo PreencheGrid_Fim
    End If

    Call CarregaGridAutomatico(VSFlexGrid, oRs, VSFlexGrid.Row)
    
    VSFlexGrid.ColHidden(VSFlexGrid.ColIndex("cxa_cd")) = True
    VSFlexGrid.ColHidden(VSFlexGrid.ColIndex("cin_cd")) = True
            
    VSFlexGrid.MergeCells = flexMergeRestrictColumns
    VSFlexGrid.MergeCol(VSFlexGrid.ColIndex("cxa_cd")) = True
    VSFlexGrid.MergeCol(VSFlexGrid.ColIndex("cin_cd")) = True
    VSFlexGrid.MergeCol(VSFlexGrid.ColIndex("cin_nm")) = True

    Call CarregaControles

    GoTo PreencheGrid_Fim
    
PreencheGrid_Erro:
    MsgBox "Erro de execução! 'PreencheGrid/frmCaixas'", vbCritical, App.ProductName
    
PreencheGrid_Fim:
    If oRs.State = 1 Then oRs.Close
    Set oRs = Nothing
    Set clsTB_CAIXA = Nothing
    
End Sub
Private Function Consiste() As Boolean
    
    Dim sMens As String
    
    If ccd_cin_cd.codigo = "" Then
        sMens = sMens & "Cinema deve ser informado!" & vbCrLf
    End If
    
    If Trim(txt_cxa_desc.Text) = "" Then
        sMens = sMens & "Nome do Caixa deve ser informado!" & vbCrLf
    End If
    
    If sMens <> "" Then
        MsgBox sMens, vbInformation, App.ProductName
        Exit Function
    End If
    
    Consiste = True
    
End Function

Private Sub VSFlexGrid_RowColChange()
    Call CarregaControles
End Sub

