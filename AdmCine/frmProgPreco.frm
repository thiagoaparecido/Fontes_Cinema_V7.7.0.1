VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{234FCAFF-8D53-4DC2-9CD1-F90F4F6CB524}#32.0#0"; "Comandos.ocx"
Begin VB.Form frmProgPreco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Período Preços"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7845
   Icon            =   "frmProgPreco.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   7845
   Begin VB.Frame fraGrid 
      Caption         =   "Programações"
      Height          =   3315
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   7695
      Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid 
         Height          =   2775
         Left            =   180
         TabIndex        =   0
         Top             =   360
         Width           =   7335
         _cx             =   12938
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmProgPreco.frx":000C
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
      Height          =   1545
      Left            =   60
      TabIndex        =   1
      Top             =   3480
      Width           =   7680
      Begin VB.TextBox txt_ppr_patrocinador 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1080
         Width           =   6255
      End
      Begin VB.TextBox txt_ppr_desc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1140
         MaxLength       =   6
         TabIndex        =   5
         Top             =   660
         Width           =   6255
      End
      Begin VB.CheckBox chk_ppr_flg_promocao 
         Caption         =   "Convênios"
         Height          =   255
         Left            =   5700
         TabIndex        =   4
         Top             =   300
         Width           =   1515
      End
      Begin MSComCtl2.DTPicker dtp_ppr_dt_ini 
         Height          =   315
         Left            =   1140
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58195969
         CurrentDate     =   38483
      End
      Begin MSComCtl2.DTPicker dtp_ppr_dt_fim 
         Height          =   315
         Left            =   3480
         TabIndex        =   3
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58195969
         CurrentDate     =   38483
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Patrocinador:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1140
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   300
         TabIndex        =   11
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Término:"
         Height          =   195
         Left            =   2760
         TabIndex        =   9
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Início:"
         Height          =   195
         Left            =   600
         TabIndex        =   8
         Top             =   300
         Width           =   450
      End
   End
   Begin Comandos.cmdComandos cmdComandos 
      Align           =   2  'Align Bottom
      Height          =   765
      Left            =   0
      TabIndex        =   7
      Top             =   5100
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   1349
      EnabledAltera   =   -1  'True
      EnabledExclui   =   -1  'True
   End
End
Attribute VB_Name = "frmProgPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sOperacao As String

Private Sub chk_ppr_flg_promocao_Click()
    txt_ppr_desc.Enabled = (chk_ppr_flg_promocao.Value = vbChecked)
    txt_ppr_patrocinador.Enabled = (chk_ppr_flg_promocao.Value = vbChecked)
    If Not txt_ppr_desc.Enabled Then txt_ppr_desc.Text = ""
    If Not txt_ppr_patrocinador.Enabled Then txt_ppr_patrocinador.Text = ""
End Sub

Private Sub cmdComandos_Click(ByVal iButtonClicked As Comandos.eButtonClicked, Cancel As Boolean)
    
    Select Case iButtonClicked
    
        Case ButtonAltera
            If permiteAltExc() Then
                sOperacao = "A"
                chk_ppr_flg_promocao.Enabled = False
                dtp_ppr_dt_ini.Enabled = False
                If VSFlexGrid.RowSel > 0 And VSFlexGrid.Rows > 1 Then
                    Call CarregaControles
                    Call HabilitaManut(True)
                Else
                    Cancel = True
                End If
            Else
                Cancel = True
                MsgBox "Não é possível alterar programação. Existe um período posterior!", vbCritical, App.ProductName
            End If
            
        Case ButtonNovo
            sOperacao = "I"
            chk_ppr_flg_promocao.Enabled = True
            dtp_ppr_dt_ini.Enabled = True
            
            Call LimpaControles
            Call HabilitaManut(True)

        Case ButtonExclui
            If permiteAltExc() Then
                If VSFlexGrid.RowSel > 0 And VSFlexGrid.Rows > 1 Then
                    If MsgBox("Confirma exclusão da Programação selecionada?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
                        If Exclui() Then
                            Call VSFlexGrid.RemoveItem(VSFlexGrid.RowSel)
                            Call CarregaControles
                        End If
                    End If
                End If
            Else
                Cancel = True
                MsgBox "Não é possível excluir programação. Existe um período posterior!", vbCritical, App.ProductName
            End If
    
        Case ButtonGrava
        
            If DateAdd("m", 4, dtp_ppr_dt_ini.Value) >= dtp_ppr_dt_fim.Value Then
                If Grava() Then
                    Call HabilitaManut(False)
                Else
                    Cancel = True
                End If
            Else
                MsgBox "O Período não deverá ser maior que 120 dias.", vbCritical + vbOKOnly, "Atenção"
                dtp_ppr_dt_fim.Value = DateAdd("m", 4, dtp_ppr_dt_ini.Value)
                Cancel = True
            End If
            dtp_ppr_dt_ini.Enabled = True
        Case ButtonFecha
            Unload Me
            
        Case ButtonCancela
            Call HabilitaManut(False)
            Call CarregaControles
            dtp_ppr_dt_ini.Enabled = True
            
    End Select

End Sub

Private Sub Form_Activate()
    Call CarregaControles
End Sub

Private Sub Form_Load()

    Call HabilitaManut(False)
    Call PreencheGrid
    
    Dim CineAx As New CineAux.CineAx
    Call CineAx.CentralizaTela(MDIFrmAdmCine, Me)
    
    Set CineAx = Nothing
    Call LimpaControles

    Dim log As New clsLog
    
    Set log.ConexaoADO = dbConnect
    log.usu_nm = strUsuario
    log.slg_descricao = "Entrou na tela de " & Me.Caption & " do módulo " & App.ProductName
    
    log.insereLog

End Sub
Private Sub HabilitaManut(ByVal bHabilita As Boolean)
    fraGrid.Enabled = Not bHabilita
    fraManut.Enabled = bHabilita
End Sub

Private Sub CarregaControles()
    Call LimpaControles
    If VSFlexGrid.RowSel > 0 And VSFlexGrid.Rows > 1 Then
        dtp_ppr_dt_ini.Value = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Início"))
        dtp_ppr_dt_fim.Value = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Término"))
        chk_ppr_flg_promocao.Value = IIf(VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Promoção")) = -1, 1, 0)
        txt_ppr_desc.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Descrição"))
        txt_ppr_patrocinador.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Patrocinador"))
    End If
End Sub
Private Sub LimpaControles()
    dtp_ppr_dt_ini.Value = Date
    dtp_ppr_dt_fim.Value = Date
    chk_ppr_flg_promocao.Value = vbUnchecked
    txt_ppr_desc.Text = ""
    txt_ppr_patrocinador.Text = ""
End Sub
Private Function Grava() As Boolean

    On Error GoTo Grava_Erro
    
    If Not Consiste() Then
        Exit Function
    End If
    
    Dim clsTB_PROG_PRECO As New Cine2005.clsTB_PROG_PRECO
    
    Set clsTB_PROG_PRECO.ConexaoADO = dbConnect
    
    clsTB_PROG_PRECO.ppr_dt_ini = dtp_ppr_dt_ini.Value
    clsTB_PROG_PRECO.ppr_dt_fim = dtp_ppr_dt_fim.Value
    clsTB_PROG_PRECO.ppr_flg_promocao = chk_ppr_flg_promocao.Value
    clsTB_PROG_PRECO.ppr_desc = txt_ppr_desc.Text
    clsTB_PROG_PRECO.ppr_patrocinador = txt_ppr_patrocinador.Text
    
    If sOperacao = "I" Then
        If Not clsTB_PROG_PRECO.Incluir() Then
            MsgBox "Não foi possível incluir a Programação!" & vbCrLf & clsTB_PROG_PRECO.MensagemErro, vbInformation, App.ProductName
            GoTo Grava_Fim
        End If
    Else
        clsTB_PROG_PRECO.ppr_cd = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, 0)
        If Not clsTB_PROG_PRECO.Alterar() Then
            MsgBox "Não foi possível alterar a Programação Selecionada!" & vbCrLf & clsTB_PROG_PRECO.MensagemErro, vbInformation, App.ProductName
            GoTo Grava_Fim
        End If
    End If
    
    Call PreencheGrid

    Grava = True
    GoTo Grava_Fim

Grava_Erro:
    MsgBox "Erro de execução! 'Grava/frmProgPreco'", vbCritical, App.ProductName
    
Grava_Fim:
    Set clsTB_PROG_PRECO = Nothing
    
End Function
Private Function Exclui() As Boolean

    On Error GoTo Exclui_Erro
    
    Dim clsTB_PROG_PRECO As New Cine2005.clsTB_PROG_PRECO

    Set clsTB_PROG_PRECO.ConexaoADO = dbConnect
    
    clsTB_PROG_PRECO.ppr_cd = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, 0)
    
    If Not clsTB_PROG_PRECO.Excluir() Then
        MsgBox "Não foi possível excluir a Programação Selecionada!" & vbCrLf & clsTB_PROG_PRECO.MensagemErro, vbInformation, App.ProductName
        GoTo Exclui_Fim
    End If
            
    Exclui = True
    GoTo Exclui_Fim
    
Exclui_Erro:
    MsgBox "Erro de execução! 'Exclui/frmProgPreco'", vbCritical, App.ProductName
    
Exclui_Fim:
    Set clsTB_PROG_PRECO = Nothing
    
End Function
Private Sub PreencheGrid()

    On Error GoTo PreencheGrid_Erro
    
    Dim oRs As New ADODB.Recordset
    Dim clsTB_PROG_PRECO As New Cine2005.clsTB_PROG_PRECO
    
    Set clsTB_PROG_PRECO.ConexaoADO = dbConnect
    
    If Not clsTB_PROG_PRECO.PreencheGrid(oRs) Then
        MsgBox clsTB_PROG_PRECO.MensagemErro, vbCritical, App.ProductName
        GoTo PreencheGrid_Fim
    End If

    Call CarregaGridAutomatico(VSFlexGrid, oRs, VSFlexGrid.Row)
    VSFlexGrid.ColHidden(0) = True
            
    Call CarregaControles

    GoTo PreencheGrid_Fim
    
PreencheGrid_Erro:
    MsgBox "Erro de execução! 'PreencheGrid/frmProgPreco'", vbCritical, App.ProductName
    
PreencheGrid_Fim:
    If oRs.State = 1 Then oRs.Close
    Set oRs = Nothing
    Set clsTB_PROG_PRECO = Nothing
    
End Sub
Private Function Consiste() As Boolean
    
    Dim sMens As String
    
    If dtp_ppr_dt_ini.Value > dtp_ppr_dt_fim.Value Then
        sMens = sMens & "Data Término deve ser superior a data de Início!" & vbCrLf
    End If
    
    If chk_ppr_flg_promocao.Value = vbChecked Then
        If Trim(txt_ppr_desc.Text) = "" Then
            sMens = sMens & "Descrição deve ser informada!" & vbCrLf
        End If
        If Trim(txt_ppr_patrocinador.Text) = "" Then
            sMens = sMens & "Patrocinador deve ser informado!" & vbCrLf
        End If
    End If
    
    If sMens <> "" Then
        MsgBox sMens, vbInformation, App.ProductName
        Exit Function
    End If
    
    Consiste = True
    
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim log As New clsLog
    
    Set log.ConexaoADO = dbConnect
    log.usu_nm = strUsuario
    log.slg_descricao = "Saiu da tela de " & Me.Caption & " do módulo " & App.ProductName
    
    log.insereLog
End Sub

Private Sub VSFlexGrid_RowColChange()
    Call CarregaControles
End Sub

Private Function permiteAltExc() As Boolean
    Dim dtAux  As Date
    Dim maxDtIni As Date
    Dim i        As Integer
    
    permiteAltExc = False
    
    maxDtIni = dtp_ppr_dt_ini.Value
    
    For i = 1 To VSFlexGrid.Rows - 1
        If txt_ppr_patrocinador.Text = VSFlexGrid.TextMatrix(i, VSFlexGrid.ColIndex("Patrocinador")) Then
            dtAux = VSFlexGrid.TextMatrix(i, VSFlexGrid.ColIndex("Início"))
            If maxDtIni < dtAux Then
                maxDtIni = dtAux
            End If
        End If
    Next i
    
    If dtp_ppr_dt_ini.Value >= maxDtIni Then
        permiteAltExc = True
   End If
End Function

