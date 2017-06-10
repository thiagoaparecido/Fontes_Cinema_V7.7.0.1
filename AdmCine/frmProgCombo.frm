VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{234FCAFF-8D53-4DC2-9CD1-F90F4F6CB524}#32.0#0"; "Comandos.ocx"
Object = "{1753B334-0119-4B34-9134-D2B3CD181550}#36.0#0"; "Combo.ocx"
Object = "{535107C7-6800-4F8A-AAE3-2BAD0FD6B0BC}#22.0#0"; "Float2.ocx"
Begin VB.Form frmProgCombo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Período Combos"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6660
   Icon            =   "frmProgCombo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   6660
   Begin VB.Frame fraManut 
      Caption         =   "Manutenção"
      Enabled         =   0   'False
      Height          =   1185
      Left            =   120
      TabIndex        =   7
      Top             =   3540
      Width           =   6420
      Begin Float2.Float_2 flt_pcb_valor 
         Height          =   315
         Left            =   5400
         TabIndex        =   4
         Top             =   660
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         Text            =   "0,00"
      End
      Begin MSComCtl2.DTPicker dtp_pcb_dt_ini 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58195969
         CurrentDate     =   38483
      End
      Begin MSComCtl2.DTPicker dtp_pcb_dt_fim 
         Height          =   315
         Left            =   3240
         TabIndex        =   3
         Top             =   660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58195969
         CurrentDate     =   38483
      End
      Begin Combo.cboCodDesc ccd_cbo_cd 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   556
         NomeTabela      =   "tb_combo"
         NomeCampoCodigo =   "cbo_cd"
         NomeCampoDescricao=   "cbo_nm"
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
         Filtro          =   "cbo_dt_des is null"
         MostraBotaoAtualiza=   0   'False
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Combo:"
         Height          =   195
         Left            =   300
         TabIndex        =   11
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Início:"
         Height          =   195
         Left            =   420
         TabIndex        =   10
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Término:"
         Height          =   195
         Left            =   2520
         TabIndex        =   9
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         Height          =   195
         Left            =   4920
         TabIndex        =   8
         Top             =   720
         Width           =   405
      End
   End
   Begin VB.Frame fraGrid 
      Caption         =   "Programações"
      Height          =   3315
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6435
      Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid 
         Height          =   2775
         Left            =   180
         TabIndex        =   0
         Top             =   360
         Width           =   6075
         _cx             =   10716
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
         FormatString    =   $"frmProgCombo.frx":000C
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
   Begin Comandos.cmdComandos cmdComandos 
      Align           =   2  'Align Bottom
      Height          =   765
      Left            =   0
      TabIndex        =   5
      Top             =   4800
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   1349
      EnabledAltera   =   -1  'True
      EnabledExclui   =   -1  'True
   End
End
Attribute VB_Name = "frmProgCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sOperacao As String

Private Sub cmdComandos_Click(ByVal iButtonClicked As Comandos.eButtonClicked, Cancel As Boolean)
    
    ccd_cbo_cd.Enabled = True
    dtp_pcb_dt_ini.Enabled = True
            
    Select Case iButtonClicked
    
        Case ButtonAltera
            If permiteAltExc() Then
                sOperacao = "A"
                
                If VSFlexGrid.RowSel > 0 And VSFlexGrid.Rows > 1 Then
                    ccd_cbo_cd.Enabled = False
                    dtp_pcb_dt_ini.Enabled = False
                    Call CarregaControles
                    Call HabilitaManut(True)
                Else
                    Cancel = True
                End If
            Else
                Cancel = True
                MsgBox "Não é possível alterar Prog. Combo. Existe um período posterior!", vbCritical, App.ProductName
            End If
            
        Case ButtonNovo
            sOperacao = "I"
            
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
                MsgBox "Não é possível excluir Prog. Combo. Existe um período posterior!", vbCritical, App.ProductName
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

Private Sub Form_Activate()
    Call CarregaControles
End Sub

Private Sub Form_Load()

    Set ccd_cbo_cd.ConexaoADO = dbConnect
    
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
        ccd_cbo_cd.codigo = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("cbo_cd"))
        dtp_pcb_dt_ini.Value = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Início"))
        dtp_pcb_dt_fim.Value = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Término"))
        flt_pcb_valor.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Valor"))
    End If
End Sub
Private Sub LimpaControles()
    ccd_cbo_cd.codigo = ""
    dtp_pcb_dt_ini.Value = Date
    dtp_pcb_dt_fim.Value = Date
    flt_pcb_valor.Text = "0"
End Sub
Private Function Grava() As Boolean

    On Error GoTo Grava_Erro
    
    If Not Consiste() Then
        Exit Function
    End If
    
    Dim clsTB_PROG_COMBO As New Cine2005.clsTB_PROG_COMBO
    
    Set clsTB_PROG_COMBO.ConexaoADO = dbConnect
    
    clsTB_PROG_COMBO.cbo_cd = ccd_cbo_cd.codigo
    clsTB_PROG_COMBO.pcb_dt_ini = dtp_pcb_dt_ini.Value
    clsTB_PROG_COMBO.pcb_dt_fim = dtp_pcb_dt_fim.Value
    clsTB_PROG_COMBO.pcb_valor = flt_pcb_valor.Text
    
    If sOperacao = "I" Then
        If Not clsTB_PROG_COMBO.Incluir() Then
            'MsgBox "Não foi possível incluir a Programação!" & vbCrLf & clsTB_PROG_COMBO.MensagemErro, vbInformation, App.ProductName
            MsgBox "Não foi possível incluir a Programação!", vbInformation, App.ProductName
            GoTo Grava_Fim
        End If
    Else
        If Not clsTB_PROG_COMBO.Alterar() Then
            'MsgBox "Não foi possível alterar a Programação Selecionada!" & vbCrLf & clsTB_PROG_COMBO.MensagemErro, vbInformation, App.ProductName
            MsgBox "Não foi possível alterar a Programação Selecionada!", vbInformation, App.ProductName
            GoTo Grava_Fim
        End If
    End If
    
    Call PreencheGrid

    Grava = True
    GoTo Grava_Fim

Grava_Erro:
    MsgBox "Erro de execução! 'Grava/frmProgPreco'", vbCritical, App.ProductName
    
Grava_Fim:
    Set clsTB_PROG_COMBO = Nothing
    
End Function
Private Function Exclui() As Boolean

    On Error GoTo Exclui_Erro
    
    Dim clsTB_PROG_COMBO As New Cine2005.clsTB_PROG_COMBO

    Set clsTB_PROG_COMBO.ConexaoADO = dbConnect
    
    clsTB_PROG_COMBO.cbo_cd = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("cbo_cd"))
    clsTB_PROG_COMBO.pcb_dt_ini = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Início"))
    
    If Not clsTB_PROG_COMBO.Excluir() Then
        MsgBox "Não foi possível excluir a Programação Selecionada!" & vbCrLf & clsTB_PROG_COMBO.MensagemErro, vbInformation, App.ProductName
        GoTo Exclui_Fim
    End If
            
    Exclui = True
    GoTo Exclui_Fim
    
Exclui_Erro:
    MsgBox "Erro de execução! 'Exclui/frmProgPreco'", vbCritical, App.ProductName
    
Exclui_Fim:
    Set clsTB_PROG_COMBO = Nothing
    
End Function
Private Sub PreencheGrid()

    On Error GoTo PreencheGrid_Erro
    
    Dim oRs As New ADODB.Recordset
    Dim clsTB_PROG_COMBO As New Cine2005.clsTB_PROG_COMBO
    
    Set clsTB_PROG_COMBO.ConexaoADO = dbConnect
    
    If Not clsTB_PROG_COMBO.PreencheGrid(oRs) Then
        MsgBox clsTB_PROG_COMBO.MensagemErro, vbCritical, App.ProductName
        GoTo PreencheGrid_Fim
    End If

    Call CarregaGridAutomatico(VSFlexGrid, oRs, VSFlexGrid.Row)
    VSFlexGrid.ColHidden(0) = True
            
    VSFlexGrid.MergeCells = flexMergeRestrictColumns
    VSFlexGrid.MergeCol(VSFlexGrid.ColIndex("Combo")) = True
            
    Call CarregaControles

    GoTo PreencheGrid_Fim
    
PreencheGrid_Erro:
    MsgBox "Erro de execução! 'PreencheGrid/frmProgPreco'", vbCritical, App.ProductName
    
PreencheGrid_Fim:
    If oRs.State = 1 Then oRs.Close
    Set oRs = Nothing
    Set clsTB_PROG_COMBO = Nothing
    
End Sub
Private Function Consiste() As Boolean
    
    Dim sMens As String
    
    If ccd_cbo_cd.codigo = "" Then
        sMens = sMens & "Combo deve ser informado!" & vbCrLf
    End If

    If dtp_pcb_dt_fim.Value < dtp_pcb_dt_ini.Value Then
        sMens = sMens & "Data Término deve ser superior ou igual a data de Início!" & vbCrLf
    End If
    
    If Val(Replace(flt_pcb_valor.Text, ",", ".")) = 0 Then
        sMens = sMens & "Valor deve ser informado!" & vbCrLf
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
    
    maxDtIni = dtp_pcb_dt_ini.Value
    
    For i = 1 To VSFlexGrid.Rows - 1
        If ccd_cbo_cd.codigo = VSFlexGrid.TextMatrix(i, VSFlexGrid.ColIndex("cbo_cd")) Then
            dtAux = VSFlexGrid.TextMatrix(i, VSFlexGrid.ColIndex("Início"))
            If maxDtIni < dtAux Then
                maxDtIni = dtAux
            End If
        End If
    Next i
    
    If dtp_pcb_dt_ini.Value >= maxDtIni Then
        permiteAltExc = True
   End If
End Function

