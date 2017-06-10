VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{234FCAFF-8D53-4DC2-9CD1-F90F4F6CB524}#31.0#0"; "Comandos.ocx"
Object = "{1753B334-0119-4B34-9134-D2B3CD181550}#35.0#0"; "Combo.ocx"
Object = "{91C13016-45DE-491B-BAC0-37F755626532}#30.0#0"; "Spin.ocx"
Begin VB.Form frmLugaresSala 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro -Alteração de Lugares"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6660
   Icon            =   "frmLugaresSala.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   6660
   Begin VB.Frame fraGrid 
      Caption         =   "Programações"
      Height          =   3315
      Left            =   120
      TabIndex        =   12
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
         FormatString    =   $"frmLugaresSala.frx":000C
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
      Height          =   1605
      Left            =   120
      TabIndex        =   7
      Top             =   3540
      Width           =   6435
      Begin VB.TextBox txt_sal_mot_alt 
         Height          =   315
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1080
         Width           =   4755
      End
      Begin MSComCtl2.DTPicker dtp_sal_dt_ini 
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Top             =   660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58785793
         CurrentDate     =   38483
      End
      Begin MSComCtl2.DTPicker dtp_sal_dt_fim 
         Height          =   315
         Left            =   3000
         TabIndex        =   3
         Top             =   660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58785793
         CurrentDate     =   38483
      End
      Begin Combo.cboCodDesc ccd_sal_cd 
         Height          =   315
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         NomeTabela      =   "tb_sala"
         NomeCampoCodigo =   "sal_cd"
         NomeCampoDescricao=   "sal_desc"
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
         Filtro          =   "sal_dt_des is null"
         MostraBotaoAtualiza=   0   'False
      End
      Begin Spin.SpinNumber spn_sal_lugares 
         Height          =   315
         Left            =   5280
         TabIndex        =   4
         Top             =   660
         Width           =   1035
         _ExtentX        =   1826
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
         Max             =   9999
         Value           =   "0"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Motivo Alteração:"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   1140
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lugares:"
         Height          =   195
         Left            =   4620
         TabIndex        =   11
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Término:"
         Height          =   195
         Left            =   2280
         TabIndex        =   10
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Início:"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sala:"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   300
         Width           =   360
      End
   End
   Begin Comandos.cmdComandos cmdComandos 
      Align           =   2  'Align Bottom
      Height          =   765
      Left            =   0
      TabIndex        =   6
      Top             =   5220
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   1349
      EnabledAltera   =   -1  'True
      EnabledExclui   =   -1  'True
   End
End
Attribute VB_Name = "frmLugaresSala"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sOperacao As String

Private Sub cmdComandos_Click(ByVal iButtonClicked As Comandos.eButtonClicked, Cancel As Boolean)
    
    ccd_sal_cd.Enabled = True
    dtp_sal_dt_ini.Enabled = True
            
    Select Case iButtonClicked
    
        Case ButtonAltera
            sOperacao = "A"
            
            If VSFlexGrid.RowSel > 0 And VSFlexGrid.Rows > 1 Then
                ccd_sal_cd.Enabled = False
                dtp_sal_dt_ini.Enabled = False
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
                If MsgBox("Confirma exclusão da Programação selecionada?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
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

Private Sub Form_Activate()
    Call CarregaControles
End Sub

Private Sub Form_Load()

    Set ccd_sal_cd.ConexaoADO = dbConnect
    
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
        ccd_sal_cd.codigo = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("sal_cd"))
        dtp_sal_dt_ini.Value = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Início"))
        dtp_sal_dt_fim.Value = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Término"))
        spn_sal_lugares.Value = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Lugares"))
        txt_sal_mot_alt.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Motivo"))
    End If
End Sub
Private Sub LimpaControles()
    ccd_sal_cd.codigo = ""
    dtp_sal_dt_ini.Value = Date
    dtp_sal_dt_fim.Value = Date
    spn_sal_lugares.Value = "0"
    txt_sal_mot_alt.Text = ""
End Sub
Private Function Grava() As Boolean

    On Error GoTo Grava_Erro
    
    If Not Consiste() Then
        Exit Function
    End If
    
    Dim clsTB_SALA_LUGAR As New Cine2005.clsTB_SALA_LUGAR
    
    Set clsTB_SALA_LUGAR.ConexaoADO = dbConnect
    
    clsTB_SALA_LUGAR.sal_cd = ccd_sal_cd.codigo
    clsTB_SALA_LUGAR.sal_dt_ini = dtp_sal_dt_ini.Value
    clsTB_SALA_LUGAR.sal_dt_fim = dtp_sal_dt_fim.Value
    clsTB_SALA_LUGAR.sal_lugares = spn_sal_lugares.Value
    clsTB_SALA_LUGAR.sal_mot_alt = txt_sal_mot_alt.Text
    clsTB_SALA_LUGAR.usu_cd = intUsuario
    
    If sOperacao = "I" Then
        If Not clsTB_SALA_LUGAR.Incluir() Then
            MsgBox "Não foi possível incluir a Programação!" & vbCrLf & clsTB_SALA_LUGAR.MensagemErro, vbInformation, App.ProductName
            GoTo Grava_Fim
        End If
    Else
        If Not clsTB_SALA_LUGAR.Alterar() Then
            MsgBox "Não foi possível alterar a Programação Selecionada!" & vbCrLf & clsTB_SALA_LUGAR.MensagemErro, vbInformation, App.ProductName
            GoTo Grava_Fim
        End If
    End If
    
    Call PreencheGrid

    Grava = True
    GoTo Grava_Fim

Grava_Erro:
    MsgBox "Erro de execução! 'Grava/frmLugaresSala'", vbCritical, App.ProductName
    
Grava_Fim:
    Set clsTB_SALA_LUGAR = Nothing
    
End Function
Private Function Exclui() As Boolean

    On Error GoTo Exclui_Erro
    
    Dim clsTB_SALA_LUGAR As New Cine2005.clsTB_SALA_LUGAR

    Set clsTB_SALA_LUGAR.ConexaoADO = dbConnect
    
    clsTB_SALA_LUGAR.sal_cd = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("sal_cd"))
    clsTB_SALA_LUGAR.sal_dt_ini = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Início"))
    
    If Not clsTB_SALA_LUGAR.Excluir() Then
        MsgBox "Não foi possível excluir a Programação Selecionada!" & vbCrLf & clsTB_SALA_LUGAR.MensagemErro, vbInformation, App.ProductName
        GoTo Exclui_Fim
    End If
            
    Exclui = True
    GoTo Exclui_Fim
    
Exclui_Erro:
    MsgBox "Erro de execução! 'Exclui/frmLugaresSala'", vbCritical, App.ProductName
    
Exclui_Fim:
    Set clsTB_SALA_LUGAR = Nothing
    
End Function
Private Sub PreencheGrid()

    On Error GoTo PreencheGrid_Erro
    
    Dim oRs As New ADODB.Recordset
    Dim clsTB_SALA_LUGAR As New Cine2005.clsTB_SALA_LUGAR
    
    Set clsTB_SALA_LUGAR.ConexaoADO = dbConnect
    
    If Not clsTB_SALA_LUGAR.PreencheGrid(oRs) Then
        MsgBox clsTB_SALA_LUGAR.MensagemErro, vbCritical, App.ProductName
        GoTo PreencheGrid_Fim
    End If

    Call CarregaGridAutomatico(VSFlexGrid, oRs, VSFlexGrid.Row)
    VSFlexGrid.ColHidden(0) = True
            
    VSFlexGrid.MergeCells = flexMergeRestrictColumns
    VSFlexGrid.MergeCol(VSFlexGrid.ColIndex("Combo")) = True
            
    Call CarregaControles

    GoTo PreencheGrid_Fim
    
PreencheGrid_Erro:
    MsgBox "Erro de execução! 'PreencheGrid/frmLugaresSala'", vbCritical, App.ProductName
    
PreencheGrid_Fim:
    If oRs.State = 1 Then oRs.Close
    Set oRs = Nothing
    Set clsTB_SALA_LUGAR = Nothing
    
End Sub
Private Function Consiste() As Boolean
    
    Dim sMens As String
    
    If ccd_sal_cd.codigo = "" Then
        sMens = sMens & "Sala deve ser informada!" & vbCrLf
    End If

    If dtp_sal_dt_fim.Value < dtp_sal_dt_ini.Value Then
        sMens = sMens & "Data Término deve ser superior ou igual a data de Início!" & vbCrLf
    End If
    
    If Val(spn_sal_lugares.Value) = 0 Then
        sMens = sMens & "Número de Lugares deve ser informado!" & vbCrLf
    End If
    
    If Trim(txt_sal_mot_alt.Text) = "" Then
        sMens = sMens & "Motivo da Alteração deve ser informado!" & vbCrLf
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




