VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{234FCAFF-8D53-4DC2-9CD1-F90F4F6CB524}#22.0#0"; "Comandos.ocx"
Object = "{1753B334-0119-4B34-9134-D2B3CD181550}#20.0#0"; "Combo.ocx"
Begin VB.Form frmCatraca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Catracas"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5385
   Icon            =   "frmCatraca.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   5385
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgSalas 
      Height          =   1020
      Left            =   255
      TabIndex        =   14
      Top             =   4680
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   1799
      _Version        =   393216
      Rows            =   7
      FixedRows       =   0
      FixedCols       =   0
      ScrollBars      =   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame fraGrid 
      Caption         =   "Catracas Cadastradas"
      Height          =   2400
      Left            =   60
      TabIndex        =   10
      Top             =   45
      Width           =   5235
      Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid 
         Height          =   1965
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Width           =   4935
         _cx             =   8705
         _cy             =   3466
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
      Height          =   4140
      Left            =   60
      TabIndex        =   8
      Top             =   2460
      Width           =   5220
      Begin VB.Frame fraNumeracao 
         Caption         =   "Numeração"
         Height          =   690
         Left            =   90
         TabIndex        =   16
         Top             =   3345
         Width           =   5025
         Begin VB.CheckBox chkAltNum 
            Caption         =   "Altera Numeração"
            Height          =   315
            Left            =   1425
            TabIndex        =   18
            Top             =   240
            Width           =   1605
         End
         Begin VB.TextBox txtNumeracao 
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   17
            Top             =   240
            Width           =   825
         End
      End
      Begin VB.TextBox txt_cat_num 
         Height          =   315
         Left            =   840
         MaxLength       =   2
         TabIndex        =   0
         Top             =   225
         Width           =   585
      End
      Begin VB.Frame frmSalas 
         Caption         =   "Salas"
         Height          =   2070
         Left            =   90
         TabIndex        =   12
         Top             =   1275
         Width           =   5025
         Begin VB.CommandButton cmdExclui 
            Caption         =   "Exclui"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2355
            TabIndex        =   5
            Top             =   570
            Width           =   1275
         End
         Begin VB.CommandButton cmdInclui 
            Caption         =   "Inclui"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   975
            TabIndex        =   4
            Top             =   570
            Width           =   1275
         End
         Begin Combo.cboCodDesc cboCodDesc1 
            Height          =   315
            Left            =   465
            TabIndex        =   3
            Top             =   195
            Width           =   4125
            _ExtentX        =   7276
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
         Begin VB.Label lblSala 
            AutoSize        =   -1  'True
            Caption         =   "Sala:"
            Height          =   195
            Left            =   60
            TabIndex        =   13
            Top             =   255
            Width           =   360
         End
      End
      Begin VB.TextBox txt_cat_nm 
         Height          =   315
         Left            =   840
         MaxLength       =   50
         TabIndex        =   2
         Top             =   960
         Width           =   4200
      End
      Begin Combo.cboCodDesc ccd_cin_cd 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   585
         Width           =   4215
         _ExtentX        =   7435
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
         Caption         =   "Número:"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cinema:"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   645
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Left            =   300
         TabIndex        =   9
         Top             =   1005
         Width           =   465
      End
   End
   Begin Comandos.cmdComandos cmdComandos 
      Align           =   2  'Align Bottom
      Height          =   765
      Left            =   0
      TabIndex        =   6
      Top             =   6600
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   1349
      EnabledAltera   =   -1  'True
      EnabledExclui   =   -1  'True
   End
End
Attribute VB_Name = "frmCatraca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sOperacao As String

Private Sub chkAltNum_Click()
    If chkAltNum.Value = vbChecked Then
        txtNumeracao.Locked = False
    Else
        txtNumeracao.Locked = True
    End If
End Sub

Private Sub cmdExclui_Click()
    If hfgSalas.Row >= 0 And hfgSalas.Rows > 0 Then
'        If hfgSalas.Rows > 1 Then
            hfgSalas.RemoveItem hfgSalas.Row
'        Else
'            hfgSalas.TextMatrix(hfgSalas.row, 0) = ""
'            hfgSalas.TextMatrix(hfgSalas.row, 1) = ""
'        End If
    End If
End Sub

Private Sub cmdInclui_Click()
    Dim i As Integer
    
    If cboCodDesc1.Descricao <> "" Then
        For i = 0 To hfgSalas.Rows - 1
            If hfgSalas.TextMatrix(i, 1) = cboCodDesc1.codigo Then
                MsgBox "Sala já associada com esta catraca!", vbInformation, App.ProductName
                Exit Sub
            End If
        Next i
        'If hfgSalas.TextMatrix(hfgSalas.row, 1) = "" Then
        '    hfgSalas.TextMatrix(hfgSalas.row, 0) = cboCodDesc1.Descricao
        '    hfgSalas.TextMatrix(hfgSalas.row, 1) = cboCodDesc1.codigo
        'Else
            hfgSalas.AddItem cboCodDesc1.Descricao & vbTab & cboCodDesc1.codigo
        'End If
    End If
End Sub

Private Sub Form_Load()

    Set ccd_cin_cd.ConexaoADO = dbConnect
    Set cboCodDesc1.ConexaoADO = dbConnect
    
    Call HabilitaManut(False)
    Call iniGridSalas
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
                If MsgBox("Confirma exclusão da Catraca selecionada?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
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
            cboCodDesc1.Clear
    End Select

End Sub

Private Sub HabilitaManut(ByVal bHabilita As Boolean)
    fraGrid.Enabled = Not bHabilita
    fraManut.Enabled = bHabilita
    fraNumeracao.Enabled = False
    
    If sOperacao = "A" Then
        txt_cat_num.Enabled = Not bHabilita
        fraNumeracao.Enabled = True
    End If
End Sub

Private Sub CarregaControles()
    Dim catraca As New clsTB_CATRACA
    Dim i As Integer
    
    Call LimpaControles
    If VSFlexGrid.RowSel > 0 And VSFlexGrid.Rows > 1 Then
        ccd_cin_cd.codigo = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("cin_cd"))
        ccd_cin_cd.Refresh
        txt_cat_nm.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Nome"))
        txt_cat_num.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("cat_cd"))
        
        catraca.cat_cd = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("cat_cd"))
        Set catraca.ConexaoADO = dbConnect
        If (catraca.CarregaSalas) Then
            For i = 1 To catraca.salas.Count
                hfgSalas.AddItem catraca.salas.Item(i).Descricao & vbTab & catraca.salas.Item(i).codigo
            Next i
        End If
        
        If (catraca.CarregaNumeracao) Then
            txtNumeracao.Text = catraca.ctc_fim_cont
        End If
    End If
End Sub

Private Sub LimpaControles()
    ccd_cin_cd.codigo = ""
    txt_cat_nm.Text = ""
    txt_cat_num.Text = ""
        
    chkAltNum.Value = vbUnchecked
    txtNumeracao.Text = ""
        
    hfgSalas.Rows = 0
End Sub

Private Function Grava() As Boolean
    Dim i     As Integer
    Dim salas As New clcSalas

    On Error GoTo Grava_Erro
    
    If Not Consiste() Then
        Exit Function
    End If
    
    Dim clsTB_CATRACA As New Cine2005.clsTB_CATRACA
    
    Set clsTB_CATRACA.ConexaoADO = dbConnect
    
    clsTB_CATRACA.cat_cd = CInt(txt_cat_num.Text)
    clsTB_CATRACA.cin_cd = ccd_cin_cd.codigo
    clsTB_CATRACA.cat_nm = txt_cat_nm.Text
    
    If sOperacao = "A" Then
        clsTB_CATRACA.AltNumeracao = (chkAltNum.Value = vbChecked)
        clsTB_CATRACA.ctc_fim_cont = CLng(txtNumeracao.Text)
    Else
        clsTB_CATRACA.AltNumeracao = False
        clsTB_CATRACA.ctc_fim_cont = 0
    End If
    'clsTB_CATRACA.cat_posicao = spn_cat_posicao.Value
    
    For i = 0 To hfgSalas.Rows - 1
        clsTB_CATRACA.salas.Add hfgSalas.TextMatrix(i, 0), hfgSalas.TextMatrix(i, 1)
    Next i
    
    'Set clsTB_CATRACA.salas = salas
    
    If sOperacao = "I" Then
        If Not clsTB_CATRACA.Incluir() Then
            MsgBox "Não foi possível incluir a Catraca!" & vbCrLf & clsTB_CATRACA.MensagemErro, vbInformation, App.ProductName
            GoTo Grava_Fim
        End If
    Else
        If Not clsTB_CATRACA.Alterar() Then
            MsgBox "Não foi possível alterar a Catraca Selecionada!" & vbCrLf & clsTB_CATRACA.MensagemErro, vbInformation, App.ProductName
            GoTo Grava_Fim
        End If
    End If
    
    Call PreencheGrid
    cboCodDesc1.Clear

    Grava = True
    GoTo Grava_Fim

Grava_Erro:
    MsgBox "Erro de execução! 'Grava/frmCatracas'", vbCritical, App.ProductName
    
Grava_Fim:
    Set clsTB_CATRACA = Nothing
    
End Function

Private Function Exclui() As Boolean

    On Error GoTo Exclui_Erro
    
    Dim clsTB_CATRACA As New Cine2005.clsTB_CATRACA

    Set clsTB_CATRACA.ConexaoADO = dbConnect
    
    clsTB_CATRACA.cat_cd = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("cat_cd"))
    
    If Not clsTB_CATRACA.Excluir() Then
        MsgBox "Não foi possível excluir a Catraca Selecionada!" & vbCrLf & clsTB_CATRACA.MensagemErro, vbInformation, App.ProductName
        GoTo Exclui_Fim
    End If
            
    Exclui = True
    GoTo Exclui_Fim
    
Exclui_Erro:
    MsgBox "Erro de execução! 'Exclui/frmCatracas'", vbCritical, App.ProductName
    
Exclui_Fim:
    Set clsTB_CATRACA = Nothing
    
End Function

Private Sub PreencheGrid()

    On Error GoTo PreencheGrid_Erro
    
    Dim oRs As New ADODB.Recordset
    Dim clsTB_CATRACA As New Cine2005.clsTB_CATRACA
    
    Set clsTB_CATRACA.ConexaoADO = dbConnect
    
    If Not clsTB_CATRACA.PreencheGrid(oRs) Then
        MsgBox clsTB_CATRACA.MensagemErro, vbCritical, App.ProductName
        GoTo PreencheGrid_Fim
    End If

    Call CarregaGridAutomatico(VSFlexGrid, oRs, VSFlexGrid.Row)
    
    VSFlexGrid.ColHidden(VSFlexGrid.ColIndex("cat_cd")) = True
    VSFlexGrid.ColHidden(VSFlexGrid.ColIndex("cin_cd")) = True

    VSFlexGrid.MergeCells = flexMergeRestrictColumns
    VSFlexGrid.MergeCol(VSFlexGrid.ColIndex("cat_cd")) = True
    VSFlexGrid.MergeCol(VSFlexGrid.ColIndex("cin_cd")) = True
    VSFlexGrid.MergeCol(VSFlexGrid.ColIndex("cin_nm")) = True
    
    Call CarregaControles

    GoTo PreencheGrid_Fim
    
PreencheGrid_Erro:
    MsgBox "Erro de execução! 'PreencheGrid/frmCatracas'", vbCritical, App.ProductName
    
PreencheGrid_Fim:
    If oRs.State = 1 Then oRs.Close
    Set oRs = Nothing
    Set clsTB_CATRACA = Nothing
    
End Sub

Private Function Consiste() As Boolean
    
    Dim sMens As String
    
    If txt_cat_num.Text = "" Then
        sMens = sMens & "Número da catraca deve ser informado!" & vbCrLf
    End If
    
    If Not IsNumeric(txt_cat_num.Text) Then
        sMens = sMens & "Número da catraca invalido!" & vbCrLf
    End If
    
    If ccd_cin_cd.codigo = "" Then
        sMens = sMens & "Cinema deve ser informado!" & vbCrLf
    End If
    
    If Trim(txt_cat_nm.Text) = "" Then
        sMens = sMens & "Nome da Catraca deve ser informado!" & vbCrLf
    End If
    
    If hfgSalas.Rows = 0 Then
        sMens = sMens & "Não existem salas associadas com esta catraca!" & vbCrLf
    End If

    If sOperacao = "A" Then
        If Not IsNumeric(txtNumeracao.Text) Then
            If chkAltNum.Value = vbChecked Then
                sMens = sMens & "Valor da numeração invalido!" & vbCrLf
            Else
                txtNumeracao.Text = 0
            End If
        End If
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

Private Sub iniGridSalas()
    hfgSalas.Cols = 2
    hfgSalas.Rows = 0
    hfgSalas.FixedCols = 0
    hfgSalas.FixedRows = 1
    hfgSalas.ColWidth(0) = 4740
    hfgSalas.ColWidth(1) = 0
End Sub
