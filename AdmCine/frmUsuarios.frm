VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{234FCAFF-8D53-4DC2-9CD1-F90F4F6CB524}#30.0#0"; "Comandos.ocx"
Object = "{1753B334-0119-4B34-9134-D2B3CD181550}#34.0#0"; "Combo.ocx"
Begin VB.Form frmUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Usuários"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   Icon            =   "frmUsuarios.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   6615
   Begin VB.Frame fraManut 
      Caption         =   "Manutenção"
      Enabled         =   0   'False
      Height          =   2025
      Left            =   60
      TabIndex        =   8
      Top             =   3420
      Width           =   6495
      Begin VB.TextBox txt_usu_senha 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   3900
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1080
         Width           =   2280
      End
      Begin VB.TextBox txt_usu_login 
         Height          =   315
         Left            =   780
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1080
         Width           =   2340
      End
      Begin VB.TextBox txt_usu_nm 
         Height          =   315
         Left            =   780
         MaxLength       =   50
         TabIndex        =   2
         Top             =   675
         Width           =   5400
      End
      Begin Combo.cboCodDesc ccd_cin_cd 
         Height          =   315
         Left            =   780
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
      Begin Combo.cboCodDesc ccd_per_cd 
         Height          =   315
         Left            =   780
         TabIndex        =   5
         Top             =   1500
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         NomeTabela      =   "tb_perfil_acesso"
         NomeCampoCodigo =   "per_cd"
         NomeCampoDescricao=   "per_desc"
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
         MostraBotaoAtualiza=   0   'False
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Perfil:"
         Height          =   195
         Left            =   300
         TabIndex        =   13
         Top             =   1560
         Width           =   390
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Senha:"
         Height          =   195
         Left            =   3285
         TabIndex        =   12
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Login:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1125
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cinema:"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   360
         Width           =   570
      End
   End
   Begin VB.Frame fraGrid 
      Caption         =   "Usuários Cadastrados"
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
   Begin Comandos.cmdComandos cmdComandos 
      Align           =   2  'Align Bottom
      Height          =   765
      Left            =   0
      TabIndex        =   6
      Top             =   5505
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1349
      EnabledAltera   =   -1  'True
      EnabledExclui   =   -1  'True
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sOperacao As String

Private Sub Form_Activate()
    Call CarregaControles
End Sub

Private Sub Form_Load()

    Set ccd_cin_cd.ConexaoADO = dbConnect
    Set ccd_per_cd.ConexaoADO = dbConnect
    
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
                If MsgBox("Confirma exclusão do Usuário selecionado?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
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
        txt_usu_nm.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Usuário"))
        txt_usu_login.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Login"))
        ccd_per_cd.codigo = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("per_cd"))
        ccd_per_cd.Refresh
        
'        Dim cCryptoVB As New CineCrypto.cnCrypto

'        cCryptoVB.DadoEntrada = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("usu_senha"))
'        cCryptoVB.Decrypt
'        txt_usu_senha.Text = cCryptoVB.Resultado
        txt_usu_senha.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("usu_senha"))
        
'        Set cCryptoVB = Nothing

    End If
End Sub
Private Sub LimpaControles()
    ccd_cin_cd.codigo = ""
    txt_usu_nm.Text = ""
    txt_usu_login.Text = ""
    txt_usu_senha.Text = ""
    ccd_per_cd.codigo = ""
End Sub
Private Function Grava() As Boolean

    On Error GoTo Grava_Erro
    
    If Not Consiste() Then
        Exit Function
    End If
    
    Dim clsTB_USUARIO As New Cine2005.clsTB_USUARIO
    
    Set clsTB_USUARIO.ConexaoADO = dbConnect
    
    clsTB_USUARIO.cin_cd = ccd_cin_cd.codigo
    clsTB_USUARIO.usu_nm = txt_usu_nm.Text
    clsTB_USUARIO.usu_login = txt_usu_login.Text
    clsTB_USUARIO.per_cd = ccd_per_cd.codigo

'    Dim cCryptoVB As New CineCrypto.cnCrypto

'    cCryptoVB.DadoEntrada = Trim(txt_usu_senha.Text)
'    cCryptoVB.Encrypt
'    clsTB_USUARIO.usu_senha = cCryptoVB.Resultado
    clsTB_USUARIO.usu_senha = Trim(txt_usu_senha.Text)
    
'    Set cCryptoVB = Nothing
    
    If sOperacao = "I" Then
        If Not clsTB_USUARIO.Incluir() Then
            MsgBox "Não foi possível incluir o Usuário!" & vbCrLf & clsTB_USUARIO.MensagemErro, vbInformation, App.ProductName
            GoTo Grava_Fim
        End If
    Else
        clsTB_USUARIO.usu_cd = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, 0)
        If Not clsTB_USUARIO.Alterar() Then
            MsgBox "Não foi possível alterar o Usuário Selecionado!" & vbCrLf & clsTB_USUARIO.MensagemErro, vbInformation, App.ProductName
            GoTo Grava_Fim
        End If
    End If
    
    Call PreencheGrid

    Grava = True
    GoTo Grava_Fim

Grava_Erro:
    MsgBox "Erro de execução! 'Grava/frmUsuarios'", vbCritical, App.ProductName
    
Grava_Fim:
    Set clsTB_USUARIO = Nothing
    
End Function
Private Function Exclui() As Boolean

    On Error GoTo Exclui_Erro
    
    Dim clsTB_USUARIO As New Cine2005.clsTB_USUARIO

    Set clsTB_USUARIO.ConexaoADO = dbConnect
    
    clsTB_USUARIO.usu_cd = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, 0)
    
    If Not clsTB_USUARIO.Excluir() Then
        MsgBox "Não foi possível excluir o Usuário Selecionado!" & vbCrLf & clsTB_USUARIO.MensagemErro, vbInformation, App.ProductName
        GoTo Exclui_Fim
    End If
            
    Exclui = True
    GoTo Exclui_Fim
    
Exclui_Erro:
    MsgBox "Erro de execução! 'Exclui/frmUsuarios'", vbCritical, App.ProductName
    
Exclui_Fim:
    Set clsTB_USUARIO = Nothing
    
End Function
Private Sub PreencheGrid()

    On Error GoTo PreencheGrid_Erro
    
    Dim oRs As New ADODB.Recordset
    Dim clsTB_USUARIO As New Cine2005.clsTB_USUARIO
    
    Set clsTB_USUARIO.ConexaoADO = dbConnect
    
    If Not clsTB_USUARIO.PreencheGrid(oRs) Then
        MsgBox clsTB_USUARIO.MensagemErro, vbCritical, App.ProductName
        GoTo PreencheGrid_Fim
    End If

    Call CarregaGridAutomatico(VSFlexGrid, oRs, VSFlexGrid.Row)
    
    VSFlexGrid.MergeCol(VSFlexGrid.ColIndex("cin_nm")) = True
    VSFlexGrid.MergeCells = flexMergeRestrictColumns
    
    VSFlexGrid.ColHidden(VSFlexGrid.ColIndex("usu_cd")) = True
    VSFlexGrid.ColHidden(VSFlexGrid.ColIndex("usu_senha")) = True
    VSFlexGrid.ColHidden(VSFlexGrid.ColIndex("cin_cd")) = True
    VSFlexGrid.ColHidden(VSFlexGrid.ColIndex("per_cd")) = True
            
    Call CarregaControles

    GoTo PreencheGrid_Fim
    
PreencheGrid_Erro:
    MsgBox "Erro de execução! 'PreencheGrid/frmUsuarios'", vbCritical, App.ProductName
    
PreencheGrid_Fim:
    If oRs.State = 1 Then oRs.Close
    Set oRs = Nothing
    Set clsTB_USUARIO = Nothing
    
End Sub
Private Function Consiste() As Boolean
    
    Dim sMens As String
    
    If ccd_cin_cd.codigo = "" Then
        sMens = sMens & "Cinema deve ser informado!" & vbCrLf
    End If
    
    If ccd_per_cd.codigo = "" Then
        sMens = sMens & "Perfil deve ser informado!" & vbCrLf
    End If
    
    If Trim(txt_usu_nm.Text) = "" Then
        sMens = sMens & "Nome do Usuário deve ser informado!" & vbCrLf
    End If
    
    If Trim(txt_usu_login.Text) = "" Then
        sMens = sMens & "Login do Usuário deve ser informado!" & vbCrLf
    End If
    
    If sMens <> "" Then
        MsgBox sMens, vbInformation, App.ProductName
        Exit Function
    End If
    
    Consiste = True
    
End Function

Private Sub txt_usu_login_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub VSFlexGrid_RowColChange()
    Call CarregaControles
End Sub


