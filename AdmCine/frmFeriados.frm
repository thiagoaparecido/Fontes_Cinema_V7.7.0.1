VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{234FCAFF-8D53-4DC2-9CD1-F90F4F6CB524}#3.0#0"; "Comandos.ocx"
Begin VB.Form frmFeriados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Feriados"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5205
   Icon            =   "frmFeriados.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   5205
   Begin VB.Frame fraManut 
      Caption         =   "Manutenção"
      Enabled         =   0   'False
      Height          =   1125
      Left            =   60
      TabIndex        =   5
      Top             =   3480
      Width           =   5040
      Begin VB.TextBox txt_fer_desc 
         Height          =   315
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   2
         Top             =   615
         Width           =   3780
      End
      Begin MSComCtl2.DTPicker dtp_fer_data 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   53936129
         CurrentDate     =   38483
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   660
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         Height          =   195
         Left            =   600
         TabIndex        =   6
         Top             =   300
         Width           =   390
      End
   End
   Begin VB.Frame fraGrid 
      Caption         =   "Feriados Cadastrados"
      Height          =   3315
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   5055
      Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid 
         Height          =   2775
         Left            =   180
         TabIndex        =   0
         Top             =   360
         Width           =   4695
         _cx             =   8281
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
      TabIndex        =   3
      Top             =   4695
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   1349
      EnabledAltera   =   -1  'True
      EnabledExclui   =   -1  'True
   End
End
Attribute VB_Name = "frmFeriados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sOperacao As String

Private Sub cmdComandos_Click(ByVal iButtonClicked As Comandos.eButtonClicked, Cancel As Boolean)
    
    Select Case iButtonClicked
    
        Case ButtonAltera
            sOperacao = "A"
            dtp_fer_data.Enabled = False
        
            If VSFlexGrid.RowSel > 0 And VSFlexGrid.Rows > 1 Then
                Call CarregaControles
                Call HabilitaManut(True)
            Else
                Cancel = True
            End If
            
        Case ButtonNovo
            sOperacao = "I"
            dtp_fer_data.Enabled = True
            
            Call LimpaControles
            Call HabilitaManut(True)

        Case ButtonExclui
            If VSFlexGrid.RowSel > 0 And VSFlexGrid.Rows > 1 Then
                If MsgBox("Confirma exclusão do Feriado selecionado?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
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
Private Sub Form_Load()

    Call HabilitaManut(False)
    Call PreencheGrid
    
    Dim CineAx As New cineAux.CineAx
    Call CineAx.CentralizaTela(MDIFrmAdmCine, Me)
    
    Set CineAx = Nothing
    
End Sub
Private Sub HabilitaManut(ByVal bHabilita As Boolean)
    fraGrid.Enabled = Not bHabilita
    fraManut.Enabled = bHabilita
End Sub

Private Sub CarregaControles()
    Call LimpaControles
    If VSFlexGrid.RowSel > 0 And VSFlexGrid.Rows > 1 Then
        dtp_fer_data.Value = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Data"))
        txt_fer_desc.Text = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, VSFlexGrid.ColIndex("Feriado"))
    End If
End Sub
Private Sub LimpaControles()
    txt_fer_desc.Text = ""
    dtp_fer_data.Value = Date
End Sub
Private Function Grava() As Boolean

    On Error GoTo Grava_Erro
    
    If Not Consiste() Then
        Exit Function
    End If
    
    Dim clsTB_Feriado As New Cine2005.clsTB_Feriado
    
    Set clsTB_Feriado.ConexaoADO = dbConnect
    
    clsTB_Feriado.fer_desc = txt_fer_desc.Text
    clsTB_Feriado.fer_data = dtp_fer_data.Value
    
    If sOperacao = "I" Then
        If Not clsTB_Feriado.Incluir() Then
            MsgBox "Não foi possível incluir o Feriado!" & vbCrLf & clsTB_Feriado.MensagemErro, vbInformation, App.ProductName
            GoTo Grava_Fim
        End If
    Else
        If Not clsTB_Feriado.Alterar() Then
            MsgBox "Não foi possível alterar o Feriado Selecionado!" & vbCrLf & clsTB_Feriado.MensagemErro, vbInformation, App.ProductName
            GoTo Grava_Fim
        End If
    End If
    
    Call PreencheGrid

    Grava = True
    GoTo Grava_Fim

Grava_Erro:
    MsgBox "Erro de execução! 'Grava/frmFeriados'", vbCritical, App.ProductName
    
Grava_Fim:
    Set clsTB_Feriado = Nothing
    
End Function
Private Function Exclui() As Boolean

    On Error GoTo Exclui_Erro
    
    Dim clsTB_Feriado As New Cine2005.clsTB_Feriado

    Set clsTB_Feriado.ConexaoADO = dbConnect
    
    clsTB_Feriado.fer_data = VSFlexGrid.TextMatrix(VSFlexGrid.RowSel, 0)
    
    If Not clsTB_Feriado.Excluir() Then
        MsgBox "Não foi possível excluir o Feriado Selecionado!" & vbCrLf & clsTB_Feriado.MensagemErro, vbInformation, App.ProductName
        GoTo Exclui_Fim
    End If
            
    Exclui = True
    GoTo Exclui_Fim
    
Exclui_Erro:
    MsgBox "Erro de execução! 'Exclui/frmFeriados'", vbCritical, App.ProductName
    
Exclui_Fim:
    Set clsTB_Feriado = Nothing
    
End Function
Private Sub PreencheGrid()

    On Error GoTo PreencheGrid_Erro
    
    Dim oRs As New ADODB.Recordset
    Dim clsTB_Feriado As New Cine2005.clsTB_Feriado
    
    Set clsTB_Feriado.ConexaoADO = dbConnect
    
    If Not clsTB_Feriado.PreencheGrid(oRs) Then
        MsgBox clsTB_Feriado.MensagemErro, vbCritical, App.ProductName
        GoTo PreencheGrid_Fim
    End If

    Call CarregaGridAutomatico(VSFlexGrid, oRs, VSFlexGrid.Row)

    Call CarregaControles

    GoTo PreencheGrid_Fim
    
PreencheGrid_Erro:
    MsgBox "Erro de execução! 'PreencheGrid/frmFeriados'", vbCritical, App.ProductName
    
PreencheGrid_Fim:
    If oRs.State = 1 Then oRs.Close
    Set oRs = Nothing
    Set clsTB_Feriado = Nothing
    
End Sub
Private Function Consiste() As Boolean
    
    Dim sMens As String
    
    If Trim(txt_fer_desc.Text) = "" Then
        sMens = sMens & "Descrição do Feriado deve ser informado!" & vbCrLf
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

