VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmImagem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imagens para Exibição"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   9465
   Begin VB.Frame fraImagem 
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   15
      TabIndex        =   6
      Top             =   0
      Width           =   9405
      Begin VB.TextBox txtArquivoImg 
         Height          =   315
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   105
         Width           =   8250
      End
      Begin VB.CommandButton cmdProcuraImg 
         Caption         =   "..."
         Height          =   315
         Left            =   750
         TabIndex        =   7
         Top             =   105
         Width           =   315
      End
      Begin VB.Label lblArqImagem 
         AutoSize        =   -1  'True
         Caption         =   "Arquivo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   15
         TabIndex        =   9
         Top             =   105
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdInserirImagem 
      Caption         =   "Inserir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1440
      TabIndex        =   5
      Top             =   630
      Width           =   1980
   End
   Begin VB.CommandButton cmdAlterarImagem 
      Caption         =   "Alterar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3675
      TabIndex        =   4
      Top             =   630
      Width           =   1980
   End
   Begin VB.CommandButton cmdExcluirImagem 
      Caption         =   "Excluir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5925
      TabIndex        =   3
      Top             =   630
      Width           =   1980
   End
   Begin VB.CommandButton cmdCancelaImagem 
      Caption         =   "Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4800
      TabIndex        =   2
      Top             =   630
      Width           =   1980
   End
   Begin VB.CommandButton cmdOkImagem 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2550
      TabIndex        =   1
      Top             =   630
      Width           =   1980
   End
   Begin MSFlexGridLib.MSFlexGrid mfgImagem 
      Height          =   5220
      Left            =   0
      TabIndex        =   0
      Top             =   1065
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   9208
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog3 
      Left            =   225
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmImagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private gRSImagem      As New ADODB.Recordset
Private modoImagem     As String
Private Function verificaImagem() As Boolean
   Dim strAux As String
   
   verificaImagem = False
   
   If Trim(txtArquivoImg.Text) = "" Then
      MsgBox "Arquivo não pode ser vazio", vbCritical, App.ProductName
      
      txtArquivoImg.SetFocus
      
      Exit Function
   End If
   
   
   verificaImagem = True

End Function

Private Function alteraImagem() As Boolean
   Dim sMsg   As String
   Dim strSql As String
   Dim i      As Integer
   Dim strAux As String
   
   On Error GoTo TrataErro
   
   alteraImagem = False
   
   If Not verificaImagem() Then
       Exit Function
   End If
   
   i = mfgImagem.Row

   gRSImagem.Open "Select * from tb_Imagens Where Arquivo = '" & Trim(txtArquivoImg.Tag) & "'", "FILE NAME=" & App.Path & "\Cinema.udl", adOpenDynamic, adLockOptimistic, adCmdText
   
   gRSImagem.Fields("arquivo").Value = Trim(txtArquivoImg.Text)
   gRSImagem.Update
   
   mfgImagem.TextMatrix(i, 0) = Trim(txtArquivoImg.Text)
   
   gRSImagem.Close
   
   alteraImagem = True
   
   Exit Function
   
TrataErro:
   If gRSImagem.State = adStateOpen Then
       gRSImagem.Close
   End If
   
   sMsg = "Ocorreu um erro em alteraImagem." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
   Exit Function
   Resume 0
End Function

Private Function insereImagem() As Boolean
   Dim sMsg     As String
   Dim strSql   As String
   Dim strAux   As String
   Dim i        As Integer
   
   On Error GoTo TrataErro
   
   insereImagem = False
   
   If Not verificaImagem() Then
       Exit Function
   End If
   
   gRSImagem.Open "tb_imagens", "FILE NAME=" & App.Path & "\Cinema.udl", adOpenDynamic, adLockOptimistic, adCmdTableDirect
   
   
   gRSImagem.AddNew
      
   gRSImagem.Fields("arquivo").Value = Trim(txtArquivoImg.Text)
   gRSImagem.Update
   
   If mfgImagem.Rows > 2 Or mfgImagem.TextMatrix(1, 0) <> "" Then
      mfgImagem.Rows = mfgImagem.Rows + 1
      i = mfgImagem.Rows - 1
   Else
      i = 1
   End If
      
   mfgImagem.TextMatrix(i, 0) = Trim(txtArquivoImg.Text)
   
   gRSImagem.Close
   
   insereImagem = True
   
   Exit Function
   
TrataErro:
   If gRSImagem.State = adStateOpen Then
       gRSImagem.Close
   End If
   
   sMsg = "Ocorreu um erro em insereImagem." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
   Exit Function
   Resume 0
End Function

Private Sub cmdProcuraImg_Click()
Dim sMsg As String

   On Error GoTo TrataErro
   
   CommonDialog3.CancelError = True
   CommonDialog3.DialogTitle = App.ProductName
   CommonDialog3.InitDir = gsArqsVideo
   CommonDialog3.Filter = "Imagens (*.jpg;*.bmp)|*.jpg;*.bmp"
   CommonDialog3.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNShareAware Or _
                         cdlOFNHideReadOnly Or cdlOFNLongNames
                         
   CommonDialog3.ShowOpen
   
   txtArquivoImg.Text = CommonDialog3.FileTitle
   'txtArquivoImg.Text = CommonDialog3.FileName
   
   Exit Sub

TrataErro:
   If Err.Number <> cdlCancel Then
      sMsg = "Ocorreu um erro em cmdProcuraImg_Click." & vbCrLf
      sMsg = sMsg & Err.Number & " - " & Err.Description
      MsgBox sMsg, vbCritical, App.ProductName
   End If
End Sub

Private Sub Form_Load()

    iTop = ((MDIFrmAdmCine.Height - Me.Height) \ 2)
    iLeft = ((MDIFrmAdmCine.width - Me.width) \ 2)
    Me.Move iLeft, iTop
    
   mfgImagem.FixedRows = 0
   mfgImagem.FixedCols = 0
   
   mfgImagem.Rows = 1
   
   mfgImagem.Rows = 2
   mfgImagem.Cols = 1
   
   mfgImagem.FixedRows = 1
   
   mfgImagem.Row = 0
   mfgImagem.Col = 0
   mfgImagem.CellFontBold = True
   
   mfgImagem.TextMatrix(0, 0) = "Arquivo"
   mfgImagem.ColWidth(0) = 9500
   mfgImagem.ColAlignment(0) = flexAlignLeftTop
   
   fraImagem.Enabled = False
   
   cmdInserirImagem.Visible = True
   cmdAlterarImagem.Visible = True
   cmdExcluirImagem.Visible = True
   cmdOkImagem.Visible = False
   cmdCancelaImagem.Visible = False
   
   cmdInserirImagem.Enabled = True
   cmdAlterarImagem.Enabled = True
   cmdExcluirImagem.Enabled = True
   cmdOkImagem.Enabled = False
   cmdCancelaImagem.Enabled = False
   
   cmdInserirImagem.ZOrder 0
   cmdAlterarImagem.ZOrder 0
   cmdExcluirImagem.ZOrder 0
   
   Call carregaImagem
   
End Sub

Private Sub mfgImagem_Click()
   Call mfgImagem_EnterCell
End Sub

Private Sub mfgImagem_EnterCell()
   Dim i As Integer
   
   i = mfgImagem.Row
   
   If i > 0 Then
      txtArquivoImg.Text = mfgImagem.TextMatrix(i, 0)
      txtArquivoImg.Tag = mfgImagem.TextMatrix(i, 0)
   Else
      txtArquivoImg.Text = ""
   End If
End Sub

Private Sub inicializaGridImagem()
   mfgImagem.FixedRows = 0
   mfgImagem.FixedCols = 0
   
   mfgImagem.Rows = 1
   
   mfgImagem.Rows = 2
   mfgImagem.Cols = 2
   
   mfgImagem.FixedRows = 1
   
   mfgImagem.Row = 0
   mfgImagem.Col = 0
   mfgImagem.CellFontBold = True
   mfgImagem.Col = 1
   mfgImagem.CellFontBold = True
   
   mfgImagem.TextMatrix(0, 0) = "Imagem"
   mfgImagem.ColWidth(0) = 800
   mfgImagem.ColAlignment(0) = flexAlignLeftTop
   
   mfgImagem.TextMatrix(0, 1) = "Arquivo"
   mfgImagem.ColWidth(1) = 9500
   mfgImagem.ColAlignment(1) = flexAlignLeftTop
End Sub

Private Sub carregaImagem()
   Dim sMsg   As String
   Dim strSql As String
   Dim i      As Integer
   Dim strAux As String
   
   On Error GoTo TrataErro
   
   strSql = "SELECT "
   strSql = strSql & "arquivo "
   strSql = strSql & "FROM tb_imagens "
   strSql = strSql & "ORDER BY descricao"
   
   gRSImagem.Open strSql, "FILE NAME=" & App.Path & "\Cinema.udl", adOpenDynamic, adLockReadOnly, adCmdText
   
   If Not (gRSImagem.EOF And gRSImagem.BOF) Then
      mfgImagem.FixedRows = 0
      mfgImagem.Rows = 1
      
      Do While Not gRSImagem.EOF
         mfgImagem.Rows = mfgImagem.Rows + 1
         i = mfgImagem.Rows - 1
         mfgImagem.TextMatrix(i, 0) = gRSImagem.Fields("arquivo").Value
         
         gRSImagem.MoveNext
      Loop
      
      If mfgImagem.Rows = 1 Then
         mfgImagem.Rows = 2
      End If
      
      mfgImagem.FixedRows = 1
      mfgImagem.Row = 1
      Call mfgImagem_Click
   Else
      mfgImagem.FixedRows = 0
      mfgImagem.Rows = 1
      
      mfgImagem.Rows = 2
      mfgImagem.FixedRows = 1
   End If
   
   gRSImagem.Close
   
   Exit Sub
   
TrataErro:
   If gRSImagem.State = adStateOpen Then
      gRSImagem.Close
   End If
   
   sMsg = "Ocorreu um erro em carregaImagem." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
End Sub

Private Sub cmdOkImagem_Click()
   If modoImagem = "inserir" Then
      If Not insereImagem Then
         Exit Sub
      End If
   ElseIf modoImagem = "alterar" Then
      If Not alteraImagem Then
         Exit Sub
      End If
   End If
   
   cmdCancelaImagem_Click
End Sub

Private Sub cmdAlterarImagem_Click()
   Dim i As Integer
   
   i = mfgImagem.Row
   
   If mfgImagem.TextMatrix(i, 0) = "" Then
      MsgBox "Não há imagem para ser alterada", vbCritical, App.ProductName
      Exit Sub
   End If
   
   fraImagem.Enabled = True
   mfgImagem.Enabled = False
   
   cmdInserirImagem.Visible = False
   cmdAlterarImagem.Visible = False
   cmdExcluirImagem.Visible = False
   cmdOkImagem.Visible = True
   cmdCancelaImagem.Visible = True
   
   cmdInserirImagem.Enabled = False
   cmdAlterarImagem.Enabled = False
   cmdExcluirImagem.Enabled = False
   cmdOkImagem.Enabled = True
   cmdCancelaImagem.Enabled = True
   
   cmdOkImagem.ZOrder 0
   cmdCancelaImagem.ZOrder 0
   
   modoImagem = "alterar"
End Sub

Private Sub cmdCancelaImagem_Click()
   fraImagem.Enabled = False
   mfgImagem.Enabled = True
   
   cmdInserirImagem.Visible = True
   cmdAlterarImagem.Visible = True
   cmdExcluirImagem.Visible = True
   cmdOkImagem.Visible = False
   cmdCancelaImagem.Visible = False
   
   cmdInserirImagem.Enabled = True
   cmdAlterarImagem.Enabled = True
   cmdExcluirImagem.Enabled = True
   cmdOkImagem.Enabled = False
   cmdCancelaImagem.Enabled = False
   
   cmdInserirImagem.ZOrder 0
   cmdAlterarImagem.ZOrder 0
   cmdExcluirImagem.ZOrder 0

   mfgImagem_Click
End Sub

Private Sub cmdExcluirImagem_Click()
   Dim sMsg   As String
   Dim strSql As String
   Dim i      As Integer
   Dim strAux As String
   
   On Error GoTo TrataErro
   
   If mfgImagem.TextMatrix(1, 0) = "" Then
      MsgBox "Não há Imagem para ser excluída.", vbCritical, App.ProductName
      Exit Sub
   End If
   
   gRSImagem.Open "Select * from tb_imagens Where Arquivo = '" & Trim(txtArquivoImg.Tag) & "'", "FILE NAME=" & App.Path & "\Cinema.udl", adOpenDynamic, adLockOptimistic, adCmdText
        
   If gRSImagem.BOF And gRSImagem.EOF Then
        gRSImagem.Close
        Exit Sub
   End If
   
   
   If MsgBox("Confirma a exclusão desta imagem?", vbYesNo, App.ProductName) = vbNo Then
        gRSImagem.Close
        Exit Sub
   End If
   gRSImagem.Delete
   
   If mfgImagem.Rows > 2 Then
      i = mfgImagem.Row
      mfgImagem.RemoveItem i
   Else
      mfgImagem.TextMatrix(1, 0) = ""
   End If
   
   gRSImagem.Close
   
   mfgImagem_Click
   
   Exit Sub
   
TrataErro:
   If gRSImagem.State = adStateOpen Then
      gRSImagem.Close
   End If
   
   sMsg = "Ocorreu um erro em cmdExcluirImagem_Click." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
   Exit Sub
   Resume 0
End Sub

Private Sub cmdInserirImagem_Click()
  txtArquivoImg.Text = ""
   
   fraImagem.Enabled = True
   mfgImagem.Enabled = False
   
   cmdInserirImagem.Visible = False
   cmdAlterarImagem.Visible = False
   cmdExcluirImagem.Visible = False
   cmdOkImagem.Visible = True
   cmdCancelaImagem.Visible = True
   
   cmdInserirImagem.Enabled = False
   cmdAlterarImagem.Enabled = False
   cmdExcluirImagem.Enabled = False
   cmdOkImagem.Enabled = True
   cmdCancelaImagem.Enabled = True
   
   cmdOkImagem.ZOrder 0
   cmdCancelaImagem.ZOrder 0
   
   modoImagem = "inserir"
End Sub
