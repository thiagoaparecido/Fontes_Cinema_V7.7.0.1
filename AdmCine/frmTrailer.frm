VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTrailer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trailers para Exibição"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   11400
   Begin VB.Frame fraTrailer 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   75
      TabIndex        =   6
      Top             =   150
      Width           =   11280
      Begin VB.CommandButton cmdProcura 
         Caption         =   "..."
         Height          =   315
         Left            =   825
         TabIndex        =   8
         Top             =   30
         Width           =   315
      End
      Begin VB.TextBox txtArquivo 
         Height          =   315
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   30
         Width           =   9975
      End
      Begin VB.Label lblArquivo 
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
         Left            =   90
         TabIndex        =   9
         Top             =   30
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdInserirTrailer 
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
      Left            =   2250
      TabIndex        =   4
      Top             =   705
      Width           =   1980
   End
   Begin VB.CommandButton cmdAlterarTrailer 
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
      Left            =   4485
      TabIndex        =   3
      Top             =   705
      Width           =   1980
   End
   Begin VB.CommandButton cmdExcluirTrailer 
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
      Left            =   6735
      TabIndex        =   2
      Top             =   705
      Width           =   1980
   End
   Begin VB.CommandButton cmdCancelaTrailer 
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
      Left            =   5610
      TabIndex        =   1
      Top             =   705
      Width           =   1980
   End
   Begin VB.CommandButton cmdOkTrailer 
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
      Left            =   3360
      TabIndex        =   0
      Top             =   705
      Width           =   1980
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   375
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid mfgTrailer 
      Height          =   5520
      Left            =   150
      TabIndex        =   5
      Top             =   1350
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   9737
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmTrailer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private gRSTrailer     As New ADODB.Recordset
Private modoTrailer    As String

Private Const tempoEntreSessoes = 10



Private Sub cmdProcura_Click()
   Dim sMsg As String

   On Error GoTo TrataErro
   
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = App.ProductName
   CommonDialog1.InitDir = gsArqsVideo
   CommonDialog1.Filter = "Todos os Formatos de video|*.Avi;*.mpeg;*.mpg;*.wmv;*.ASF;*.SWF;*.F4V;*.F4P;*.F4A;*.F4B;*.FLV;*.MP4;*.3GP;*.3G2;*.mov;*.QT;*.MKV; *.MKA;*.RMVB;*.RM|Videos AVI (*.avi)|*.Avi|Videos Mpeg (*.MPEG, *.MPG)|*.mpeg;*.mpg|Windows Media Movie (*.WMV)|*.wmv|Advanced Systems Format(*.ASF)|*.ASF|Flash (*.SWF, *.F4V, *.F4P, *.F4A, *.F4B, *.FLV)|*.SWF;*.F4V;*.F4P;*.F4A;*.F4B;*.FLV|MP4 (*.MP4, *.3GP,*.3G2)|*.MP4;*.3GP;*.3G2|QuickTime (*.MOV,*.QT)|*.mov;*.QT|Matroska (*.MKV, *.MKA)|*.MKV; *.MKA|Real Video (*.RMVB,*.RM)|*.RMVB;*.RM"

   CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNShareAware Or _
                         cdlOFNHideReadOnly Or cdlOFNLongNames
                         
   CommonDialog1.ShowOpen
   
   txtArquivo.Text = CommonDialog1.FileTitle
   
   Exit Sub

TrataErro:
   If Err.Number <> cdlCancel Then
      sMsg = "Ocorreu um erro em cmdProcura_Click." & vbCrLf
      sMsg = sMsg & Err.Number & " - " & Err.Description
      MsgBox sMsg, vbCritical, App.ProductName
   End If
End Sub

Private Sub Command1_Click()
    Call alteraTrailer
End Sub

Private Sub Form_Load()
    iTop = ((MDIFrmAdmCine.Height - Me.Height) \ 2)
    iLeft = ((MDIFrmAdmCine.width - Me.width) \ 2)
    Me.Move iLeft, iTop
    
    gsArqsVideo = GetSetting("CineProg", "Diretorios", "arqsVideo", "")
    If gsArqsVideo = "" Then
        gsArqsVideo = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
        SaveSetting "CineProg", "Diretorios", "arqsVideo", gsArqsVideo
    End If
   
   modoFilme = "consulta"
   modoSala = "consulta"
   modoSessao = "consulta"
   modoPreco = "consulta"
   modoTrailer = "consulta"
   modoFeriado = "consulta"
   modoParametros = "consulta"
   modoImagem = "consulta"
   
   
   fraTrailer.Enabled = False
   
   cmdInserirTrailer.Visible = True
   cmdAlterarTrailer.Visible = True
   cmdExcluirTrailer.Visible = True
   cmdOkTrailer.Visible = False
   cmdCancelaTrailer.Visible = False
   
   cmdInserirTrailer.Enabled = True
   cmdAlterarTrailer.Enabled = True
   cmdExcluirTrailer.Enabled = True
   cmdOkTrailer.Enabled = False
   cmdCancelaTrailer.Enabled = False
   
   cmdInserirTrailer.ZOrder 0
   cmdAlterarTrailer.ZOrder 0
   cmdExcluirTrailer.ZOrder 0
    
   mfgTrailer.FixedRows = 0
   mfgTrailer.FixedCols = 0
   
   mfgTrailer.Rows = 1
   
   mfgTrailer.Rows = 2
   mfgTrailer.Cols = 1
   
   mfgTrailer.FixedRows = 1
   
   mfgTrailer.Row = 0
   mfgTrailer.Col = 0
   mfgTrailer.CellFontBold = True
   mfgTrailer.Col = 0
   mfgTrailer.CellFontBold = True
   
   mfgTrailer.TextMatrix(0, 0) = "Arquivo"
   mfgTrailer.ColWidth(0) = 9500
   mfgTrailer.ColAlignment(0) = flexAlignLeftTop
     
   Call carregaTrailer

End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Call gFechaBase
End Sub

Private Sub carregaTrailer()
   Dim sMsg   As String
   Dim strSql As String
   Dim i      As Integer
   Dim strAux As String
   
   On Error GoTo TrataErro
   
   strSql = "SELECT "
   strSql = strSql & "arquivo "
   strSql = strSql & "FROM tb_trailer "
   strSql = strSql & "ORDER BY descricao"
   
   'gRSTrailer.Open strSql, gConnect, adOpenDynamic, adLockReadOnly, adCmdText
   gRSTrailer.Open strSql, "FILE NAME=" & App.Path & "\Cinema.udl", adOpenDynamic, adLockReadOnly, adCmdText
   
   
   If Not (gRSTrailer.EOF And gRSTrailer.BOF) Then
      mfgTrailer.FixedRows = 0
      mfgTrailer.Rows = 1
      
      Do While Not gRSTrailer.EOF
         mfgTrailer.Rows = mfgTrailer.Rows + 1
         i = mfgTrailer.Rows - 1
         
         'mfgTrailer.TextMatrix(i, 0) = gRSTrailer.Fields("descricao").Value
         mfgTrailer.TextMatrix(i, 0) = gRSTrailer.Fields("Arquivo").Value
         'mfgTrailer.TextMatrix(i, 1) = gRSTrailer.Fields("arquivo").Value
         
         gRSTrailer.MoveNext
      Loop
      
      If mfgTrailer.Rows = 1 Then
         mfgTrailer.Rows = 2
      End If
      
      mfgTrailer.FixedRows = 1
      mfgTrailer.Row = 1
      Call mfgTrailer_Click
   Else
      mfgTrailer.FixedRows = 0
      mfgTrailer.Rows = 1
      
      mfgTrailer.Rows = 2
      mfgTrailer.FixedRows = 1
   End If
   
   gRSTrailer.Close
   
   Exit Sub
   
TrataErro:
   If gRSTrailer.State = adStateOpen Then
      gRSTrailer.Close
   End If
   
   sMsg = "Ocorreu um erro em carregaTrailer." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
End Sub

Private Sub cmdOkTrailer_Click()
   If modoTrailer = "inserir" Then
      If Not insereTrailer Then
         Exit Sub
      End If
   ElseIf modoTrailer = "alterar" Then
      If Not alteraTrailer Then
         Exit Sub
      End If
   End If
   
   cmdCancelaTrailer_Click
End Sub

Private Sub cmdAlterarTrailer_Click()
   Dim i As Integer
   
   i = mfgTrailer.Row
   
   If mfgTrailer.TextMatrix(i, 0) = "" Then
      MsgBox "Não há trailer para ser alterado.", vbCritical, App.ProductName
      Exit Sub
   End If
   
   fraTrailer.Enabled = True
   mfgTrailer.Enabled = False
   
   cmdInserirTrailer.Visible = False
   cmdAlterarTrailer.Visible = False
   cmdExcluirTrailer.Visible = False
   cmdOkTrailer.Visible = True
   cmdCancelaTrailer.Visible = True
   
   cmdInserirTrailer.Enabled = False
   cmdAlterarTrailer.Enabled = False
   cmdExcluirTrailer.Enabled = False
   cmdOkTrailer.Enabled = True
   cmdCancelaTrailer.Enabled = True
   
   cmdOkTrailer.ZOrder 0
   cmdCancelaTrailer.ZOrder 0
   
   modoTrailer = "alterar"
End Sub

Private Sub cmdCancelaTrailer_Click()
   fraTrailer.Enabled = False
   mfgTrailer.Enabled = True
   
   cmdInserirTrailer.Visible = True
   cmdAlterarTrailer.Visible = True
   cmdExcluirTrailer.Visible = True
   cmdOkTrailer.Visible = False
   cmdCancelaTrailer.Visible = False
   
   cmdInserirTrailer.Enabled = True
   cmdAlterarTrailer.Enabled = True
   cmdExcluirTrailer.Enabled = True
   cmdOkTrailer.Enabled = False
   cmdCancelaTrailer.Enabled = False
   
   cmdInserirTrailer.ZOrder 0
   cmdAlterarTrailer.ZOrder 0
   cmdExcluirTrailer.ZOrder 0

   mfgTrailer_Click
End Sub

Private Sub cmdExcluirTrailer_Click()
   Dim sMsg   As String
   Dim strSql As String
   Dim i      As Integer
   Dim strAux As String
   
   On Error GoTo TrataErro
   
   If mfgTrailer.TextMatrix(1, 0) = "" Then
      MsgBox "Não há trailer para ser excluído.", vbCritical, App.ProductName
      Exit Sub
   End If
   
      
   gRSTrailer.Open "Select * from tb_trailer Where Arquivo = '" & Trim(txtArquivo.Tag) & "'", "FILE NAME=" & App.Path & "\Cinema.udl", adOpenDynamic, adLockOptimistic, adCmdText
        
   If gRSTrailer.BOF And gRSTrailer.EOF Then
        gRSTrailer.Close
        Exit Sub
   End If
   
   If MsgBox("Confirma a exclusão deste trailer?", vbYesNo, App.ProductName) = vbNo Then
        gRSTrailer.Close
        Exit Sub
   End If
   gRSTrailer.Delete
   
   If mfgTrailer.Rows > 2 Then
      i = mfgTrailer.Row
      mfgTrailer.RemoveItem i
   Else
      mfgTrailer.TextMatrix(1, 0) = ""
   End If
   
   gRSTrailer.Close
   
   mfgTrailer_Click
   
   Exit Sub
   
TrataErro:
   If gRSTrailer.State = adStateOpen Then
      gRSTrailer.Close
   End If
   
   sMsg = "Ocorreu um erro em cmdExcluirTrailer_Click." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
   Exit Sub
   Resume 0
End Sub

Private Sub cmdInserirTrailer_Click()
   'txtFilmeTrailer.Text = ""
   txtArquivo.Text = ""
   
   fraTrailer.Enabled = True
   mfgTrailer.Enabled = False
   
   cmdInserirTrailer.Visible = False
   cmdAlterarTrailer.Visible = False
   cmdExcluirTrailer.Visible = False
   cmdOkTrailer.Visible = True
   cmdCancelaTrailer.Visible = True
   
   cmdInserirTrailer.Enabled = False
   cmdAlterarTrailer.Enabled = False
   cmdExcluirTrailer.Enabled = False
   cmdOkTrailer.Enabled = True
   cmdCancelaTrailer.Enabled = True
   
   cmdOkTrailer.ZOrder 0
   cmdCancelaTrailer.ZOrder 0
   
   modoTrailer = "inserir"
End Sub

Private Function verificaTrailer() As Boolean
   Dim strAux As String

   verificaTrailer = False

   'If Trim(txtFilmeTrailer.Text) = "" Then
   '   MsgBox "Filme não pode ser vazio", vbCritical, App.ProductName''

'      txtFilmeTrailer.SetFocus

      'Exit Function
   'End If

   If Trim(txtArquivo.Text) = "" Then
      MsgBox "Arquivo não pode ser vazio", vbCritical, App.ProductName

      txtArquivo.SetFocus

      Exit Function
   End If


   verificaTrailer = True

End Function

Private Function alteraTrailer() As Boolean
   Dim sMsg   As String
   Dim strSql As String
   Dim i      As Integer
   Dim strAux As String

   On Error GoTo TrataErro
  
   alteraTrailer = False
   
   If Not verificaTrailer() Then
       Exit Function
   End If
   
   i = mfgTrailer.Row

   gRSTrailer.Open "Select * from tb_trailer Where Arquivo = '" & Trim(txtArquivo.Tag) & "'", "FILE NAME=" & App.Path & "\Cinema.udl", adOpenDynamic, adLockOptimistic, adCmdText
   
   gRSTrailer.Fields("arquivo").Value = Trim(txtArquivo.Text)
   gRSTrailer.Update
   
   mfgTrailer.TextMatrix(i, 0) = Trim(txtArquivo.Text)
   
   gRSTrailer.Close
   
   alteraTrailer = True
   
   Exit Function
   
TrataErro:
   If gRSTrailer.State = adStateOpen Then
       gRSTrailer.Close
   End If
   
   sMsg = "Ocorreu um erro em alteraTrailer." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
   Exit Function
   Resume 0
End Function

Private Function insereTrailer() As Boolean
   Dim sMsg     As String
   Dim strSql   As String
   Dim strAux   As String
   Dim i        As Integer
   
   On Error GoTo TrataErro
   
   insereTrailer = False
   
   If Not verificaTrailer() Then
       Exit Function
   End If
   
   gRSTrailer.Open "tb_trailer", "FILE NAME=" & App.Path & "\Cinema.udl", adOpenDynamic, adLockOptimistic, adCmdTableDirect
   'gRSTrailer.Open strSql, "FILE NAME=" & App.Path & "\Cinema.udl", adOpenDynamic, adLockReadOnly, adCmdText
   
   

   gRSTrailer.AddNew
   
   'gRSTrailer.Fields("descricao").Value = Trim(txtFilmeTrailer.Text)
   gRSTrailer.Fields("arquivo").Value = Trim(txtArquivo.Text)
   gRSTrailer.Update
   
   If mfgTrailer.Rows > 2 Or mfgTrailer.TextMatrix(1, 0) <> "" Then
      mfgTrailer.Rows = mfgTrailer.Rows + 1
      i = mfgTrailer.Rows - 1
   Else
      i = 1
   End If
   
   'mfgTrailer.TextMatrix(i, 0) = Trim(txtFilmeTrailer.Text)
   mfgTrailer.TextMatrix(i, 0) = Trim(txtArquivo.Text)
   
   gRSTrailer.Close
   
   insereTrailer = True
   
   Exit Function
   
TrataErro:
   If gRSTrailer.State = adStateOpen Then
       gRSTrailer.Close
   End If
   
   sMsg = "Ocorreu um erro em insereTrailer." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
   Exit Function
   Resume 0
End Function

Private Sub mfgTrailer_Click()
   Call mfgTrailer_EnterCell
End Sub

Private Sub mfgTrailer_EnterCell()
   Dim i As Integer
   
   i = mfgTrailer.Row
   
   If i > 0 Then
      'txtFilmeTrailer.Text = mfgTrailer.TextMatrix(i, 0)
      txtArquivo.Text = mfgTrailer.TextMatrix(i, 0)
      txtArquivo.Tag = mfgTrailer.TextMatrix(i, 0)
   Else
      'txtFilmeTrailer.Text = ""
      txtArquivo.Text = ""
   End If
End Sub

Private Sub carregaFilmes()
   Dim sMsg   As String
   Dim strSql As String
   Dim i      As Integer
   Dim strAux As String
   
   On Error GoTo TrataErro
   
   strSql = "SELECT codFilme, "
   strSql = strSql & "descricao, "
   strSql = strSql & "duracao, "
   strSql = strSql & "censura "
   strSql = strSql & "FROM tb_filmes "
   strSql = strSql & "WHERE codFilme <> 0 "
   strSql = strSql & "ORDER BY descricao"
   
   gRSFilme.Open strSql, "FILE NAME=" & App.Path & "\Cinema.udl", adOpenDynamic, adLockReadOnly, adCmdText
   
   If Not (gRSFilme.EOF And gRSFilme.BOF) Then
      mfgFilme.FixedRows = 0
      mfgFilme.Rows = 1
      
      Do While Not gRSFilme.EOF
         mfgFilme.Rows = mfgFilme.Rows + 1
         i = mfgFilme.Rows - 1
         
         mfgFilme.TextMatrix(i, 0) = gRSFilme.Fields("descricao").Value
         strAux = Format(gRSFilme.Fields("duracao").Value, "00.00")
         strAux = Mid(strAux, 1, 2) & ":" & Mid(strAux, 4, 2)
         mfgFilme.TextMatrix(i, 1) = strAux
         mfgFilme.TextMatrix(i, 2) = CStr(gRSFilme.Fields("codFilme").Value)
         If Not IsNull(gRSFilme.Fields("censura").Value) Then
            mfgFilme.TextMatrix(i, 3) = gRSFilme.Fields("censura").Value
         Else
            mfgFilme.TextMatrix(i, 3) = ""
         End If
         
         gRSFilme.MoveNext
      Loop
      
      If mfgFilme.Rows = 1 Then
         mfgFilme.Rows = 2
      End If
      
      mfgFilme.FixedRows = 1
      mfgFilme.Row = 1
'      Call mfgFilme_Click
   Else
      mfgFilme.FixedRows = 0
      mfgFilme.Rows = 1
      
      mfgFilme.Rows = 2
      mfgFilme.FixedRows = 1
   End If
   
   gRSFilme.Close
   
   Exit Sub
   
TrataErro:
   If gRSFilme.State = adStateOpen Then
      gRSFilme.Close
   End If
   
   sMsg = "Ocorreu um erro em carregaFilmes." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
End Sub

