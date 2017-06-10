VERSION 5.00
Begin VB.Form frmExpurgo 
   Caption         =   "Expurgo"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1755
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancela 
      Cancel          =   -1  'True
      Caption         =   "Cancela"
      Height          =   420
      Left            =   3262
      TabIndex        =   3
      Top             =   990
      Width           =   1590
   End
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
      Left            =   1342
      TabIndex        =   2
      Top             =   990
      Width           =   1590
   End
   Begin VB.TextBox txtDias 
      Height          =   345
      Left            =   4500
      MaxLength       =   4
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblDias 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade de dias (Base de dados disponíveis):"
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
      Left            =   240
      TabIndex        =   1
      Top             =   435
      Width           =   4200
   End
End
Attribute VB_Name = "frmExpurgo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim expurgo As New clsExpurgo
    
    If Not IsNumeric(txtDias.Text) Then
        MsgBox "Valor invalido para numero de dias!", vbCritical, App.ProductName
        Exit Sub
    End If
    
    If CInt(txtDias.Text) < 30 Then
        MsgBox "Número de dias deve ser maior que 30!", vbCritical, App.ProductName
        Exit Sub
    End If
    
    Set expurgo.ConexaoADO = dbConnect
    expurgo.dias = CInt(txtDias.Text)
    Call expurgo.expurgoCentral
    
    If expurgo.CodigoErro <> 0 Then
        MsgBox "Ocorreu um erro no processo de expurgo: " & expurgo.MensagemErro, vbCritical, App.ProductName
        
        Exit Sub
    End If
    
    MsgBox "Expurgo realizado com sucesso!", vbInformation, App.ProductName
    
    Unload Me
End Sub

Private Sub Form_Load()
    txtDias.Text = pDiasExpurgo
End Sub
