VERSION 5.00
Begin VB.Form frmExpurgo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2280
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Aguarde, Carregando Sistema...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   1132
      TabIndex        =   0
      Top             =   923
      Width           =   5220
   End
End
Attribute VB_Name = "frmExpurgo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    Dim expurgo As New clsExpurgo
    Dim msgErro As String
    
    DoEvents
    
    Set expurgo.ConexaoADO = dbConnect
    expurgo.dias = CInt(pDiasExpurgo)
    Call expurgo.expurgo
        
    If expurgo.CodigoErro <> 0 Then
       MsgBox "Ocorreu um erro no processo de expurgo: " & expurgo.MensagemErro, vbCritical, App.ProductName
    End If

    Unload Me
End Sub

