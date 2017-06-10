VERSION 5.00
Begin VB.Form frmTipoCor 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3690
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancela 
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
      Left            =   1965
      TabIndex        =   4
      Top             =   1905
      Width           =   1260
   End
   Begin VB.CommandButton cmdOk 
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
      Left            =   420
      TabIndex        =   3
      Top             =   1905
      Width           =   1260
   End
   Begin VB.OptionButton optCor 
      Caption         =   "Cor do Texto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   1275
      Width           =   3480
   End
   Begin VB.OptionButton optCor 
      Caption         =   "Cor do Fundo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Top             =   825
      Width           =   3480
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Altera qual Cor?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   735
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmTipoCor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mOk  As Boolean
Private mCor As Integer

Public Function loadCor(ByRef cor As Integer) As Boolean
   
   On Error GoTo TrataErro
   
   mOk = False
   optCor(0).Value = True
   
   Me.Show vbModal
   
   loadCor = mOk
   cor = mCor
   
   Unload Me
   Exit Function
TrataErro:
   If Err.Number <> 400 Then
      MsgBox Err.Number & " - " & Err.Description
      End
   End If
End Function

Private Sub cmdCancela_Click()
   Unload Me
End Sub

Private Sub cmdOk_Click()
   If optCor(0).Value Then
      mCor = 1
   Else
      mCor = 2
   End If
   
   mOk = True
   
   Unload Me
End Sub

