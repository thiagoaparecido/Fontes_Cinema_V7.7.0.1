VERSION 5.00
Begin VB.Form frmTipoImp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Impressora"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4605
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancela 
      Cancel          =   -1  'True
      Caption         =   "&Cancela"
      Height          =   390
      Left            =   2895
      TabIndex        =   3
      Top             =   585
      Width           =   1438
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   2895
      TabIndex        =   2
      Top             =   105
      Width           =   1438
   End
   Begin VB.OptionButton opTipoImp 
      Caption         =   "80 Colunas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   135
      TabIndex        =   1
      Top             =   570
      Width           =   2565
   End
   Begin VB.OptionButton opTipoImp 
      Caption         =   "40 Colunas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Value           =   -1  'True
      Width           =   2565
   End
End
Attribute VB_Name = "frmTipoImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private tpImp As Integer '1 - 80 colunas
                         '2 - 40 colunas
                         '3 - cancela

Public Function tipoImpo() As Integer
    tpImp = 3
    
    Me.Show vbModal
    
    tipoImpo = tpImp
End Function

Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If opTipoImp(0).Value Then
        tpImp = 2
    Else
        tpImp = 1
    End If
    
    Unload Me
End Sub
