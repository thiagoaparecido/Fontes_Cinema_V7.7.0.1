VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPainelControle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Win32
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

'Dados Internos
Private msSeparador     As String
Private msDecimal       As String
Private msMilhar        As String
Private msMascaraData   As String
Private msSeparadorData As String
Private msMascaraHora   As String
Private msSeparadorHora As String
Private msMoeda         As String
Private msGrouping      As String

Private Sub Class_Initialize()
    Dim sAux As String
    Dim iRet As Integer
    Dim sX   As String
    
    sAux = Space(10)
    iRet = GetProfileString("intl", "sList", ";", sAux, Len(sAux))
    msSeparador = Left$(sAux, iRet)

    sAux = Space(10)
    iRet = GetProfileString("intl", "sDecimal", ",", sAux, Len(sAux))
    msDecimal = Left$(sAux, iRet)

    sAux = Space(10)
    iRet = GetProfileString("intl", "sThousand", ".", sAux, Len(sAux))
    msMilhar = Left$(sAux, iRet)

    sAux = Space(3)
    iRet = GetProfileString("intl", "sGrouping", "3;0", sAux, Len(sAux))
    msGrouping = Left$(sAux, 1)

    sAux = Space(50)
    iRet = GetProfileString("intl", "sDate", "/", sAux, Len(sAux))
    msSeparadorData = Left$(sAux, iRet)
    sX = msSeparadorData

    sAux = Space(50)
    iRet = GetProfileString("intl", "sShortDate", "dd" & sX & "mm" & sX & "yy", sAux, Len(sAux))
    msMascaraData = Left$(sAux, iRet)

    sAux = Space(50)
    iRet = GetProfileString("intl", "sTime", ":", sAux, Len(sAux))
    msSeparadorHora = Left$(sAux, iRet)
    sX = msSeparadorHora

    sAux = Space(50)
    iRet = GetProfileString("intl", "sTimeFormat", "H" & sX & "mm" & sX & "ss", sAux, Len(sAux))
    msMascaraHora = Left$(sAux, iRet)

    sAux = Space(10)
    iRet = GetProfileString("intl", "sCurrency", "R$ ", sAux, Len(sAux))
    msMoeda = Left$(sAux, iRet)

End Sub

Public Property Get Separador() As String
    Separador = msSeparador
End Property

Public Property Get SeparadorDecimal() As String
    SeparadorDecimal = msDecimal
End Property

Public Property Get Milhar() As String
    Milhar = msMilhar
End Property

Public Property Get NumDigitosGrupo() As Integer
    NumDigitosGrupo = Val(msGrouping)
End Property

Public Property Get SeparadorData() As String
    SeparadorData = msSeparadorData
End Property

Public Property Get MascaraData() As String
    MascaraData = msMascaraData
End Property

Public Property Get SeparadorHora() As String
    SeparadorHora = msSeparadorHora
End Property

Public Property Get MascaraHora() As String
    MascaraHora = msMascaraHora
End Property

Public Property Get Moeda() As String
    Moeda = msMoeda
End Property
