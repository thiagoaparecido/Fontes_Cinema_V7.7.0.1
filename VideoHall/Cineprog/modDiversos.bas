Attribute VB_Name = "modDiversos"
'Option Explicit


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Declare Function GetShortPathName Lib "kernel32" _
                        Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
                        ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Declare Function SendMessageAsLong Lib "user32" _
Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function SendMessageAsString Lib "user32" _
Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, ByVal lParam As String) As Long


'Variaveis do "Registry" da aplica��o
Public gsRegistryKey   As String               'Chave do Registro da aplica��o no "Registry"
Public gsRegistryKeyDB As String               'Chave das informa��es para conex�o com o banco de dados

Public gsCharMil       As String
Public gsCharDec       As String
Public giCharDec       As Integer

Public Const giMAX_PATH = 260
Public Const EM_GETLINE = 196
Public Const EM_GETLINECOUNT = 186
Public Const MAX_CHAR_PER_LINE = 200  ' Scale this to size of text box.


Public Function EsconderMouse()

Do
  LMostraCursor = LMostraCursor - 1
  rtn = ShowCursor(False)
  Loop Until rtn < 0

End Function
Public Function ExibirMouse()

Do
  LMostraCursor = LMostraCursor - 1
  rtn = ShowCursor(True)
Loop Until rtn >= 0

End Function



Public Function fGetLine(LineNumber As Long, ctrl1 As Control) As String
    ' This function fills the buffer with a line of text
    ' specified by LineNumber from the text-box control.
    ' The first line starts at zero.
    Dim byteLo As Integer
    Dim byteHi As Integer
    Dim X      As Long
    Dim Buffer As String
    
    byteLo = MAX_CHAR_PER_LINE And (255)  '[changed 5/15/92]
    byteHi = Int(MAX_CHAR_PER_LINE / 256) '[changed 5/15/92]
    Buffer = Chr$(byteLo) + Chr$(byteHi) + Space$(MAX_CHAR_PER_LINE - 2)

    X = SendMessageAsString(ctrl1.hWnd, EM_GETLINE, LineNumber, Buffer)

    fGetLine = Left$(Buffer, X)
End Function

Public Function fGetLineCount(ctrl1 As Control) As Long
    Dim lcount As Long
    ' This function will return the number of lines
    ' currently in the text-box control.
    ' Setfocus method illegal while in resize event,
    ' so use global flag to see if called from there
    ' (or use setfocus before this function call in general case).

    lcount = SendMessageAsLong(ctrl1.hWnd, EM_GETLINECOUNT, 0, 0)

    fGetLineCount = lcount
End Function

Public Function converteNumero(numero As String) As Double
    Dim aux As String
    
    aux = Replace(numero, clsPC.Milhar, "")
    aux = Replace(aux, clsPC.SeparadorDecimal, ".")
    
    converteNumero = Val(aux)
End Function

'**********************************************************************************************
'  Rotina      : gCentralizaTela
'  Descri��o   : Centraliza um FORM em rela��o a um MDIFORM ou Screen
'  Argumentos  : TelaPrinc  - um objeto MDIForm ou Screen
'                TelaSecund - Tela a ser centralizada
'  Retorno     : N�o Possui
'  Autor       : Pedro Am�rico Abril/2000
'  Altera��o   :
'**********************************************************************************************
Public Sub gCentralizaTela(TelaPrinc As Object, TelaSecund As Object)
    If TypeOf TelaPrinc Is MDIForm Then
        TelaSecund.Top = IIf(TelaSecund.Height < TelaPrinc.ScaleHeight, (TelaPrinc.ScaleHeight - TelaSecund.Height) / 2, 0)
        TelaSecund.Left = IIf(TelaSecund.Width < TelaPrinc.ScaleWidth, (TelaPrinc.ScaleWidth - TelaSecund.Width) / 2, 0)
    ElseIf TypeOf TelaPrinc Is Screen Then
        TelaSecund.Top = IIf(TelaSecund.Height < TelaPrinc.Height, (TelaPrinc.Height - TelaSecund.Height) / 2, 0)
        TelaSecund.Left = IIf(TelaSecund.Width < TelaPrinc.Width, (TelaPrinc.Width - TelaSecund.Width) / 2, 0)
    End If
End Sub

Public Function gGetShortPathName(ByVal vsLongPath As String) As String
   Dim sShortPath As String
   Dim lPathLen   As Long
   Dim lLen       As Long
   
   sShortPath = Space$(giMAX_PATH)
   lLen = Len(sShortPath)
   lPathLen = GetShortPathName(vsLongPath, sShortPath, lLen)
   gGetShortPathName = Left$(sShortPath, lPathLen)
End Function

Public Function gsEliminaAcentos(sString As String) As String
    Dim I   As Integer
    Dim str As String

    str = ""
    
    For I = 1 To Len(sString)
        Select Case Mid(sString, I, 1)
            Case "�", "�", "�", "�", "�"
                str = str & "a"
            Case "�", "�", "�", "�", "�"
                str = str & "A"
            Case "�", "�", "�", "�"
                str = str & "e"
            Case "�", "�", "�", "�"
                str = str & "E"
            Case "�", "�", "�", "�"
                str = str & "i"
            Case "�", "�", "�", "�"
                str = str & "I"
            Case "�", "�", "�", "�", "�"
                str = str & "o"
            Case "�", "�", "�", "�", "�"
                str = str & "O"
            Case "�", "�", "�", "�"
                str = str & "u"
            Case "�", "�", "�", "�"
                str = str & "U"
            Case "�"
                str = str & "n"
            Case "�"
                str = str & "N"
            Case "�"
                str = str & "c"
            Case "�"
                str = str & "C"
            Case "�", "�"
                str = str & "y"
            Case "�", "�"
                str = str & "Y"
            Case Else
                If Asc(Mid(sString, I, 1)) >= 32 And _
                   Asc(Mid(sString, I, 1)) <= 126 Then
                    str = str & Mid(sString, I, 1)
                Else
                    str = str & " "
                End If
        End Select
    Next I

    gsEliminaAcentos = str
End Function

