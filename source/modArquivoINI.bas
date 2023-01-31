Attribute VB_Name = "modArquivoINI"
Option Explicit
'///////////////////////////////////////////////////////////////////////
' API para Ler/Escrever Arquivo .INI e Pegar o Usuário Logado

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function GravarINI(ByVal sSecao As String, ByVal sChave As String, ByVal sValor As String, ByVal sArqIni As String) As Long
    GravarINI = WritePrivateProfileString(sSecao, sChave, sValor, sArqIni)
End Function

Public Function LerINI(ByVal sSecao As String, ByVal sChave As String, ByVal sArqIni As String, Optional ByVal sDefault As String = "(Nada)") As String
    Dim lngRet As Long
    Dim strRet As String
    
    strRet = Space(255)
    lngRet = GetPrivateProfileString(sSecao, sChave, sDefault, strRet, Len(strRet), sArqIni)
    LerINI = Trim(Left$(strRet, lngRet))
End Function
