Attribute VB_Name = "Funcoes"
Option Explicit

Public Function GFcn_MsgBox(Prompt, Optional Buttons As VbMsgBoxStyle = vbInformation, Optional Title = "Austin Rating") As VbMsgBoxResult
    Dim VLStr_Prefixo   As String
    Dim WLStr_Msg       As String

    VLStr_Prefixo = "Atenção!" & Chr(10) & Chr(10)
    WLStr_Msg = Prompt
    
    Beep
    GFcn_MsgBox = MsgBox(VLStr_Prefixo & WLStr_Msg, Buttons, Title)
    
End Function

Public Function GFcn_PegaNomeUsuario() As String
    Dim VLStr_NomeUsr   As String
    Dim VLLng_Tamanho   As Long
    
    VLStr_NomeUsr = Space$(255)
    VLLng_Tamanho = Len(VLStr_NomeUsr)
    
    GetUserName VLStr_NomeUsr, VLLng_Tamanho
    
    GFcn_PegaNomeUsuario = Trim$(VLStr_NomeUsr)
End Function
Public Sub SendKeysTab()
    keybd_event KeyTab, 0, KeyDown, 0
    keybd_event KeyTab, 0, KeyUP, 0
End Sub

