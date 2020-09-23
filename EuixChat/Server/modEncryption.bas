Attribute VB_Name = "modEncryption"
Public Function Encrypt(StringToEncrypt As String, Optional AlphaEncoding As Boolean = False) As String
    On Error GoTo ErrorHandler
    Dim Char As String
    Encrypt = ""
    
    For i = 1 To Len(StringToEncrypt)
        Char = Asc(Mid(StringToEncrypt, i, 1))
        Encrypt = Encrypt & Len(Char) & Char
    Next i
    
    If AlphaEncoding Then
    
        StringToEncrypt = Encrypt
        Encrypt = ""
        
        For i = 1 To Len(StringToEncrypt)
            Encrypt = Encrypt & Chr(Mid(StringToEncrypt, i, 1) + 147)
        Next i
        
    End If
    Exit Function
ErrorHandler:
    Encrypt = "Error encrypting string"
End Function

Public Function Decrypt(StringToDecrypt As String, Optional AlphaDecoding As Boolean = False) As String
    On Error GoTo ErrorHandler
    Dim CharCode As String
    Dim CharPos As Integer
    Dim Char As String
    
    If AlphaDecoding Then
    
        Decrypt = StringToDecrypt
        StringToDecrypt = ""
        
        For i = 1 To Len(Decrypt)
            StringToDecrypt = StringToDecrypt & (Asc(Mid(Decrypt, i, 1)) - 147)
        Next i
        
    End If
    
    Decrypt = ""
    
    Do
    
        CharPos = Left(StringToDecrypt, 1)
        StringToDecrypt = Mid(StringToDecrypt, 2)
        CharCode = Left(StringToDecrypt, CharPos)
        StringToDecrypt = Mid(StringToDecrypt, Len(CharCode) + 1)
        Decrypt = Decrypt & Chr(CharCode)
        
    Loop Until StringToDecrypt = ""
    Exit Function
ErrorHandler:
    Decrypt = "Error decrypting string"
End Function
