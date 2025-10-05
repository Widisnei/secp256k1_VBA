Attribute VB_Name = "Hash160_VBA"
Option Explicit
Option Compare Binary
' =====================================================================================
'  Hash160_VBA.bas — hash160(data) = RIPEMD-160(SHA-256(data))
'  Requer: módulos SHA256_VBA e RIPEMD160_VBA importados
'
'  API pública:
'    Hash160_String(ByVal s As String) As String
'    Hash160_Hex(ByVal hexStr As String) As String
'    Hash160_Bytes(ByRef data() As Byte) As String
'
'  (c) 2025, MIT License.
' =====================================================================================

Private Function HexToBytes(ByVal hexStr As String) As Byte()
    If (Len(hexStr) Mod 2) <> 0 Then
        Err.Raise vbObjectError + &H160, "Hash160_VBA.HexToBytes", _
                  "Hex string must contain an even number of characters."
    End If

    Dim n As Long, i As Long
    n = Len(hexStr) \ 2
    Dim b() As Byte
    If n > 0 Then
        ReDim b(0 To n - 1)
        For i = 0 To n - 1
            b(i) = CByte("&H" & Mid$(hexStr, 2 * i + 1, 2))
        Next
    End If
    HexToBytes = b
End Function

Public Function Hash160_Bytes(ByRef data() As Byte) As String
    Dim sha As String
    sha = SHA256_Bytes(data)          ' do módulo SHA256_VBA (hex de 32 bytes)

    Dim tmp() As Byte
    tmp = HexToBytes(sha)             ' bytes do SHA-256

    Hash160_Bytes = RIPEMD160_Bytes(tmp) ' hex de 20 bytes
End Function

Public Function Hash160_String(ByVal s As String) As String
    Dim b() As Byte
    If LenB(s) > 0 Then b = StrConv(s, vbFromUnicode)
    Hash160_String = Hash160_Bytes(b)
End Function

Public Function Hash160_Hex(ByVal hexStr As String) As String
    Dim b() As Byte
    On Error GoTo HexParseError
    b = HexToBytes(hexStr)
    On Error GoTo 0
    Hash160_Hex = Hash160_Bytes(b)
    Exit Function

HexParseError:
    Err.Raise Err.Number, "Hash160_VBA.Hash160_Hex", Err.Description
End Function

' ============================== Auto-Testes =========================================
Public Sub Hash160_SelfTest()
    Dim h As String, h2 As String, b() As Byte

    ' Vetor conhecido: hash160("") = B472A266D0BD89C13706A4132CCFB16F7C3B9F61
    h = Hash160_String("")
    Debug.Print "hash160("""") = " & h
    Debug.Print "Expect       = B472A266D0BD89C13706A4132CCFB16F7C3B9F61"

    ' Mesma entrada via HEX vazio
    h2 = Hash160_Hex("")
    Debug.Print "hash160_hex("""") = " & h2 & "  (igual ao anterior)"
    Debug.Print "Equal?        = " & CStr(h = h2)

    ' Checagem cruzada: Bytes -> String
    ' Entrada: bytes 00 01 02
    ReDim b(0 To 2)
    b(0) = &H0: b(1) = &H1: b(2) = &H2
    h = Hash160_Bytes(b)
    h2 = Hash160_Hex("000102")
    Debug.Print "hash160_bytes(000102) = " & h
    Debug.Print "hash160_hex(000102)   = " & h2
    Debug.Print "Equal?                = " & CStr(h = h2)

    ' Checagem por composição (sanity): RIPEMD160(SHA256(s)) == Hash160_String(s)
    Dim s As String
    s = "abc"
    Dim sha As String
    sha = SHA256_String(s)
    ' Converte sha (hex) em bytes e calcula RIPEMD-160
    Dim tmp() As Byte
    tmp = HexToBytes(sha)
    h2 = RIPEMD160_Bytes(tmp)
    h = Hash160_String(s)
    Debug.Print "hash160(""abc"") = " & h
    Debug.Print "via compose   = " & h2
    Debug.Print "Equal?        = " & CStr(h = h2)
End Sub