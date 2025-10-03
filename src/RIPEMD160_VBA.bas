Attribute VB_Name = "RIPEMD160_VBA"
Option Explicit
Option Compare Binary
' =====================================================================================
'  RIPEMD160_VBA.bas — Implementação pura em VBA do RIPEMD-160 (compatível com test vectors)
'  - Sem dependências externas
'  - Opera com palavras de 32 bits usando Long (assinado) com utilitários "unsigned"
'  - Saída em HEX maiúsculo (20 bytes)
'  - Padding estilo MD4/MD5 (comprimento em bits little-endian)
'
'  API pública:
'    RIPEMD160_String(ByVal s As String) As String
'    RIPEMD160_Hex(ByVal hexStr As String) As String
'    RIPEMD160_Bytes(ByRef data() As Byte) As String
'
'  Vetores conhecidos:
'    ""    -> 9C1185A5C5E9FC54612808977EE8F548B2258D31
'    "abc" -> 8EB208F7E05D987A9B044A8E98C6B087F15A0BFC
'
'  (c) 2025, MIT License.
' =====================================================================================

Private Const BLOCK_SIZE As Long = 64                 ' 512 bits
Private Const U32_MAX As Double = 4294967296#         ' 2^32
Private Const U31 As Double = 2147483648#             ' 2^31

' ========================= Utilitários Unsigned 32-bit ===============================
Private Function U32(ByVal x As Long) As Double
    Dim d As Double
    d = (x And &H7FFFFFFF)
    If (x And &H80000000) <> 0 Then d = d + U31
    U32 = d
End Function

Private Function D2L(ByVal u As Double) As Long
    If u >= U31 Then u = u - U32_MAX
    D2L = CLng(u)
End Function

Private Function UAdd(ByVal a As Long, ByVal b As Long) As Long
    Dim s As Double
    s = (U32(a) + U32(b))
    s = s - (Int(s / U32_MAX) * U32_MAX)
    UAdd = D2L(s)
End Function

Private Function UAdd3(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long
    Dim s As Double
    s = U32(a) + U32(b) + U32(c)
    s = s - (Int(s / U32_MAX) * U32_MAX)
    UAdd3 = D2L(s)
End Function

Private Function UAdd4(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long) As Long
    Dim s As Double
    s = U32(a) + U32(b) + U32(c) + U32(d)
    s = s - (Int(s / U32_MAX) * U32_MAX)
    UAdd4 = D2L(s)
End Function

Private Function ULShift(ByVal x As Long, ByVal n As Long) As Long
    If n = 0 Then ULShift = x: Exit Function
    Dim r As Double
    r = (U32(x) * (2 ^ n))
    r = r - (Int(r / U32_MAX) * U32_MAX)
    ULShift = D2L(r)
End Function

Private Function URShift(ByVal x As Long, ByVal n As Long) As Long
    If n = 0 Then URShift = x: Exit Function
    Dim r As Double
    r = Fix(U32(x) / (2 ^ n))
    URShift = D2L(r)
End Function

Private Function ROTL(ByVal x As Long, ByVal n As Long) As Long
    ROTL = (ULShift(x, n) Or URShift(x, 32 - n))
End Function

' ============================ API Pública ===========================================
Public Function RIPEMD160_String(ByVal s As String) As String
    Dim b() As Byte
    If LenB(s) > 0 Then
        b = StrConv(s, vbFromUnicode)
    End If
    RIPEMD160_String = RIPEMD160_Bytes(b)
End Function

Public Function RIPEMD160_Hex(ByVal hexStr As String) As String
    Dim n As Long, i As Long, b() As Byte
    n = Len(hexStr) \ 2
    If n > 0 Then
        ReDim b(0 To n - 1)
        For i = 0 To n - 1
            b(i) = CByte("&H" & Mid$(hexStr, 2 * i + 1, 2))
        Next
    End If
    RIPEMD160_Hex = RIPEMD160_Bytes(b)
End Function

Public Function RIPEMD160_Bytes(ByRef data() As Byte) As String
    Dim H0 As Long, H1 As Long, H2 As Long, H3 As Long, H4 As Long
    H0 = &H67452301
    H1 = &HEFCDAB89
    H2 = &H98BADCFE
    H3 = &H10325476
    H4 = &HC3D2E1F0

    Dim padded() As Byte
    Call PadMessage(data, padded)

    Dim r(0 To 79) As Byte, rp(0 To 79) As Byte
    Dim s(0 To 79) As Byte, sp(0 To 79) As Byte
    Call Init_r_s(r, s, rp, sp)

    Dim K(0 To 4) As Long, Kp(0 To 4) As Long
    K(0) = &H0:          K(1) = &H5A827999: K(2) = &H6ED9EBA1: K(3) = &H8F1BBCDC: K(4) = &HA953FD4E
    Kp(0) = &H50A28BE6:  Kp(1) = &H5C4DD124: Kp(2) = &H6D703EF3: Kp(3) = &H7A6D76E9: Kp(4) = &H0

    Dim i As Long, j As Long
    Dim X(0 To 15) As Long
    Dim a1 As Long, b1 As Long, c1 As Long, d1 As Long, e1 As Long
    Dim a2 As Long, b2 As Long, c2 As Long, d2 As Long, e2 As Long
    Dim T As Long, t0 As Long

    For i = 0 To UBound(padded) Step BLOCK_SIZE
        ' Carrega 16 palavras little-endian
        For j = 0 To 15
            X(j) = CLng(padded(i + 4 * j)) _
                 Or ULShift(CLng(padded(i + 4 * j + 1)), 8) _
                 Or ULShift(CLng(padded(i + 4 * j + 2)), 16) _
                 Or ULShift(CLng(padded(i + 4 * j + 3)), 24)
            X(j) = D2L(CDbl(X(j))) ' garantir wrap 32-bit
        Next j

        a1 = H0: b1 = H1: c1 = H2: d1 = H3: e1 = H4
        a2 = H0: b2 = H1: c2 = H2: d2 = H3: e2 = H4

        For j = 0 To 79
            T = UAdd4(a1, f(j, b1, c1, d1), X(r(j)), K(j \ 16))
            T = ROTL(T, s(j))
            T = UAdd(T, e1)
            a1 = e1: e1 = d1: d1 = ROTL(c1, 10): c1 = b1: b1 = T

            T = UAdd4(a2, fp(j, b2, c2, d2), X(rp(j)), Kp(j \ 16))
            T = ROTL(T, sp(j))
            T = UAdd(T, e2)
            a2 = e2: e2 = d2: d2 = ROTL(c2, 10): c2 = b2: b2 = T
        Next j

        t0 = UAdd3(H1, c1, d2)
        H1 = UAdd3(H2, d1, e2)
        H2 = UAdd3(H3, e1, a2)
        H3 = UAdd3(H4, a1, b2)
        H4 = UAdd3(H0, b1, c2)
        H0 = t0
    Next i

    RIPEMD160_Bytes = _
        UCase$(Word32ToHexLE(H0) & Word32ToHexLE(H1) & Word32ToHexLE(H2) & Word32ToHexLE(H3) & Word32ToHexLE(H4))
End Function

' ============================= Padding ==============================================
Private Sub PadMessage(ByRef data() As Byte, ByRef padded() As Byte)
    Dim msgLen As Long
    Dim lb As Long, ub As Long

    On Error Resume Next
    lb = LBound(data)
    ub = UBound(data)
    If Err.Number <> 0 Then
        Err.Clear
        msgLen = 0
    ElseIf ub < lb Then
        msgLen = 0
    Else
        msgLen = ub - lb + 1
    End If
    On Error GoTo 0

    Dim newLen As Long
    newLen = msgLen + 1
    Do While (newLen Mod BLOCK_SIZE) <> 56
        newLen = newLen + 1
    Loop
    newLen = newLen + 8 ' comprimento 64-bit little-endian

    ReDim padded(0 To newLen - 1) As Byte

    Dim i As Long
    For i = 0 To msgLen - 1
        padded(i) = data(i)
    Next i

    padded(msgLen) = &H80

    ' comprimento em bits 64-bit little-endian
    Dim bitLen As Double
    bitLen = CDbl(msgLen) * 8#
    Dim hi As Double, lo As Double
    hi = Fix(bitLen / U32_MAX)
    lo = bitLen - hi * U32_MAX

    ' low 32
    padded(newLen - 8) = D2L(lo) And &HFF
    padded(newLen - 7) = URShift(D2L(lo), 8) And &HFF
    padded(newLen - 6) = URShift(D2L(lo), 16) And &HFF
    padded(newLen - 5) = URShift(D2L(lo), 24) And &HFF
    ' high 32
    padded(newLen - 4) = D2L(hi) And &HFF
    padded(newLen - 3) = URShift(D2L(hi), 8) And &HFF
    padded(newLen - 2) = URShift(D2L(hi), 16) And &HFF
    padded(newLen - 1) = URShift(D2L(hi), 24) And &HFF
End Sub

' ============================ Funções f / f' ========================================
Private Function f(ByVal j As Long, ByVal x As Long, ByVal y As Long, ByVal z As Long) As Long
    Select Case j \ 16
        Case 0: f = (x Xor y Xor z)
        Case 1: f = (x And y) Or ((Not x) And z)
        Case 2: f = (x Or (Not y)) Xor z
        Case 3: f = (x And z) Or (y And (Not z))
        Case 4: f = (x Xor (y Or (Not z)))
    End Select
End Function

Private Function fp(ByVal j As Long, ByVal x As Long, ByVal y As Long, ByVal z As Long) As Long
    Select Case j \ 16
        Case 0: fp = (x Xor (y Or (Not z)))
        Case 1: fp = (x And z) Or (y And (Not z))
        Case 2: fp = (x Or (Not y)) Xor z
        Case 3: fp = (x And y) Or ((Not x) And z)
        Case 4: fp = (x Xor y Xor z)
    End Select
End Function

' ============================ Permutações/Rotações ==================================
Private Sub Init_r_s(ByRef r() As Byte, ByRef s() As Byte, ByRef rp() As Byte, ByRef sp() As Byte)
    Dim rVals As Variant, sVals As Variant, rpVals As Variant, spVals As Variant

    rVals = Array( _
    0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15, _
    7,4,13,1,10,6,15,3,12,0,9,5,2,14,11,8, _
    3,10,14,4,9,15,8,1,2,7,0,6,13,11,5,12, _
    1,9,11,10,0,8,12,4,13,3,7,15,14,5,6,2, _
    4,0,5,9,7,12,2,10,14,1,3,8,11,6,15,13)

    sVals = Array( _
    11,14,15,12,5,8,7,9,11,13,14,15,6,7,9,8, _
    7,6,8,13,11,9,7,15,7,12,15,9,11,7,13,12, _
    11,13,6,7,14,9,13,15,14,8,13,6,5,12,7,5, _
    11,12,14,15,14,15,9,8,9,14,5,6,8,6,5,12, _
    9,15,5,11,6,8,13,12,5,12,13,14,11,8,5,6)

    rpVals = Array( _
    5,14,7,0,9,2,11,4,13,6,15,8,1,10,3,12, _
    6,11,3,7,0,13,5,10,14,15,8,12,4,9,1,2, _
    15,5,1,3,7,14,6,9,11,8,12,2,10,0,4,13, _
    8,6,4,1,3,11,15,0,5,12,2,13,9,7,10,14, _
    12,15,10,4,1,5,8,7,6,2,13,14,0,3,9,11)

    spVals = Array( _
    8,9,9,11,13,15,15,5,7,7,8,11,14,14,12,6, _
    9,13,15,7,12,8,9,11,7,7,12,7,6,15,13,11, _
    9,7,15,11,8,6,6,14,12,13,5,14,13,13,7,5, _
    15,5,8,11,14,14,6,14,6,9,12,9,12,5,15,8, _
    8,5,12,9,12,5,14,6,8,13,6,5,15,13,11,11)

    Dim i As Long
    For i = 0 To 79
        r(i) = rVals(i)
        s(i) = sVals(i)
        rp(i) = rpVals(i)
        sp(i) = spVals(i)
    Next i
End Sub

' ========================== Helpers de conversão ====================================
Private Function Word32ToHexLE(ByVal w As Long) As String
    Dim b0 As Long, b1 As Long, b2 As Long, b3 As Long
    b0 = w And &HFF
    b1 = URShift(w, 8) And &HFF
    b2 = URShift(w, 16) And &HFF
    b3 = URShift(w, 24) And &HFF
    Word32ToHexLE = ByteToHex(b0) & ByteToHex(b1) & ByteToHex(b2) & ByteToHex(b3)
End Function

Private Function ByteToHex(ByVal x As Long) As String
    Dim s As String
    s = Hex$(x And &HFF)
    If Len(s) < 2 Then s = "0" & s
    ByteToHex = s
End Function

' ============================== Auto-Testes =========================================
Public Sub RIPEMD160_SelfTest()
    Dim h As String
    h = RIPEMD160_String("")
    Debug.Print "RIPEMD160("""") = " & h
    Debug.Print "Expect       = 9C1185A5C5E9FC54612808977EE8F548B2258D31"

    h = RIPEMD160_String("abc")
    Debug.Print "RIPEMD160(""abc"") = " & h
    Debug.Print "Expect       = 8EB208F7E05D987A9B044A8E98C6B087F15A0BFC"
End Sub