Attribute VB_Name = "SHA256_VBA"
Option Explicit
Option Compare Binary
' =====================================================================================
'  SHA256_VBA.bas — Implementação pura em VBA do SHA-256 (compatível com test vectors)
'  - Sem dependências de BigInt
'  - Opera com palavras de 32 bits usando Long (assinado) com utilitários "unsigned"
'  - Saída em HEX maiúsculo, byte-order big-endian (padrão SHA-256)
'  - Compatível com VBA7 (Office 64-bit) e versões anteriores
'
'  API pública:
'    SHA256_String(ByVal s As String) As String     ' hash de string (ASCII, via codepage local)
'    SHA256_Hex(ByVal hexStr As String) As String   ' hash de bytes de uma string hex
'    SHA256_Bytes(ByRef data() As Byte) As String   ' hash de um array de bytes
'
'  Testes rápidos (Sub de auto-teste no final):
'    Call SHA256_SelfTest
'
'  Vectors esperados:
'    ""     -> E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855
'    "abc"  -> BA7816BF8F01CFEA414140DE5DAE2223B00361A396177A9CB410FF61F20015AD
'
'  (c) 2025, MIT License.
' =====================================================================================

' =============================== Constantes =========================================
Private Const BLOCK_SIZE As Long = 64                 ' 512 bits
Private Const U32_MAX As Double = 4294967296#         ' 2^32
Private Const U31 As Double = 2147483648#             ' 2^31

' ========================= Utilitários Unsigned 32-bit ===============================
' Converte Long assinado para Double no range [0 .. 2^32-1]
Private Function U32(ByVal x As Long) As Double
    Dim d As Double
    d = (x And &H7FFFFFFF)
    If (x And &H80000000) <> 0 Then d = d + U31
    U32 = d
End Function

' Converte Double no range [0 .. 2^32-1] de volta para Long (wrap 32-bit)
Private Function D2L(ByVal u As Double) As Long
    If u >= U31 Then u = u - U32_MAX
    D2L = CLng(u)
End Function

' Soma unsigned 32-bit: (a + b) mod 2^32
Private Function UAdd(ByVal a As Long, ByVal b As Long) As Long
    Dim s As Double
    s = (U32(a) + U32(b))
    s = s - (Int(s / U32_MAX) * U32_MAX)
    UAdd = D2L(s)
End Function

' Soma de 3 termos unsigned 32-bit
Private Function UAdd3(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long
    Dim s As Double
    s = U32(a) + U32(b) + U32(c)
    s = s - (Int(s / U32_MAX) * U32_MAX)
    UAdd3 = D2L(s)
End Function

' Soma de 5 termos unsigned 32-bit
Private Function UAdd5(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal e As Long) As Long
    Dim s As Double
    s = U32(a) + U32(b) + U32(c) + U32(d) + U32(e)
    s = s - (Int(s / U32_MAX) * U32_MAX)
    UAdd5 = D2L(s)
End Function

' Deslocamento lógico à direita (x >>> n), 0 <= n <= 31
Private Function URShift(ByVal x As Long, ByVal n As Long) As Long
    If n = 0 Then
        URShift = x
        Exit Function
    End If
    Dim r As Double
    r = Fix(U32(x) / (2 ^ n))
    URShift = D2L(r)
End Function

' Deslocamento à esquerda com máscara de 32 bits (x << n) & 0xFFFFFFFF
Private Function ULShift(ByVal x As Long, ByVal n As Long) As Long
    If n = 0 Then
        ULShift = x
        Exit Function
    End If
    Dim r As Double
    r = (U32(x) * (2 ^ n))
    r = r - (Int(r / U32_MAX) * U32_MAX)
    ULShift = D2L(r)
End Function

' Rotação direita de 32 bits: ROTR(x, n)
Private Function ROTR(ByVal x As Long, ByVal n As Long) As Long
    ROTR = (URShift(x, n) Or ULShift(x, 32 - n))
End Function

' ============================ Funções SHA-256 =======================================
' Funções menores sigma0, sigma1
Private Function sigma0(ByVal x As Long) As Long
    sigma0 = (ROTR(x, 7) Xor ROTR(x, 18) Xor URShift(x, 3))
End Function

Private Function sigma1(ByVal x As Long) As Long
    sigma1 = (ROTR(x, 17) Xor ROTR(x, 19) Xor URShift(x, 10))
End Function

' Funções maiores Σ0, Σ1
Private Function bigSigma0(ByVal x As Long) As Long
    bigSigma0 = (ROTR(x, 2) Xor ROTR(x, 13) Xor ROTR(x, 22))
End Function

Private Function bigSigma1(ByVal x As Long) As Long
    bigSigma1 = (ROTR(x, 6) Xor ROTR(x, 11) Xor ROTR(x, 25))
End Function

' Escolha: Ch(x,y,z) = (x AND y) XOR ((NOT x) AND z)
Private Function Ch32(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Long
    Ch32 = ((x And y) Xor ((Not x) And z))
End Function

' Maioria: Maj(x,y,z) = (x AND y) XOR (x AND z) XOR (y AND z)
Private Function Maj32(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Long
    Maj32 = ((x And y) Xor (x And z) Xor (y And z))
End Function

' ============================ API Pública ===========================================
Public Function SHA256_String(ByVal s As String) As String
    Dim b() As Byte
    If LenB(s) = 0 Then
        ' array não alocado = mensagem vazia
    Else
        b = StrConv(s, vbFromUnicode)
    End If
    SHA256_String = SHA256_Bytes(b)
End Function

Public Function SHA256_Hex(ByVal hexStr As String) As String
    Dim n As Long, i As Long
    n = Len(hexStr) \ 2
    Dim b() As Byte
    If n > 0 Then
        ReDim b(0 To n - 1)
        For i = 0 To n - 1
            b(i) = CByte("&H" & Mid$(hexStr, 2 * i + 1, 2))
        Next
    End If
    SHA256_Hex = SHA256_Bytes(b)
End Function

Public Function SHA256_Bytes(ByRef data() As Byte) As String
    Dim H(0 To 7) As Long
    H(0) = &H6A09E667
    H(1) = &HBB67AE85
    H(2) = &H3C6EF372
    H(3) = &HA54FF53A
    H(4) = &H510E527F
    H(5) = &H9B05688C
    H(6) = &H1F83D9AB
    H(7) = &H5BE0CD19

    Dim K(0 To 63) As Long
    Call K_Init(K)

    Dim padded() As Byte
    Call PadMessage(data, padded)

    Dim i As Long, j As Long, t As Long
    Dim a1 As Long, b1 As Long, c1 As Long, d1 As Long, e1 As Long, f1 As Long, g1 As Long, h1 As Long
    Dim W(0 To 63) As Long

    For i = 0 To UBound(padded) Step BLOCK_SIZE
        ' 1) Preparar W[0..63] (big-endian) sem multiplicações em Double
        For j = 0 To 15
            W(j) = _
                  ULShift(CLng(padded(i + 4 * j)), 24) _
                Or ULShift(CLng(padded(i + 4 * j + 1)), 16) _
                Or ULShift(CLng(padded(i + 4 * j + 2)), 8) _
                Or CLng(padded(i + 4 * j + 3))
        Next j
        For j = 16 To 63
            W(j) = UAdd3( sigma1(W(j - 2)), W(j - 7), sigma0(W(j - 15)) )
            W(j) = UAdd(W(j), W(j - 16))
        Next j

        ' 2) Inicializar registradores de trabalho
        a1 = H(0): b1 = H(1): c1 = H(2): d1 = H(3)
        e1 = H(4): f1 = H(5): g1 = H(6): h1 = H(7)

        ' 3) 64 rodadas
        Dim T1 As Long, T2 As Long
        For t = 0 To 63
            T1 = UAdd5(h1, bigSigma1(e1), Ch32(e1, f1, g1), K(t), W(t))
            T2 = UAdd(bigSigma0(a1), Maj32(a1, b1, c1))

            h1 = g1
            g1 = f1
            f1 = e1
            e1 = UAdd(d1, T1)
            d1 = c1
            c1 = b1
            b1 = a1
            a1 = UAdd(T1, T2)
        Next t

        ' 4) Adicionar ao hash
        H(0) = UAdd(H(0), a1)
        H(1) = UAdd(H(1), b1)
        H(2) = UAdd(H(2), c1)
        H(3) = UAdd(H(3), d1)
        H(4) = UAdd(H(4), e1)
        H(5) = UAdd(H(5), f1)
        H(6) = UAdd(H(6), g1)
        H(7) = UAdd(H(7), h1)
    Next i

    ' 5) Converter para HEX (big-endian por palavra)
    Dim outHex As String
    outHex = ""
    For i = 0 To 7
        outHex = outHex & Word32ToHex(H(i))
    Next i
    SHA256_Bytes = UCase$(outHex)
End Function

Public Function SHA256_HMAC(ByRef key() As Byte, ByRef data() As Byte) As Byte()
    Const BLOCK_SIZE As Long = 64

    Dim blockKey() As Byte
    ReDim blockKey(0 To BLOCK_SIZE - 1)

    Dim keyLen As Long
    keyLen = ByteArrayLength(key)

    Dim i As Long
    If keyLen > BLOCK_SIZE Then
        Dim hashedHex As String
        hashedHex = SHA256_Bytes(key)
        Dim hashed() As Byte
        hashed = HexToBytes(hashedHex)
        Dim hashedLen As Long
        hashedLen = ByteArrayLength(hashed)
        Dim baseHashed As Long
        baseHashed = LBoundSafe(hashed)
        For i = 0 To hashedLen - 1
            blockKey(i) = hashed(baseHashed + i)
        Next i
    ElseIf keyLen > 0 Then
        Dim baseKey As Long
        baseKey = LBoundSafe(key)
        For i = 0 To keyLen - 1
            blockKey(i) = key(baseKey + i)
        Next i
    End If

    Dim ipad() As Byte, opad() As Byte
    ReDim ipad(0 To BLOCK_SIZE - 1)
    ReDim opad(0 To BLOCK_SIZE - 1)
    For i = 0 To BLOCK_SIZE - 1
        ipad(i) = blockKey(i) Xor &H36
        opad(i) = blockKey(i) Xor &H5C
    Next i

    Dim innerInput() As Byte
    innerInput = ByteArrayConcat(ipad, data)
    Dim innerHashHex As String
    innerHashHex = SHA256_Bytes(innerInput)
    Dim innerHash() As Byte
    innerHash = HexToBytes(innerHashHex)

    Dim outerInput() As Byte
    outerInput = ByteArrayConcat(opad, innerHash)
    Dim outerHashHex As String
    outerHashHex = SHA256_Bytes(outerInput)
    SHA256_HMAC = HexToBytes(outerHashHex)
End Function

Private Function ByteArrayLength(ByRef arr() As Byte) As Long
    On Error GoTo EmptyArray
    ByteArrayLength = UBound(arr) - LBound(arr) + 1
    Exit Function
EmptyArray:
    ByteArrayLength = 0
End Function

Private Function LBoundSafe(ByRef arr() As Byte) As Long
    On Error GoTo EmptyArray
    LBoundSafe = LBound(arr)
    Exit Function
EmptyArray:
    LBoundSafe = 0
End Function

Private Function ByteArrayConcat(ByRef a() As Byte, ByRef b() As Byte) As Byte()
    Dim lenA As Long, lenB As Long
    lenA = ByteArrayLength(a)
    lenB = ByteArrayLength(b)

    Dim total As Long
    total = lenA + lenB

    Dim result() As Byte
    If total <= 0 Then
        ByteArrayConcat = result
        Exit Function
    End If

    ReDim result(0 To total - 1)

    Dim baseA As Long, baseB As Long
    Dim i As Long

    If lenA > 0 Then
        baseA = LBoundSafe(a)
        For i = 0 To lenA - 1
            result(i) = a(baseA + i)
        Next i
    End If

    If lenB > 0 Then
        baseB = LBoundSafe(b)
        For i = 0 To lenB - 1
            result(lenA + i) = b(baseB + i)
        Next i
    End If

    ByteArrayConcat = result
End Function

Private Function HexToBytes(ByVal hexStr As String) As Byte()
    Dim lengthBytes As Long
    lengthBytes = Len(hexStr) \ 2

    Dim result() As Byte
    If lengthBytes <= 0 Then
        HexToBytes = result
        Exit Function
    End If

    ReDim result(0 To lengthBytes - 1)

    Dim i As Long
    For i = 0 To lengthBytes - 1
        result(i) = CByte("&H" & Mid$(hexStr, 2 * i + 1, 2))
    Next i

    HexToBytes = result
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
    newLen = msgLen + 1 ' 0x80
    Do While (newLen Mod BLOCK_SIZE) <> (BLOCK_SIZE - 8)
        newLen = newLen + 1
    Loop
    newLen = newLen + 8 ' espaço pros 64 bits de comprimento

    ReDim padded(0 To newLen - 1) As Byte

    Dim i As Long
    For i = 0 To msgLen - 1
        padded(i) = data(i)
    Next i

    ' 0x80
    padded(msgLen) = &H80

    ' zeros no meio já default

    ' comprimento em bits (uint64 big-endian)
    Dim bitLen As Double
    bitLen = CDbl(msgLen) * 8#

    Dim hi As Double, lo As Double
    hi = Fix(bitLen / U32_MAX)
    lo = bitLen - hi * U32_MAX

    ' write high 32
    padded(newLen - 8) = D2L(Fix(hi / 16777216#)) And &HFF
    padded(newLen - 7) = D2L(Fix((hi Mod 16777216#) / 65536#)) And &HFF
    padded(newLen - 6) = D2L(Fix((hi Mod 65536#) / 256#)) And &HFF
    padded(newLen - 5) = D2L(hi Mod 256#) And &HFF
    ' write low 32
    padded(newLen - 4) = D2L(Fix(lo / 16777216#)) And &HFF
    padded(newLen - 3) = D2L(Fix((lo Mod 16777216#) / 65536#)) And &HFF
    padded(newLen - 2) = D2L(Fix((lo Mod 65536#) / 256#)) And &HFF
    padded(newLen - 1) = D2L(lo Mod 256#) And &HFF
End Sub

' ============================ Constantes K ==========================================
Private Sub K_Init(ByRef K() As Long)
    K(0) = &H428A2F98: K(1) = &H71374491: K(2) = &HB5C0FBCF: K(3) = &HE9B5DBA5
    K(4) = &H3956C25B: K(5) = &H59F111F1: K(6) = &H923F82A4: K(7) = &HAB1C5ED5
    K(8) = &HD807AA98: K(9) = &H12835B01: K(10) = &H243185BE: K(11) = &H550C7DC3
    K(12) = &H72BE5D74: K(13) = &H80DEB1FE: K(14) = &H9BDC06A7: K(15) = &HC19BF174
    K(16) = &HE49B69C1: K(17) = &HEFBE4786: K(18) = &HFC19DC6: K(19) = &H240CA1CC
    K(20) = &H2DE92C6F: K(21) = &H4A7484AA: K(22) = &H5CB0A9DC: K(23) = &H76F988DA
    K(24) = &H983E5152: K(25) = &HA831C66D: K(26) = &HB00327C8: K(27) = &HBF597FC7
    K(28) = &HC6E00BF3: K(29) = &HD5A79147: K(30) = &H6CA6351: K(31) = &H14292967
    K(32) = &H27B70A85: K(33) = &H2E1B2138: K(34) = &H4D2C6DFC: K(35) = &H53380D13
    K(36) = &H650A7354: K(37) = &H766A0ABB: K(38) = &H81C2C92E: K(39) = &H92722C85
    K(40) = &HA2BFE8A1: K(41) = &HA81A664B: K(42) = &HC24B8B70: K(43) = &HC76C51A3
    K(44) = &HD192E819: K(45) = &HD6990624: K(46) = &HF40E3585: K(47) = &H106AA070
    K(48) = &H19A4C116: K(49) = &H1E376C08: K(50) = &H2748774C: K(51) = &H34B0BCB5
    K(52) = &H391C0CB3: K(53) = &H4ED8AA4A: K(54) = &H5B9CCA4F: K(55) = &H682E6FF3
    K(56) = &H748F82EE: K(57) = &H78A5636F: K(58) = &H84C87814: K(59) = &H8CC70208
    K(60) = &H90BEFFFA: K(61) = &HA4506CEB: K(62) = &HBEF9A3F7: K(63) = &HC67178F2
End Sub

' ========================== Helpers de conversão ====================================
Private Function Word32ToHex(ByVal w As Long) As String
    Dim b0 As Long, b1 As Long, b2 As Long, b3 As Long
    b0 = URShift(w, 24) And &HFF
    b1 = URShift(w, 16) And &HFF
    b2 = URShift(w, 8) And &HFF
    b3 = w And &HFF
    Word32ToHex = _
        ByteToHex(b0) & ByteToHex(b1) & ByteToHex(b2) & ByteToHex(b3)
End Function

Private Function ByteToHex(ByVal x As Long) As String
    Dim s As String
    s = Hex$(x And &HFF)
    If Len(s) < 2 Then s = "0" & s
    ByteToHex = s
End Function

' ============================== Auto-Testes =========================================
Public Sub SHA256_SelfTest()
    Dim h As String
    
    h = SHA256_String("")
    Debug.Print "SHA256("""") = " & h
    Debug.Print "Expect       = E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855"
    
    h = SHA256_String("abc")
    Debug.Print "SHA256(""abc"") = " & h
    Debug.Print "Expect       = BA7816BF8F01CFEA414140DE5DAE2223B00361A396177A9CB410FF61F20015AD"
End Sub
