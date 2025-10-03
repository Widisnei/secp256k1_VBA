Attribute VB_Name = "Base58_VBA"
Option Explicit
Option Compare Binary
' =====================================================================================
'  Base58_VBA.bas — Base58 e Base58Check (Bitcoin) em VBA puro
'  Requer: SHA256_VBA.bas (para o checksum do Base58Check)
'
'  API pública (principais):
'    Base58_Encode(ByRef data() As Byte) As String
'    Base58_Decode(ByVal s As String) As Byte()        ' lança erro se inválido
'
'    Base58Check_Encode(version As Byte, ByRef payload() As Byte) As String
'    Base58Check_Decode(ByVal s As String, ByRef version As Byte, ByRef payload() As Byte) As Boolean
'
'  Auto-teste:
'    Call Base58_SelfTest
'
'  (c) 2025, MIT License.
' =====================================================================================

Private Const ALPH As String = "123456789ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz"

' ------------------------ Utilitários de conversão ----------------------------------
Private Function HexToBytes(ByVal hexStr As String) As Byte()
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

Private Function BytesToHex(ByRef data() As Byte) As String
    Dim i As Long, s As String
    If (Not Not data) = 0 Then BytesToHex = "": Exit Function
    For i = LBound(data) To UBound(data)
        s = s & Right$("0" & Hex$(data(i)), 2)
    Next
    BytesToHex = s
End Function

' ----------------------------- Base58 raw -------------------------------------------
Public Function Base58_Encode(ByRef data() As Byte) As String
    Dim zeros As Long, i As Long, j As Long
    Dim size As Long, carry As Long
    Dim arr() As Byte, tmp() As Byte

    If (Not Not data) = 0 Then
        Base58_Encode = ""
        Exit Function
    End If

    ' Conta prefixo zero (0x00) -> vira '1'
    For i = LBound(data) To UBound(data)
        If data(i) = 0 Then zeros = zeros + 1 Else Exit For
    Next

    ' Conversão por divisão repetida / 58
    ReDim arr(0 To UBound(data) - LBound(data)) As Byte
    For i = LBound(data) To UBound(data)
        arr(i - LBound(data)) = data(i)
    Next

    Dim enc As String: enc = ""
    Do While UBound(arr) >= 0 And Not (UBound(arr) = 0 And arr(0) = 0)
        carry = 0
        For i = 0 To UBound(arr)
            Dim cur As Long
            cur = carry * 256& + arr(i)
            arr(i) = CByte(cur \ 58&)
            carry = cur Mod 58&
        Next
        enc = Mid$(ALPH, carry + 1, 1) & enc

        ' remove leading zeros do quociente (robusto)
        Dim ub As Long
        On Error Resume Next
        ub = UBound(arr)
        If Err.Number <> 0 Then ub = -1: Err.Clear
        On Error GoTo 0

                i = 0
        If ub >= 0 Then
            Do While i <= ub
                If arr(i) <> 0 Then Exit Do
                i = i + 1
            Loop
        End If
        If i > 0 Then
            If i <= ub Then
                tmp = arr
                ReDim arr(0 To ub - i)
                For j = 0 To UBound(arr)
                    arr(j) = tmp(j + i)
                Next
            Else
                ReDim arr(0 To 0): arr(0) = 0
            End If
        End If
    Loop

    ' prefixo zeros -> '1'
    For i = 1 To zeros
        enc = "1" & enc
    Next
    Base58_Encode = enc
End Function

Public Function Base58_Decode(ByVal s As String) As Byte()
    Dim i As Long, j As Long, zeros As Long
    Dim arr() As Byte, tmp() As Byte

    If Len(s) = 0 Then Exit Function

    ' conta '1' prefixo -> 0x00 bytes
    For i = 1 To Len(s)
        If Mid$(s, i, 1) = "1" Then zeros = zeros + 1 Else Exit For
    Next

    ReDim arr(0 To Len(s) - 1) As Byte

    For i = 1 To Len(s)
        Dim ch As Long, val As Long, k As Long
        ch = Asc(Mid$(s, i, 1))

        val = -1
        For k = 1 To 58
            If Asc(Mid$(ALPH, k, 1)) = ch Then val = k - 1: Exit For
        Next
        If val = -1 Then Err.Raise 5, , "Base58: caractere inválido: " & Chr$(ch)

        ' arr = arr * 58 + val
        Dim carry As Long
        carry = val
        For j = UBound(arr) To 0 Step -1
            Dim cur As Long
            cur = arr(j) * 58& + carry
            arr(j) = CByte(cur And &HFF)
            carry = cur \ 256&
        Next
        If carry <> 0 Then
            ' aumenta tamanho e move
            tmp = arr
            ReDim arr(0 To UBound(tmp) + 1)
            arr(0) = CByte(carry)
            For j = 0 To UBound(tmp)
                arr(j + 1) = tmp(j)
            Next
        End If
    Next

    ' remove zeros à esquerda e injeta prefixo zeros (robusto p/ arrays vazios)
    Dim ub As Long
    On Error Resume Next
    ub = UBound(arr)
    If Err.Number <> 0 Then ub = -1: Err.Clear
    On Error GoTo 0

            i = 0
        If ub >= 0 Then
            Do While i <= ub
                If arr(i) <> 0 Then Exit Do
                i = i + 1
            Loop
        End If

    Dim out() As Byte, p As Long, coreLen As Long
    If ub >= i Then
        coreLen = ub - i + 1
    Else
        coreLen = 0
    End If

    If zeros + coreLen <= 0 Then
        ' vazio
        Erase out
    Else
        ReDim out(0 To zeros + coreLen - 1)
        For p = 0 To zeros - 1
            out(p) = 0
        Next
        For j = 0 To coreLen - 1
            out(p) = arr(i + j): p = p + 1
        Next
    End If
    Base58_Decode = out
End Function

' ----------------------------- Base58Check -------------------------------------------
Public Function Base58Check_Encode(ByVal version As Byte, ByRef payload() As Byte) As String
    Dim buf() As Byte, i As Long
    Dim n As Long: n = 1
    If (Not Not payload) <> 0 Then n = n + (UBound(payload) - LBound(payload) + 1)

    ReDim buf(0 To n - 1)
    buf(0) = version
    If n > 1 Then
        For i = 0 To (UBound(payload) - LBound(payload))
            buf(i + 1) = payload(LBound(payload) + i)
        Next
    End If

    ' checksum = 4 bytes do double-SHA256
    Dim chk() As Byte
    chk = DoubleSHA256_4(buf)

    Dim all() As Byte
    ReDim all(0 To UBound(buf) + 4)
    For i = 0 To UBound(buf)
        all(i) = buf(i)
    Next
    For i = 0 To 3
        all(UBound(buf) + 1 + i) = chk(i)
    Next

    Base58Check_Encode = Base58_Encode(all)
End Function

Public Function Base58Check_Decode(ByVal s As String, ByRef version As Byte, ByRef payload() As Byte) As Boolean
    On Error GoTo Fail
    Dim raw() As Byte
    raw = Base58_Decode(s)
    If (Not Not raw) = 0 Then GoTo Fail
    If UBound(raw) < 4 Then GoTo Fail

    Dim dataLen As Long: dataLen = UBound(raw) - 3
    Dim i As Long
    Dim data() As Byte
    ReDim data(0 To dataLen - 1)
    For i = 0 To dataLen - 1
        data(i) = raw(i)
    Next

    Dim gotChk(0 To 3) As Byte
    For i = 0 To 3
        gotChk(i) = raw(dataLen + i)
    Next

    Dim expChk() As Byte
    expChk = DoubleSHA256_4(data)

    For i = 0 To 3
        If gotChk(i) <> expChk(i) Then GoTo Fail
    Next

    version = data(0)
    If dataLen > 1 Then
        ReDim payload(0 To dataLen - 2)
        For i = 0 To dataLen - 2
            payload(i) = data(i + 1)
        Next
    End If
    Base58Check_Decode = True
    Exit Function
Fail:
    version = 0
    Erase payload
    Base58Check_Decode = False
End Function

' ------------------------- Helpers de checksum ---------------------------------------
Private Function DoubleSHA256_4(ByRef data() As Byte) As Byte()
    Dim hex1 As String, hex2 As String
    hex1 = SHA256_Bytes(data)
    Dim b1() As Byte
    b1 = HexToBytes(hex1)
    hex2 = SHA256_Bytes(b1)
    Dim b2() As Byte
    b2 = HexToBytes(hex2)
    Dim out4(0 To 3) As Byte
    out4(0) = b2(0): out4(1) = b2(1): out4(2) = b2(2): out4(3) = b2(3)
    DoubleSHA256_4 = out4
End Function

' -------------------------------- Auto-testes ----------------------------------------
Public Sub Base58_SelfTest()
    Dim payload() As Byte, ver As Byte, addr As String, ok As Boolean
    ' P2PKH fictício: version=0x00 e hash160 de "abc"
    Dim h160 As String
    h160 = Hash160_String("abc")
    payload = HexToBytes(h160)
    addr = Base58Check_Encode(&H0, payload)
    Debug.Print "P2PKH(abc) = "; addr

    ' Round-trip decode
    Dim ver2 As Byte, pay2() As Byte
    ok = Base58Check_Decode(addr, ver2, pay2)
    Debug.Print "Decode OK? "; ok, "version: "; ver2, "payload ==? "; (BytesToHex(pay2) = h160)
End Sub