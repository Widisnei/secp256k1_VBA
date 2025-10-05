Attribute VB_Name = "Bech32_VBA"
Option Explicit
Option Compare Binary
' =====================================================================================
'  Bech32_VBA.bas — Bech32 / Bech32m (BIP-0173 / BIP-0350) em VBA puro
'
'  API pública (principalmente para endereços SegWit):
'    Bech32_SegwitEncode(hrp As String, witver As Byte, ByRef prog() As Byte) As String
'    Bech32_SegwitDecode(addr As String, ByRef hrpOut As String, ByRef witver As Byte, ByRef prog() As Byte) As Boolean
'
'  Observações:
'   - Versão 0 usa checksum Bech32 (const = 1)
'   - Versão 1..16 usam checksum Bech32m (const = &H2BC830A3)
'   - Conversão de bits: 8->5 (para encode) e 5->8 (para decode) com padding controlado
'
'  Auto-teste:
'    Call Bech32_SelfTest
'
'  (c) 2025, MIT License.
' =====================================================================================

Private Const U32_MAX As Double = 4294967296#
Private Const U31 As Double = 2147483648#
Private Const GEN0 As Long = &H3B6A57B2
Private Const GEN1 As Long = &H26508E6D
Private Const GEN2 As Long = &H1EA119FA
Private Const GEN3 As Long = &H3D4233DD
Private Const GEN4 As Long = &H2A1462B3

Private Function U32(ByVal x As Long) As Double
    Dim d As Double: d = (x And &H7FFFFFFF)
    If (x And &H80000000) <> 0 Then d = d + U31
    U32 = d
End Function
Private Function D2L(ByVal u As Double) As Long
    If u >= U31 Then u = u - U32_MAX
    D2L = CLng(u)
End Function
Private Function URShift(ByVal x As Long, ByVal n As Long) As Long
    If n = 0 Then URShift = x: Exit Function
    URShift = D2L(Fix(U32(x) / (2 ^ n)))
End Function
Private Function ULShift(ByVal x As Long, ByVal n As Long) As Long
    If n = 0 Then ULShift = x: Exit Function
    Dim r As Double: r = (U32(x) * (2 ^ n))
    r = r - (Int(r / U32_MAX) * U32_MAX)
    ULShift = D2L(r)
End Function

Private Function Polymod(ByRef values() As Byte) As Long
    Dim c As Long: c = 1
    Dim i As Long
    For i = LBound(values) To UBound(values)
        Dim c0 As Long: c0 = URShift(c, 25)
        c = ULShift((c And &H1FFFFFF), 5) Xor values(i)
        If (c0 And 1) <> 0 Then c = c Xor GEN0
        If (c0 And 2) <> 0 Then c = c Xor GEN1
        If (c0 And 4) <> 0 Then c = c Xor GEN2
        If (c0 And 8) <> 0 Then c = c Xor GEN3
        If (c0 And 16) <> 0 Then c = c Xor GEN4
    Next
    Polymod = c
End Function

Private Function HRPExpand(ByVal hrp As String) As Byte()
    Dim i As Long, n As Long: n = Len(hrp)
    Dim out() As Byte: ReDim out(0 To 2 * n)
    For i = 1 To n: out(i - 1) = URShift(Asc(mid$(hrp, i, 1)), 5) And &H1F: Next
    out(n) = 0
    For i = 1 To n: out(n + i) = Asc(mid$(hrp, i, 1)) And &H1F: Next
    HRPExpand = out
End Function

Private Function CreateChecksum(ByVal hrp As String, ByRef data() As Byte, ByVal constv As Long) As Byte()
    Dim hrpex() As Byte, tmp() As Byte, i As Long, p As Long
    hrpex = HRPExpand(hrp)
    Dim dataLen As Long: If (Not Not data) = 0 Then dataLen = 0 Else dataLen = UBound(data) - LBound(data) + 1
    ReDim tmp(0 To (UBound(hrpex) - LBound(hrpex) + 1) + 1 + dataLen + 6 - 1)
    p = 0
    For i = LBound(hrpex) To UBound(hrpex): tmp(p) = hrpex(i): p = p + 1: Next
    tmp(p) = 0: p = p + 1
    If dataLen > 0 Then For i = 0 To dataLen - 1: tmp(p) = data(LBound(data) + i): p = p + 1: Next
    For i = 1 To 6: tmp(p) = 0: p = p + 1: Next
    Dim pm As Long: pm = Polymod(tmp) Xor constv
    Dim chk(0 To 5) As Byte
    For i = 0 To 5: chk(i) = URShift(pm, 5 * (5 - i)) And &H1F: Next
    CreateChecksum = chk
End Function

Private Function VerifyChecksum(ByVal hrp As String, ByRef data() As Byte) As Long
    Dim hrpex() As Byte, tmp() As Byte, i As Long, p As Long
    hrpex = HRPExpand(hrp)
    ReDim tmp(0 To (UBound(hrpex) - LBound(hrpex) + 1) + 1 + (UBound(data) - LBound(data) + 1) - 1)
    p = 0
    For i = LBound(hrpex) To UBound(hrpex): tmp(p) = hrpex(i): p = p + 1: Next
    tmp(p) = 0: p = p + 1
    For i = LBound(data) To UBound(data): tmp(p) = data(i): p = p + 1: Next
    Dim pm As Long: pm = Polymod(tmp)
    If pm = 1& Then
        VerifyChecksum = 1&
    ElseIf pm = &H2BC830A3 Then
        VerifyChecksum = &H2BC830A3
    Else
        VerifyChecksum = -1
    End If
End Function

Private Function ConvertBits(ByRef inb() As Byte, ByVal fromBits As Long, ByVal toBits As Long, ByVal pad As Boolean, ByRef outb() As Byte) As Boolean
    Dim acc As Long: acc = 0
    Dim bits As Long: bits = 0
    Dim maxv As Long: maxv = (2 ^ toBits) - 1
    Dim maxAcc As Long: maxAcc = (2 ^ (fromBits + toBits - 1)) - 1

    Dim i As Long, v As Long, p As Long
    Erase outb
    Dim tmp() As Byte: ReDim tmp(0 To 0): p = -1

    If (Not Not inb) <> 0 Then
        For i = LBound(inb) To UBound(inb)
            v = inb(i)
            If v < 0 Or v >= (2 ^ fromBits) Then ConvertBits = False: Exit Function
            acc = ((acc * (2 ^ fromBits)) Or v) And maxAcc
            bits = bits + fromBits
            Do While bits >= toBits
                bits = bits - toBits
                p = p + 1
                If p = 0 Then ReDim tmp(0 To 0) Else ReDim Preserve tmp(0 To p)
                tmp(p) = (acc \ (2 ^ bits)) And maxv
            Loop
        Next
    End If

    If pad And bits > 0 Then
        p = p + 1
        If p = 0 Then ReDim tmp(0 To 0) Else ReDim Preserve tmp(0 To p)
        tmp(p) = (acc * (2 ^ (toBits - bits))) And maxv
    ElseIf (bits >= fromBits) Or (((acc * (2 ^ (toBits - bits))) And maxv) <> 0) Then
        ConvertBits = False: Exit Function
    End If

    outb = tmp: ConvertBits = True
End Function

Public Function Bech32_SegwitEncode(ByVal hrp As String, ByVal witver As Byte, ByRef prog() As Byte) As String
    Dim data5() As Byte, payload() As Byte, chk() As Byte
    Dim ok As Boolean, i As Long, constv As Long
    Dim s As String
    Const ALPH As String = "qpzry9x8gf2tvdw0s3jn54khce6mua7l"
    Dim hrpLower As String, hrpUpper As String

    hrpLower = LCase$(hrp)
    hrpUpper = UCase$(hrp)
    If (hrp <> hrpLower) And (hrp <> hrpUpper) Then
        Bech32_SegwitEncode = ""
        Exit Function
    End If
    hrp = hrpLower

    If (Not Not prog) <> 0 Then
        ok = ConvertBits(prog, 8, 5, True, data5)
        If Not ok Then Bech32_SegwitEncode = "": Exit Function
    End If

    If (Not Not data5) <> 0 Then
        ReDim payload(0 To UBound(data5) + 1) ' FIX: +1 para acomodar witver
        payload(0) = witver
        For i = 0 To UBound(data5): payload(i + 1) = data5(i): Next
    Else
        ReDim payload(0 To 0): payload(0) = witver
    End If

    If witver = 0 Then constv = 1& Else constv = &H2BC830A3
    chk = CreateChecksum(hrp, payload, constv)

    s = hrp & "1"
    For i = 0 To UBound(payload): s = s & mid$(ALPH, payload(i) + 1, 1): Next
    For i = 0 To UBound(chk):     s = s & mid$(ALPH, chk(i) + 1, 1):     Next
    Bech32_SegwitEncode = s
End Function

Public Function Bech32_SegwitDecode(ByVal addr As String, ByRef hrpOut As String, ByRef witver As Byte, ByRef prog() As Byte) As Boolean
    Dim i As Long, p As Long, hrp As String, data() As Byte
    Dim c As String, val As Long
    Const ALPH As String = "qpzry9x8gf2tvdw0s3jn54khce6mua7l"

    Dim addrLower As String, addrUpper As String
    addrLower = LCase$(addr)
    addrUpper = UCase$(addr)
    If (addr <> addrLower) And (addr <> addrUpper) Then
        Bech32_SegwitDecode = False
        Exit Function
    End If

    addr = addrLower
    p = InStr(1, addr, "1")
    If p = 0 Or p < 1 Or p + 7 > Len(addr) Then Bech32_SegwitDecode = False: Exit Function
    hrp = left$(addr, p - 1)

    ReDim data(0 To Len(addr) - p - 1)
    For i = 1 To Len(addr) - p
        c = mid$(addr, p + i, 1)
        val = InStr(1, ALPH, c) - 1
        If val < 0 Then Bech32_SegwitDecode = False: Exit Function
        data(i - 1) = val
    Next

    Dim constv As Long: constv = VerifyChecksum(hrp, data)
    If constv = -1 Then Bech32_SegwitDecode = False: Exit Function

    If UBound(data) < 6 Then Bech32_SegwitDecode = False: Exit Function
    ReDim Preserve data(0 To UBound(data) - 6)
    If (Not Not data) = 0 Then Bech32_SegwitDecode = False: Exit Function

    witver = data(0): hrpOut = hrp

    Dim data5() As Byte, ok As Boolean
    If UBound(data) >= 1 Then
        ReDim data5(0 To UBound(data) - 1)
        For i = 1 To UBound(data): data5(i - 1) = data(i): Next
        ok = ConvertBits(data5, 5, 8, False, prog)
        If Not ok Then Bech32_SegwitDecode = False: Exit Function
    Else
        Erase prog
    End If

    If witver = 0 And constv <> 1& Then Bech32_SegwitDecode = False: Exit Function
    If witver <> 0 And constv <> &H2BC830A3 Then Bech32_SegwitDecode = False: Exit Function
    Bech32_SegwitDecode = True
End Function
' -------------------------------- Auto-testes ----------------------------------------
Public Sub Bech32_SelfTest()
    Dim hrp As String, witver As Byte, prog() As Byte, addr As String
    Dim ok As Boolean, hrp2 As String, ver2 As Byte, prog2() As Byte
    Dim i As Long
    hrp = "bc": witver = 0
    ReDim prog(0 To 19): For i = 0 To 19: prog(i) = 0: Next
    addr = Bech32_SegwitEncode(hrp, witver, prog)
    Debug.Print "addr = "; addr
    ok = Bech32_SegwitDecode(addr, hrp2, ver2, prog2)
    Debug.Print "decode OK? "; ok, "hrp="; hrp2, "ver="; ver2, "len(prog)="; (IIf((Not Not prog2) = 0, 0, UBound(prog2) + 1))
End Sub

' v1 (bech32m) – P2TR: programa de 32 bytes
Sub Teste_Bech32_v1_Taproot()
    Dim hrp As String, witver As Byte, prog() As Byte, addr As String
    Dim ok As Boolean, hrp2 As String, ver2 As Byte, prog2() As Byte, i As Long
    hrp = "bc": witver = 1
    ReDim prog(0 To 31): For i = 0 To 31: prog(i) = 0: Next
    addr = Bech32_SegwitEncode(hrp, witver, prog)
    Debug.Print "addr(v1,32B) = "; addr
    ok = Bech32_SegwitDecode(addr, hrp2, ver2, prog2)
    Debug.Print "decode OK? "; ok; "  hrp="; hrp2; "  ver="; ver2; "  len(prog)="; (UBound(prog2) + 1)
End Sub

' Testnet (tb)
Sub Teste_Bech32_Testnet()
    Dim hrp As String, witver As Byte, prog() As Byte, addr As String
    Dim ok As Boolean, hrp2 As String, ver2 As Byte, prog2() As Byte, i As Long
    hrp = "tb": witver = 0
    ReDim prog(0 To 19): For i = 0 To 19: prog(i) = i: Next
    addr = Bech32_SegwitEncode(hrp, witver, prog)
    Debug.Print "addr(tb,v0,20B) = "; addr
    ok = Bech32_SegwitDecode(addr, hrp2, ver2, prog2)
    Debug.Print "decode OK? "; ok; "  hrp="; hrp2; "  ver="; ver2; "  len(prog)="; (UBound(prog2) + 1)
End Sub