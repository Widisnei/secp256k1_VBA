VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockBatchRNG"
Option Explicit

Private storedBytes() As Byte
Private hasPattern As Boolean

Public ShouldRaiseError As Boolean
Public RaiseErrorAfter As Long
Public CallCount As Long
Public ErrorCount As Long

Public ShouldReturnFalse As Boolean
Public ReturnFalseAfter As Long
Public FalseCount As Long

Public Sub ResetState()
    Call ClearPattern
    ShouldRaiseError = False
    RaiseErrorAfter = 0
    CallCount = 0
    ErrorCount = 0

    ShouldReturnFalse = False
    ReturnFalseAfter = 0
    FalseCount = 0
End Sub

Public Sub ClearPattern()
    Erase storedBytes
    hasPattern = False
End Sub

Public Sub SetFixedHex(ByVal hexString As String)
    Dim cleaned As String
    cleaned = Replace(hexString, " ", "")
    cleaned = Replace(cleaned, vbCr, "")
    cleaned = Replace(cleaned, vbLf, "")

    If Len(cleaned) = 0 Or (Len(cleaned) Mod 2) <> 0 Then
        Err.Raise vbObjectError + &H4301&, "MockBatchRNG.SetFixedHex", _
                  "A string hexadecimal deve conter um número par de dígitos."
    End If

    Dim byteCount As Long
    byteCount = Len(cleaned) \ 2
    ReDim storedBytes(0 To byteCount - 1)

    Dim i As Long
    For i = 0 To byteCount - 1
        storedBytes(i) = CByte(CLng("&H" & Mid$(cleaned, (i * 2) + 1, 2)))
    Next i

    hasPattern = True
End Sub

Public Function FillRandomBytes(buffer() As Byte) As Boolean
    CallCount = CallCount + 1

    If ShouldRaiseError Then
        If RaiseErrorAfter <= 0 Then
            ErrorCount = ErrorCount + 1
            Err.Raise vbObjectError + &H4300&, "MockBatchRNG.FillRandomBytes", _
                      "Erro simulado no provedor de aleatoriedade."
        ElseIf CallCount >= RaiseErrorAfter Then
            ErrorCount = ErrorCount + 1
            Err.Raise vbObjectError + &H4300&, "MockBatchRNG.FillRandomBytes", _
                      "Erro simulado no provedor de aleatoriedade."
        End If
    End If

    If ShouldReturnFalse Then
        If ReturnFalseAfter <= 0 Then
            FalseCount = FalseCount + 1
            FillRandomBytes = False
            Exit Function
        ElseIf CallCount >= ReturnFalseAfter Then
            FalseCount = FalseCount + 1
            FillRandomBytes = False
            Exit Function
        End If
    End If

    If Not hasPattern Then
        ReDim storedBytes(LBound(buffer) To UBound(buffer))

        Dim idx As Long
        For idx = LBound(storedBytes) To UBound(storedBytes)
            storedBytes(idx) = CByte(((idx - LBound(storedBytes)) * 37) And &HFF&)
        Next idx

        hasPattern = True
    End If

    Dim patternLen As Long
    patternLen = UBound(storedBytes) - LBound(storedBytes) + 1

    Dim outIndex As Long
    For outIndex = LBound(buffer) To UBound(buffer)
        buffer(outIndex) = storedBytes(((outIndex - LBound(buffer)) Mod patternLen) + LBound(storedBytes))
    Next outIndex

    FillRandomBytes = True
End Function
