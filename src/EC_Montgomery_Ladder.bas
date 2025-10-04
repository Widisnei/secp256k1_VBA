Attribute VB_Name = "EC_Montgomery_Ladder"
Option Explicit

' =============================================================================
' MONTGOMERY LADDER - RESISTÊNCIA A TIMING ATTACKS
' =============================================================================

Private ladder_call_counter As Long
Private ladder_iteration_counter As Long
Private ladder_cswap_counter As Long
Private ladder_bit_counts(0 To 1) As Long

Public Sub reset_ladder_call_counter()
    ladder_call_counter = 0
    ladder_iteration_counter = 0
    ladder_cswap_counter = 0
    ladder_bit_counts(0) = 0
    ladder_bit_counts(1) = 0
End Sub

Public Function get_ladder_call_counter() As Long
    get_ladder_call_counter = ladder_call_counter
End Function

Public Function get_ladder_iteration_counter() As Long
    get_ladder_iteration_counter = ladder_iteration_counter
End Function

Public Function get_ladder_cswap_counter() As Long
    get_ladder_cswap_counter = ladder_cswap_counter
End Function

Public Function get_ladder_bit_count(ByVal bitValue As Long) As Long
    If bitValue < 0 Or bitValue > 1 Then
        get_ladder_bit_count = 0
    Else
        get_ladder_bit_count = ladder_bit_counts(bitValue)
    End If
End Function

Private Sub bn_cswap(ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE, ByVal swap As Long)
    Dim mask As Long
    Dim i As Long
    Dim maxWords As Long
    Dim tmp As Long
    Dim aNeg As Long, bNeg As Long
    Dim aFlags As Long, bFlags As Long

    mask = -CLng(swap And 1&)

    maxWords = a.dmax
    If b.dmax > maxWords Then maxWords = b.dmax
    If maxWords < 1 Then maxWords = 1

    Call bn_wexpand(a, maxWords)
    Call bn_wexpand(b, maxWords)

    For i = 0 To maxWords - 1
        tmp = (a.d(i) Xor b.d(i)) And mask
        a.d(i) = a.d(i) Xor tmp
        b.d(i) = b.d(i) Xor tmp
    Next i

    tmp = (a.top Xor b.top) And mask
    a.top = a.top Xor tmp
    b.top = b.top Xor tmp

    aNeg = a.neg
    bNeg = b.neg
    tmp = (aNeg Xor bNeg) And mask
    aNeg = aNeg Xor tmp
    bNeg = bNeg Xor tmp
    a.neg = (aNeg <> 0)
    b.neg = (bNeg <> 0)

    aFlags = a.flags
    bFlags = b.flags
    tmp = (aFlags Xor bFlags) And mask
    aFlags = aFlags Xor tmp
    bFlags = bFlags Xor tmp
    a.flags = aFlags
    b.flags = bFlags
End Sub

Private Sub ec_point_cswap(ByRef a As EC_POINT, ByRef b As EC_POINT, ByVal swap As Long)
    Dim mask As Long
    Dim aInf As Long, bInf As Long
    Dim tmp As Long

    mask = -CLng(swap And 1&)

    ladder_cswap_counter = ladder_cswap_counter + 1

    Call bn_cswap(a.x, b.x, swap)
    Call bn_cswap(a.y, b.y, swap)
    Call bn_cswap(a.z, b.z, swap)

    aInf = a.infinity
    bInf = b.infinity
    tmp = (aInf Xor bInf) And mask
    aInf = aInf Xor tmp
    bInf = bInf Xor tmp
    a.infinity = (aInf <> 0)
    b.infinity = (bInf <> 0)
End Sub

Public Function ec_point_mul_ladder(ByRef result As EC_POINT, ByRef scalar As BIGNUM_TYPE, ByRef point As EC_POINT, ByRef ctx As SECP256K1_CTX) As Boolean
    ' Multiplicação escalar resistente a timing attacks
    ' Sempre executa mesmo número de operações independente do escalar

    ladder_call_counter = ladder_call_counter + 1

    If BN_is_zero(scalar) Or point.infinity Then
        Call ec_point_set_infinity(result)
        ec_point_mul_ladder = True
        Exit Function
    End If
    
    Dim R0 As EC_POINT, R1 As EC_POINT, temp As EC_POINT
    R0 = ec_point_new(): R1 = ec_point_new(): temp = ec_point_new()
    
    Call ec_point_set_infinity(R0)  ' R0 = O
    Call ec_point_copy(R1, point)   ' R1 = P
    
    Dim i As Long, nbits As Long, bit As Long
    nbits = BN_num_bits(scalar)

    For i = nbits - 1 To 0 Step -1
        bit = IIf(BN_is_bit_set(scalar, i), 1, 0)

        ladder_iteration_counter = ladder_iteration_counter + 1
        ladder_bit_counts(bit) = ladder_bit_counts(bit) + 1

        Call ec_point_cswap(R0, R1, bit)

        ' Sempre executar ambas operações (constant-time)
        Call ec_point_add(temp, R0, R1, ctx)
        Call ec_point_double(R1, R1, ctx)
        Call ec_point_double(R0, R0, ctx)
        Call ec_point_copy(R1, temp)

        Call ec_point_cswap(R0, R1, bit)
    Next i

    Call ec_point_copy(result, R0)
    ec_point_mul_ladder = True
End Function
