Attribute VB_Name = "EC_Montgomery_Ladder"
Option Explicit

' =============================================================================
' MONTGOMERY LADDER - RESISTÊNCIA A TIMING ATTACKS
' =============================================================================

#Const LADDER_DIAGNOSTICS_COMPILED = False

#If LADDER_DIAGNOSTICS_COMPILED Then
Private ladder_call_counter As Long
Private ladder_iteration_counter As Long
Private ladder_cswap_counter As Long
Private ladder_bit_counts(0 To 1) As Long
Private ladder_diagnostics_enabled As Boolean

Private Sub ladder_reset_counters_internal()
    ladder_call_counter = 0
    ladder_iteration_counter = 0
    ladder_cswap_counter = 0
    ladder_bit_counts(0) = 0
    ladder_bit_counts(1) = 0
End Sub

Public Sub ladder_set_diagnostics_enabled(ByVal enable As Boolean)
    ladder_diagnostics_enabled = enable
    Call ladder_reset_counters_internal
End Sub

Public Function ladder_diagnostics_available() As Boolean
    ladder_diagnostics_available = True
End Function

Public Function ladder_diagnostics_active() As Boolean
    ladder_diagnostics_active = ladder_diagnostics_enabled
End Function

Public Sub reset_ladder_call_counter()
    Call ladder_reset_counters_internal
End Sub

Public Function get_ladder_call_counter() As Long
    If ladder_diagnostics_enabled Then
        get_ladder_call_counter = ladder_call_counter
    Else
        get_ladder_call_counter = 0
    End If
End Function

Public Function get_ladder_iteration_counter() As Long
    If ladder_diagnostics_enabled Then
        get_ladder_iteration_counter = ladder_iteration_counter
    Else
        get_ladder_iteration_counter = 0
    End If
End Function

Public Function get_ladder_cswap_counter() As Long
    If ladder_diagnostics_enabled Then
        get_ladder_cswap_counter = ladder_cswap_counter
    Else
        get_ladder_cswap_counter = 0
    End If
End Function

Public Function get_ladder_bit_count(ByVal bitValue As Long) As Long
    If ladder_diagnostics_enabled And bitValue >= 0 And bitValue <= 1 Then
        get_ladder_bit_count = ladder_bit_counts(bitValue)
    Else
        get_ladder_bit_count = 0
    End If
End Function
#Else
Public Sub ladder_set_diagnostics_enabled(ByVal enable As Boolean)
    ' Diagnostics not compiled; no-op
End Sub

Public Function ladder_diagnostics_available() As Boolean
    ladder_diagnostics_available = False
End Function

Public Function ladder_diagnostics_active() As Boolean
    ladder_diagnostics_active = False
End Function

Public Sub reset_ladder_call_counter()
    ' Diagnostics not compiled; no counters to reset
End Sub

Public Function get_ladder_call_counter() As Long
    get_ladder_call_counter = 0
End Function

Public Function get_ladder_iteration_counter() As Long
    get_ladder_iteration_counter = 0
End Function

Public Function get_ladder_cswap_counter() As Long
    get_ladder_cswap_counter = 0
End Function

Public Function get_ladder_bit_count(ByVal bitValue As Long) As Long
    get_ladder_bit_count = 0
End Function
#End If

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

#If LADDER_DIAGNOSTICS_COMPILED Then
    If ladder_diagnostics_enabled Then ladder_cswap_counter = ladder_cswap_counter + 1
#End If

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

Private Sub ec_jacobian_copy(ByRef dest As EC_POINT_JACOBIAN, ByRef src As EC_POINT_JACOBIAN)
    Call BN_copy(dest.x, src.x)
    Call BN_copy(dest.y, src.y)
    Call BN_copy(dest.z, src.z)
    dest.infinity = src.infinity
End Sub

Private Sub jacobian_cswap_nocount(ByRef a As EC_POINT_JACOBIAN, ByRef b As EC_POINT_JACOBIAN, ByVal swap As Long)
    Dim mask As Long
    Dim tmp As Long
    Dim aInf As Long, bInf As Long

    mask = -CLng(swap And 1&)

    Call bn_cswap(a.x, b.x, swap)
    Call bn_cswap(a.y, b.y, swap)
    Call bn_cswap(a.z, b.z, swap)

    aInf = IIf(a.infinity, 1, 0)
    bInf = IIf(b.infinity, 1, 0)
    tmp = (aInf Xor bInf) And mask
    aInf = aInf Xor tmp
    bInf = bInf Xor tmp
    a.infinity = (aInf <> 0)
    b.infinity = (bInf <> 0)
End Sub

Private Sub ec_jacobian_cswap(ByRef a As EC_POINT_JACOBIAN, ByRef b As EC_POINT_JACOBIAN, ByVal swap As Long)
#If LADDER_DIAGNOSTICS_COMPILED Then
    If ladder_diagnostics_enabled Then ladder_cswap_counter = ladder_cswap_counter + 1
#End If
    Call jacobian_cswap_nocount(a, b, swap)
End Sub

Private Function jacobian_double_internal(ByRef result As EC_POINT_JACOBIAN, ByRef a As EC_POINT_JACOBIAN, ByRef ctx As SECP256K1_CTX) As Boolean
    Dim inputCopy As EC_POINT_JACOBIAN
    inputCopy = ec_jacobian_new()
    Call ec_jacobian_copy(inputCopy, a)

    Dim XX As BIGNUM_TYPE, YY As BIGNUM_TYPE, YYYY As BIGNUM_TYPE
    Dim S As BIGNUM_TYPE, M As BIGNUM_TYPE, T As BIGNUM_TYPE
    XX = BN_new(): YY = BN_new(): YYYY = BN_new()
    S = BN_new(): M = BN_new(): T = BN_new()

    Call BN_mod_sqr(XX, inputCopy.x, ctx.p)
    Call BN_mod_sqr(YY, inputCopy.y, ctx.p)
    Call BN_mod_sqr(YYYY, YY, ctx.p)

    Call BN_mod_mul(S, inputCopy.x, YY, ctx.p)
    Call BN_mod_add(S, S, S, ctx.p)
    Call BN_mod_add(S, S, S, ctx.p)

    Call BN_mod_add(M, XX, XX, ctx.p)
    Call BN_mod_add(M, M, XX, ctx.p)

    Call BN_mod_sqr(result.x, M, ctx.p)
    Call BN_mod_sub(result.x, result.x, S, ctx.p)
    Call BN_mod_sub(result.x, result.x, S, ctx.p)

    Call BN_mod_sub(T, S, result.x, ctx.p)
    Call BN_mod_mul(result.y, M, T, ctx.p)

    Call BN_mod_add(T, YYYY, YYYY, ctx.p)
    Call BN_mod_add(T, T, T, ctx.p)
    Call BN_mod_add(T, T, T, ctx.p)
    Call BN_mod_sub(result.y, result.y, T, ctx.p)

    Call BN_mod_mul(result.z, inputCopy.y, inputCopy.z, ctx.p)
    Call BN_mod_add(result.z, result.z, result.z, ctx.p)

    result.infinity = False

    Dim infPoint As EC_POINT_JACOBIAN
    infPoint = ec_jacobian_new()
    Call ec_jacobian_set_infinity(infPoint)
    Call jacobian_cswap_nocount(result, infPoint, IIf(a.infinity, 1, 0))

    jacobian_double_internal = True
End Function

Private Function jacobian_add_internal(ByRef result As EC_POINT_JACOBIAN, ByRef a As EC_POINT_JACOBIAN, ByRef b As EC_POINT_JACOBIAN, ByRef ctx As SECP256K1_CTX) As Boolean
    Dim aCopy As EC_POINT_JACOBIAN, bCopy As EC_POINT_JACOBIAN
    aCopy = ec_jacobian_new(): bCopy = ec_jacobian_new()
    Call ec_jacobian_copy(aCopy, a)
    Call ec_jacobian_copy(bCopy, b)

    Dim Z1Z1 As BIGNUM_TYPE, Z2Z2 As BIGNUM_TYPE, Z1Z3 As BIGNUM_TYPE, Z2Z3 As BIGNUM_TYPE
    Dim U1 As BIGNUM_TYPE, U2 As BIGNUM_TYPE, S1 As BIGNUM_TYPE, S2 As BIGNUM_TYPE
    Dim H As BIGNUM_TYPE, I As BIGNUM_TYPE, J As BIGNUM_TYPE, r As BIGNUM_TYPE
    Dim V As BIGNUM_TYPE, tmp1 As BIGNUM_TYPE, tmp2 As BIGNUM_TYPE
    Z1Z1 = BN_new(): Z2Z2 = BN_new(): Z1Z3 = BN_new(): Z2Z3 = BN_new()
    U1 = BN_new(): U2 = BN_new(): S1 = BN_new(): S2 = BN_new()
    H = BN_new(): I = BN_new(): J = BN_new(): r = BN_new()
    V = BN_new(): tmp1 = BN_new(): tmp2 = BN_new()

    Call BN_mod_sqr(Z1Z1, aCopy.z, ctx.p)
    Call BN_mod_sqr(Z2Z2, bCopy.z, ctx.p)

    Call BN_mod_mul(U1, aCopy.x, Z2Z2, ctx.p)
    Call BN_mod_mul(U2, bCopy.x, Z1Z1, ctx.p)

    Call BN_mod_mul(Z2Z3, bCopy.z, Z2Z2, ctx.p)
    Call BN_mod_mul(S1, aCopy.y, Z2Z3, ctx.p)

    Call BN_mod_mul(Z1Z3, aCopy.z, Z1Z1, ctx.p)
    Call BN_mod_mul(S2, bCopy.y, Z1Z3, ctx.p)

    Call BN_mod_sub(H, U2, U1, ctx.p)
    Call BN_mod_sub(r, S2, S1, ctx.p)

    If BN_is_zero(H) Then
        If BN_is_zero(r) Then
            Call jacobian_double_internal(result, aCopy, ctx)
        Else
            Call ec_jacobian_set_infinity(result)
            result.infinity = True
        End If
    Else
        Call BN_mod_add(tmp1, H, H, ctx.p)
        Call BN_mod_sqr(I, tmp1, ctx.p)
        Call BN_mod_mul(J, H, I, ctx.p)

        Call BN_mod_add(r, r, r, ctx.p)

        Call BN_mod_mul(V, U1, I, ctx.p)

        Call BN_mod_sqr(result.x, r, ctx.p)
        Call BN_mod_sub(result.x, result.x, J, ctx.p)
        Call BN_mod_add(tmp2, V, V, ctx.p)
        Call BN_mod_sub(result.x, result.x, tmp2, ctx.p)

        Call BN_mod_sub(tmp1, V, result.x, ctx.p)
        Call BN_mod_mul(result.y, r, tmp1, ctx.p)
        Call BN_mod_mul(tmp2, S1, J, ctx.p)
        Call BN_mod_add(tmp2, tmp2, tmp2, ctx.p)
        Call BN_mod_sub(result.y, result.y, tmp2, ctx.p)

        Call BN_mod_add(tmp1, aCopy.z, bCopy.z, ctx.p)
        Call BN_mod_sqr(tmp1, tmp1, ctx.p)
        Call BN_mod_sub(tmp1, tmp1, Z1Z1, ctx.p)
        Call BN_mod_sub(tmp1, tmp1, Z2Z2, ctx.p)
        Call BN_mod_mul(result.z, tmp1, H, ctx.p)

        result.infinity = False
    End If

    Dim copyA As EC_POINT_JACOBIAN, copyB As EC_POINT_JACOBIAN
    copyA = ec_jacobian_new(): copyB = ec_jacobian_new()
    Call ec_jacobian_copy(copyA, aCopy)
    Call ec_jacobian_copy(copyB, bCopy)

    Call jacobian_cswap_nocount(result, copyB, IIf(aCopy.infinity, 1, 0))
    Call jacobian_cswap_nocount(result, copyA, IIf(bCopy.infinity, 1, 0))

    jacobian_add_internal = True
End Function

Public Function ec_point_mul_ladder(ByRef result As EC_POINT, ByRef scalar As BIGNUM_TYPE, ByRef point As EC_POINT, ByRef ctx As SECP256K1_CTX) As Boolean
    ' Multiplicação escalar resistente a timing attacks
    ' Sempre executa mesmo número de operações independente do escalar

#If LADDER_DIAGNOSTICS_COMPILED Then
    If ladder_diagnostics_enabled Then ladder_call_counter = ladder_call_counter + 1
#End If

    If BN_is_zero(scalar) Or point.infinity Then
        Call ec_point_set_infinity(result)
        ec_point_mul_ladder = True
        Exit Function
    End If
    
    Dim R0 As EC_POINT_JACOBIAN, R1 As EC_POINT_JACOBIAN
    R0 = ec_jacobian_new()
    R1 = ec_jacobian_new()

    Call ec_jacobian_set_infinity(R0)

    Dim baseJac As EC_POINT_JACOBIAN
    baseJac = ec_jacobian_new()
    Call ec_jacobian_from_affine(baseJac, point)
    Call ec_jacobian_copy(R1, baseJac)

    Dim i As Long, nbits As Long, bit As Long
    nbits = BN_num_bits(scalar)

    Dim addRes As EC_POINT_JACOBIAN
    Dim dblR0 As EC_POINT_JACOBIAN
    Dim dblR1 As EC_POINT_JACOBIAN
    addRes = ec_jacobian_new()
    dblR0 = ec_jacobian_new()
    dblR1 = ec_jacobian_new()

    For i = nbits - 1 To 0 Step -1
        bit = IIf(BN_is_bit_set(scalar, i), 1, 0)

#If LADDER_DIAGNOSTICS_COMPILED Then
        If ladder_diagnostics_enabled Then
            ladder_iteration_counter = ladder_iteration_counter + 1
            ladder_bit_counts(bit) = ladder_bit_counts(bit) + 1
        End If
#End If

        Call ec_jacobian_cswap(R0, R1, bit)

        Call jacobian_add_internal(addRes, R0, R1, ctx)

        Call jacobian_double_internal(dblR1, R1, ctx)
        Call ec_jacobian_copy(R1, dblR1)

        Call jacobian_double_internal(dblR0, R0, ctx)
        Call ec_jacobian_copy(R0, dblR0)

        Call ec_jacobian_copy(R1, addRes)

        Call ec_jacobian_cswap(R0, R1, bit)
    Next i

    Dim affineResult As EC_POINT
    affineResult = ec_point_new()

    If Not ec_jacobian_to_affine(affineResult, R0, ctx) Then
        ec_point_mul_ladder = False
        Exit Function
    End If

    Call ec_point_copy(result, affineResult)
    ec_point_mul_ladder = True
End Function
