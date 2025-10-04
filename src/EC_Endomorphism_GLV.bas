Attribute VB_Name = "EC_Endomorphism_GLV"
Option Explicit

' =============================================================================
' EC ENDOMORPHISM GLV - MÉTODO GLV PARA SECP256K1
' =============================================================================
' Implementa endomorphism específico da curva secp256k1 para 40-50% melhoria
' Baseado nas propriedades especiais: β³ ≡ 1 (mod p), λ³ ≡ 1 (mod n)
' =============================================================================

' Constantes do endomorphism secp256k1
Private Const BETA_HEX As String = "7AE96A2B657C07106E64479EAC3434E99CF0497512F58995C1396C28719501EE"
Private Const LAMBDA_HEX As String = "5363AD4CC05C30E0A5261C028812645A122E22EA20816678DF02967C1B23BD72"
Private Const GLV_A1_HEX As String = "3086D221A7D46BCDE86C90E49284EB15"
Private Const GLV_B1_HEX As String = "E4437ED6010E88286F547FA90ABFE4C3"
Private Const GLV_A2_HEX As String = "114CA50F7A8E2F3F657C1108D9D44CFD8"
Private Const GLV_B2_HEX As String = "3086D221A7D46BCDE86C90E49284EB15"
Private Const GLV_SQRT_N_HEX As String = "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF"

Public Function ec_point_mul_glv(ByRef result As EC_POINT, ByRef scalar As BIGNUM_TYPE, ByRef point As EC_POINT, ByRef ctx As SECP256K1_CTX) As Boolean
    ' Multiplicação escalar usando método GLV - 40-50% mais rápido

    If require_constant_time() Then
        ec_point_mul_glv = ec_point_mul_ladder(result, scalar, point, ctx)
        Exit Function
    End If

    ' Decompor escalar: k = k1 + k2*λ onde |k1|,|k2| ≈ √n
    Dim k1 As BIGNUM_TYPE, k2 As BIGNUM_TYPE
    k1 = BN_new(): k2 = BN_new()

    Call glv_decompose_scalar(k1, k2, scalar, ctx)

    Dim k1_abs As BIGNUM_TYPE, k2_abs As BIGNUM_TYPE
    k1_abs = BN_new(): k2_abs = BN_new()
    Call BN_copy(k1_abs, k1)
    Call BN_copy(k2_abs, k2)
    k1_abs.neg = False
    k2_abs.neg = False
    
    ' Calcular ponto endomorphism: P2 = β*P
    Dim point2 As EC_POINT
    point2 = ec_point_new()
    Call apply_endomorphism(point2, point, ctx)

    Dim p1_local As EC_POINT, p2_local As EC_POINT
    p1_local = ec_point_new()
    p2_local = ec_point_new()
    Call ec_point_copy(p1_local, point)
    Call ec_point_copy(p2_local, point2)

    If k1.neg Then
        If Not ec_point_negate(p1_local, p1_local, ctx) Then
            ec_point_mul_glv = False
            Exit Function
        End If
    End If

    If k2.neg Then
        If Not ec_point_negate(p2_local, p2_local, ctx) Then
            ec_point_mul_glv = False
            Exit Function
        End If
    End If

    ' Calcular k1*P + k2*P2 usando Strauss com escalares não negativos
    ec_point_mul_glv = ec_point_mul_strauss(result, k1_abs, p1_local, k2_abs, p2_local, ctx)
End Function

Private Sub glv_decompose_scalar(ByRef k1 As BIGNUM_TYPE, ByRef k2 As BIGNUM_TYPE, ByRef k As BIGNUM_TYPE, ByRef ctx As SECP256K1_CTX)
    ' Decompõe escalar k = k1 + k2*λ com |k1|,|k2| ≈ √n
    ' Usa algoritmo de Babai para encontrar vetor curto na lattice
    
    Dim a1 As BIGNUM_TYPE, a2 As BIGNUM_TYPE
    Dim mu1 As BIGNUM_TYPE, mu2 As BIGNUM_TYPE
    Dim sqrt_n As BIGNUM_TYPE
    Dim half_n As BIGNUM_TYPE
    Dim k_mod_n As BIGNUM_TYPE
    Dim c1 As BIGNUM_TYPE, c2 As BIGNUM_TYPE
    Dim numerator As BIGNUM_TYPE
    Dim prod1 As BIGNUM_TYPE, prod2 As BIGNUM_TYPE, sum_prod As BIGNUM_TYPE
    Dim mu_prod1 As BIGNUM_TYPE, mu_prod2 As BIGNUM_TYPE

    a1 = BN_hex2bn(GLV_A1_HEX)
    a2 = BN_hex2bn(GLV_A2_HEX)

    mu1 = BN_hex2bn(GLV_B1_HEX)
    mu1.neg = True

    mu2 = BN_hex2bn(GLV_B2_HEX)
    mu2.neg = True

    sqrt_n = BN_hex2bn(GLV_SQRT_N_HEX)

    half_n = BN_new()
    Call BN_copy(half_n, ctx.n)
    Call BN_rshift(half_n, half_n, 1)

    k_mod_n = BN_new()
    Call BN_mod(k_mod_n, k, ctx.n)

    c1 = BN_new()
    c2 = BN_new()
    numerator = BN_new()

    If Not BN_mul(numerator, k_mod_n, mu1) Then GoTo GLV_FAIL
    If Not rounded_division(c1, numerator, ctx.n, half_n) Then GoTo GLV_FAIL

    If Not BN_mul(numerator, k_mod_n, mu2) Then GoTo GLV_FAIL
    If Not rounded_division(c2, numerator, ctx.n, half_n) Then GoTo GLV_FAIL

    prod1 = BN_new()
    prod2 = BN_new()
    sum_prod = BN_new()

    If Not BN_mul(prod1, c1, a1) Then GoTo GLV_FAIL
    If Not BN_mul(prod2, c2, a2) Then GoTo GLV_FAIL
    If Not BN_add(sum_prod, prod1, prod2) Then GoTo GLV_FAIL
    If Not BN_sub(k1, k_mod_n, sum_prod) Then GoTo GLV_FAIL

    mu_prod1 = BN_new()
    mu_prod2 = BN_new()

    If Not BN_mul(mu_prod1, c1, mu1) Then GoTo GLV_FAIL
    If Not BN_mul(mu_prod2, c2, mu2) Then GoTo GLV_FAIL
    If Not BN_add(k2, mu_prod1, mu_prod2) Then GoTo GLV_FAIL

    If Not BN_mod(k1, k1, ctx.n) Then GoTo GLV_FAIL
    If Not BN_mod(k2, k2, ctx.n) Then GoTo GLV_FAIL

    If BN_cmp(k1, half_n) > 0 Then
        If Not BN_sub(k1, k1, ctx.n) Then GoTo GLV_FAIL
    End If

    If BN_cmp(k2, half_n) > 0 Then
        If Not BN_sub(k2, k2, ctx.n) Then GoTo GLV_FAIL
    End If

    Dim k1_abs As BIGNUM_TYPE
    Dim k2_abs As BIGNUM_TYPE
    k1_abs = BN_new()
    k2_abs = BN_new()
    Call BN_copy(k1_abs, k1)
    Call BN_copy(k2_abs, k2)
    k1_abs.neg = False
    k2_abs.neg = False

    If BN_cmp(k1_abs, sqrt_n) > 0 Or BN_cmp(k2_abs, sqrt_n) > 0 Then GoTo GLV_FAIL
    Exit Sub

GLV_FAIL:
    Call BN_zero(k1)
    Call BN_zero(k2)
End Sub

Public Function glv_decompose_scalar_for_tests(ByRef k1 As BIGNUM_TYPE, ByRef k2 As BIGNUM_TYPE, ByRef k As BIGNUM_TYPE, ByRef ctx As SECP256K1_CTX) As Boolean
    ' Exposta apenas para testes: delega para glv_decompose_scalar para inspecionar k1/k2
    Call glv_decompose_scalar(k1, k2, k, ctx)
    glv_decompose_scalar_for_tests = True
End Function

Private Sub reduce_to_sqrt_range(ByRef value As BIGNUM_TYPE, ByRef sqrt_n As BIGNUM_TYPE, ByRef half_sqrt As BIGNUM_TYPE)
    Dim quotient As BIGNUM_TYPE, remainder As BIGNUM_TYPE
    Dim adjusted As BIGNUM_TYPE, abs_value As BIGNUM_TYPE
    Dim correction As BIGNUM_TYPE

    quotient = BN_new()
    remainder = BN_new()
    adjusted = BN_new()
    abs_value = BN_new()
    correction = BN_new()

    If value.neg Then
        Call BN_copy(abs_value, value)
        abs_value.neg = False
        If Not BN_add(adjusted, abs_value, half_sqrt) Then GoTo RANGE_FAIL
        If Not BN_div(quotient, remainder, adjusted, sqrt_n) Then GoTo RANGE_FAIL
        If Not BN_is_zero(quotient) Then quotient.neg = True
    Else
        If Not BN_add(adjusted, value, half_sqrt) Then GoTo RANGE_FAIL
        If Not BN_div(quotient, remainder, adjusted, sqrt_n) Then GoTo RANGE_FAIL
    End If

    If Not BN_mul(correction, quotient, sqrt_n) Then GoTo RANGE_FAIL
    If Not BN_sub(value, value, correction) Then GoTo RANGE_FAIL
    Exit Sub

RANGE_FAIL:
    Call BN_zero(value)
End Sub

Private Function rounded_division(ByRef result As BIGNUM_TYPE, ByRef numerator As BIGNUM_TYPE, ByRef denominator As BIGNUM_TYPE, ByRef half_den As BIGNUM_TYPE) As Boolean
    Dim adjusted As BIGNUM_TYPE
    Dim abs_num As BIGNUM_TYPE
    Dim quotient As BIGNUM_TYPE
    Dim remainder As BIGNUM_TYPE

    adjusted = BN_new()
    abs_num = BN_new()
    quotient = BN_new()
    remainder = BN_new()

    If numerator.neg Then
        Call BN_copy(abs_num, numerator)
        abs_num.neg = False
        If Not BN_add(adjusted, abs_num, half_den) Then GoTo ROUND_FAIL
        If Not BN_div(quotient, remainder, adjusted, denominator) Then GoTo ROUND_FAIL
        Call BN_copy(result, quotient)
        If Not BN_is_zero(result) Then result.neg = True Else result.neg = False
    Else
        If Not BN_add(adjusted, numerator, half_den) Then GoTo ROUND_FAIL
        If Not BN_div(quotient, remainder, adjusted, denominator) Then GoTo ROUND_FAIL
        Call BN_copy(result, quotient)
        result.neg = False
    End If

    rounded_division = True
    Exit Function

ROUND_FAIL:
    Call BN_zero(result)
    result.neg = False
End Function

Private Sub apply_endomorphism(ByRef result As EC_POINT, ByRef point As EC_POINT, ByRef ctx As SECP256K1_CTX)
    ' Aplica endomorphism: (x,y) → (β*x, y)
    Dim beta As BIGNUM_TYPE
    beta = BN_hex2bn(BETA_HEX)
    
    Call BN_mod_mul(result.x, point.x, beta, ctx.p)
    Call BN_copy(result.y, point.y)
    Call BN_set_word(result.z, 1)
    result.infinity = point.infinity
End Sub

Private Function ec_point_mul_strauss(ByRef result As EC_POINT, ByRef k1 As BIGNUM_TYPE, ByRef p1 As EC_POINT, ByRef k2 As BIGNUM_TYPE, ByRef p2 As EC_POINT, ByRef ctx As SECP256K1_CTX) As Boolean
    ' Algoritmo de Strauss para k1*P1 + k2*P2
    Dim max_bits As Long, i As Long
    max_bits = IIf(BN_num_bits(k1) > BN_num_bits(k2), BN_num_bits(k1), BN_num_bits(k2))
    
    Call ec_point_set_infinity(result)
    
    For i = max_bits - 1 To 0 Step -1
        Call ec_point_double(result, result, ctx)
        
        If BN_is_bit_set(k1, i) And BN_is_bit_set(k2, i) Then
            Dim temp As EC_POINT: temp = ec_point_new()
            Call ec_point_add(temp, p1, p2, ctx)
            Call ec_point_add(result, result, temp, ctx)
        ElseIf BN_is_bit_set(k1, i) Then
            Call ec_point_add(result, result, p1, ctx)
        ElseIf BN_is_bit_set(k2, i) Then
            Call ec_point_add(result, result, p2, ctx)
        End If
    Next i
    
    ec_point_mul_strauss = True
End Function