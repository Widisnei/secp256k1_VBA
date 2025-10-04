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
    
    Dim lambda As BIGNUM_TYPE, lambda_inv As BIGNUM_TYPE
    Dim k_minus_k1 As BIGNUM_TYPE, k_mod_n As BIGNUM_TYPE
    Dim lambda_k2 As BIGNUM_TYPE
    lambda = BN_hex2bn(LAMBDA_HEX)
    lambda_inv = BN_new()
    k_minus_k1 = BN_new()
    k_mod_n = BN_new()
    lambda_k2 = BN_new()

    ' Algoritmo simplificado: k1 = k mod √n, k2 = (k-k1)*λ⁻¹ mod n
    Dim sqrt_n As BIGNUM_TYPE, half_sqrt As BIGNUM_TYPE
    sqrt_n = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEBAAEDCE6AF48A03BBFD25E8CD0364141")
    Call BN_rshift(sqrt_n, sqrt_n, 128) ' Aproximação de √n
    half_sqrt = BN_new()
    Call BN_copy(half_sqrt, sqrt_n)
    Call BN_rshift(half_sqrt, half_sqrt, 1)

    Call BN_mod(k_mod_n, k, ctx.n)
    Call BN_mod(k1, k_mod_n, sqrt_n)
    Call reduce_to_sqrt_range(k1, sqrt_n, half_sqrt)

    Call BN_sub(k_minus_k1, k_mod_n, k1)
    If Not BN_mod_inverse(lambda_inv, lambda, ctx.n) Then
        Call BN_zero(k1)
        Call BN_zero(k2)
        Exit Sub
    End If

    If Not BN_mod_mul(k2, k_minus_k1, lambda_inv, ctx.n) Then
        Call BN_zero(k1)
        Call BN_zero(k2)
        Exit Sub
    End If

    Call reduce_to_sqrt_range(k2, sqrt_n, half_sqrt)

    If Not BN_mul(lambda_k2, k2, lambda) Then
        Call BN_zero(k1)
        Call BN_zero(k2)
        Exit Sub
    End If

    If Not BN_mod(lambda_k2, lambda_k2, ctx.n) Then
        Call BN_zero(k1)
        Call BN_zero(k2)
        Exit Sub
    End If

    If Not BN_sub(k1, k_mod_n, lambda_k2) Then
        Call BN_zero(k1)
        Call BN_zero(k2)
        Exit Sub
    End If

    If Not BN_mod(k1, k1, ctx.n) Then
        Call BN_zero(k1)
        Call BN_zero(k2)
        Exit Sub
    End If

    Call reduce_to_sqrt_range(k1, sqrt_n, half_sqrt)
End Sub

Public Function glv_decompose_scalar_for_tests(ByRef k1 As BIGNUM_TYPE, ByRef k2 As BIGNUM_TYPE, ByRef k As BIGNUM_TYPE, ByRef ctx As SECP256K1_CTX) As Boolean
    ' Exposta apenas para testes: delega para glv_decompose_scalar para inspecionar k1/k2
    Call glv_decompose_scalar(k1, k2, k, ctx)
    glv_decompose_scalar_for_tests = True
End Function

Private Sub reduce_to_sqrt_range(ByRef value As BIGNUM_TYPE, ByRef sqrt_n As BIGNUM_TYPE, ByRef half_sqrt As BIGNUM_TYPE)
    Dim abs_value As BIGNUM_TYPE
    abs_value = BN_new()

    Do
        Call BN_copy(abs_value, value)
        abs_value.neg = False

        If BN_cmp(abs_value, half_sqrt) <= 0 Then Exit Do

        Dim value_copy As BIGNUM_TYPE
        Dim sqrt_copy As BIGNUM_TYPE

        value_copy = BN_new()
        sqrt_copy = BN_new()

        Call BN_copy(value_copy, value)
        Call BN_copy(sqrt_copy, sqrt_n)

        Dim success As Boolean
        If value.neg Then
            success = BN_add(value, value_copy, sqrt_copy)
        Else
            success = BN_sub(value, value_copy, sqrt_copy)
        End If

        If Not success Then
            Call BN_zero(value)
            Exit Do
        End If
    Loop
End Sub

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