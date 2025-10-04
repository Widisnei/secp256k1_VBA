Attribute VB_Name = "EC_Multiplication_Dispatch"
Option Explicit

' =============================================================================
' EC MULTIPLICATION DISPATCH - ESCOLHA AUTOMÁTICA DE ALGORITMOS
' =============================================================================
' Responsável por selecionar a melhor estratégia de multiplicação escalar
' considerando tamanho do escalar, segurança (constant-time) e heurísticas
' de uso de pontos. Implementação derivada dos testes de performance, com
' dependências ajustadas para o ambiente de produção.
' =============================================================================

Private Const CACHE_USAGE_THRESHOLD As Long = 3

Private Type POINT_USAGE_ENTRY
    signature As String
    hits As Long
End Type

Private point_usage_initialized As Boolean
Private point_usage() As POINT_USAGE_ENTRY
Private next_usage_slot As Long

' =============================================================================
' MULTIPLICAÇÃO ESCALAR "ULTIMATE"
' =============================================================================

Public Function ec_point_mul_ultimate(ByRef result As EC_POINT, ByRef scalar As BIGNUM_TYPE, ByRef point As EC_POINT, ByRef ctx As SECP256K1_CTX) As Boolean
    Dim scalar_bits As Long
    scalar_bits = BN_num_bits(scalar)

    If require_constant_time() Then
        ec_point_mul_ultimate = ec_point_mul_ladder(result, scalar, point, ctx)
        Exit Function
    End If

    If is_generator_point(point, ctx) Then
        If EC_Precomputed_Manager.use_precomputed_gen_tables() Then
            ec_point_mul_ultimate = ec_generator_mul_precomputed_naf(result, scalar, ctx)
        Else
            ec_point_mul_ultimate = EC_Precomputed_Manager.ec_generator_mul_fast(result, scalar, ctx)
        End If
    ElseIf scalar_bits > 200 Then
        ec_point_mul_ultimate = ec_point_mul_glv(result, scalar, point, ctx)
    ElseIf scalar_bits > 128 Then
        If require_constant_time() Then
            ec_point_mul_ultimate = ec_point_mul_ladder(result, scalar, point, ctx)
        ElseIf should_use_cached_point(point) Then
            ec_point_mul_ultimate = ec_point_mul_cached(result, scalar, point, ctx)
        Else
            ec_point_mul_ultimate = ec_point_mul_sliding_naf(result, scalar, point, ctx)
        End If
    Else
        ec_point_mul_ultimate = ec_point_mul(result, scalar, point, ctx)
    End If
End Function

' =============================================================================
' EXPONENCIAÇÃO MODULAR AUTOMÁTICA
' =============================================================================

Public Function BN_mod_exp_ultimate(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef e As BIGNUM_TYPE, ByRef m As BIGNUM_TYPE) As Boolean
    Dim exp_bits As Long
    exp_bits = BN_num_bits(e)

    If is_secp256k1_prime(m) And exp_bits > 64 Then
        Dim mont_ctx As MONT_CTX
        mont_ctx = BN_MONT_CTX_new()
        If Not BN_MONT_CTX_set(mont_ctx, m) Then
            BN_mod_exp_ultimate = False
            Exit Function
        End If
        BN_mod_exp_ultimate = BN_mod_exp_mont(r, a, e, m, mont_ctx)
    ElseIf exp_bits > 128 Then
        BN_mod_exp_ultimate = BN_mod_exp_win4(r, a, e, m)
    Else
        BN_mod_exp_ultimate = BN_mod_exp(r, a, e, m)
    End If
End Function

' =============================================================================
' HELPER FUNCTIONS
' =============================================================================

Private Function is_generator_point(ByRef point As EC_POINT, ByRef ctx As SECP256K1_CTX) As Boolean
    If point.infinity Then
        is_generator_point = False
    Else
        is_generator_point = (ec_point_cmp(point, ctx.g, ctx) = 0)
    End If
End Function

Private Function should_use_cached_point(ByRef point As EC_POINT) As Boolean
    If point.infinity Then
        should_use_cached_point = False
        Exit Function
    End If

    If Not point_usage_initialized Then
        ReDim point_usage(0 To 7)
        point_usage_initialized = True
        next_usage_slot = 0
    End If

    Dim signature As String
    signature = point_signature(point)

    Dim i As Long
    For i = LBound(point_usage) To UBound(point_usage)
        If point_usage(i).signature = signature Then
            point_usage(i).hits = point_usage(i).hits + 1
            should_use_cached_point = (point_usage(i).hits >= CACHE_USAGE_THRESHOLD)
            Exit Function
        End If
    Next i

    For i = LBound(point_usage) To UBound(point_usage)
        If point_usage(i).signature = "" Then
            point_usage(i).signature = signature
            point_usage(i).hits = 1
            should_use_cached_point = False
            Exit Function
        End If
    Next i

    point_usage(next_usage_slot).signature = signature
    point_usage(next_usage_slot).hits = 1
    next_usage_slot = (next_usage_slot + 1) Mod (UBound(point_usage) + 1)

    should_use_cached_point = False
End Function

Private Function point_signature(ByRef point As EC_POINT) As String
    If point.infinity Then
        point_signature = "INF"
    Else
        point_signature = Left$(BN_bn2hex(point.x), 32) & ":" & Left$(BN_bn2hex(point.y), 32)
    End If
End Function

Private Function is_secp256k1_prime(ByRef m As BIGNUM_TYPE) As Boolean
    Static secp_p As BIGNUM_TYPE
    Static initialized As Boolean

    If Not initialized Then
        secp_p = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")
        initialized = True
    End If

    is_secp256k1_prime = (BN_cmp(m, secp_p) = 0)
End Function

