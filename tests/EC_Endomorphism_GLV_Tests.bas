Attribute VB_Name = "EC_Endomorphism_GLV_Tests"
Option Explicit
Option Compare Binary
Option Base 0

'==============================================================================
' TESTES: Endomorphism GLV
'------------------------------------------------------------------------------
' Garante que ec_point_mul_glv produz o mesmo resultado que ec_point_mul para
' pontos arbitrários da curva secp256k1.
'==============================================================================

Public Sub Run_Endomorphism_GLV_Tests()
    Debug.Print "=== Testes Endomorphism GLV ==="

    Call secp256k1_init
    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()

    Randomize 20241005

    Dim iterations As Long
    iterations = 16

    Dim i As Long
    Dim passed As Long, total As Long

    For i = 1 To iterations
        total = total + 1

        Dim scalar As BIGNUM_TYPE
        Dim base_scalar As BIGNUM_TYPE
        Dim base_point As EC_POINT
        Dim reference As EC_POINT
        Dim glv_result As EC_POINT

        scalar = random_scalar_mod_n(ctx)
        base_scalar = random_scalar_mod_n(ctx)

        base_point = ec_point_new()
        reference = ec_point_new()
        glv_result = ec_point_new()

        If Not ec_point_mul(base_point, base_scalar, ctx.g, ctx) Then
            Debug.Print "FALHOU: geração do ponto base aleatório (iteração " & i & ")"
            GoTo NextIteration
        End If

        If Not ec_point_mul(reference, scalar, base_point, ctx) Then
            Debug.Print "FALHOU: multiplicação de referência (iteração " & i & ")"
            GoTo NextIteration
        End If

        If Not ec_point_mul_glv(glv_result, scalar, base_point, ctx) Then
            Debug.Print "FALHOU: multiplicação GLV (iteração " & i & ")"
            GoTo NextIteration
        End If

        If ec_point_cmp(reference, glv_result, ctx) = 0 Then
            passed = passed + 1
        Else
            Debug.Print "FALHOU: divergência entre GLV e referência (iteração " & i & ")"
        End If
NextIteration:
    Next i

    Dim negative_cases As Long
    Dim negative_target As Long
    Dim attempts As Long
    Dim max_attempts As Long

    negative_cases = 0
    negative_target = 4
    attempts = 0
    max_attempts = 256

    Do While negative_cases < negative_target And attempts < max_attempts
        attempts = attempts + 1

        Dim neg_scalar As BIGNUM_TYPE
        Dim k1_dec As BIGNUM_TYPE
        Dim k2_dec As BIGNUM_TYPE

        neg_scalar = random_scalar_mod_n(ctx)
        k1_dec = BN_new()
        k2_dec = BN_new()

        If glv_decompose_scalar_for_tests(k1_dec, k2_dec, neg_scalar, ctx) Then
            If k1_dec.neg Or k2_dec.neg Then
                total = total + 1

                base_scalar = random_scalar_mod_n(ctx)
                base_point = ec_point_new()
                reference = ec_point_new()
                glv_result = ec_point_new()

                If Not ec_point_mul(base_point, base_scalar, ctx.g, ctx) Then
                    Debug.Print "FALHOU: geração do ponto base (caso negativo " & negative_cases + 1 & ")"
                    GoTo NextNegativeIteration
                End If

                If Not ec_point_mul(reference, neg_scalar, base_point, ctx) Then
                    Debug.Print "FALHOU: multiplicação de referência (caso negativo " & negative_cases + 1 & ")"
                    GoTo NextNegativeIteration
                End If

                If Not ec_point_mul_glv(glv_result, neg_scalar, base_point, ctx) Then
                    Debug.Print "FALHOU: multiplicação GLV (caso negativo " & negative_cases + 1 & ")"
                    GoTo NextNegativeIteration
                End If

                If ec_point_cmp(reference, glv_result, ctx) = 0 Then
                    passed = passed + 1
                    negative_cases = negative_cases + 1
                Else
                    Debug.Print "FALHOU: divergência GLV em decomposição negativa (caso " & negative_cases + 1 & ")"
                End If
            End If
        End If

NextNegativeIteration:
    Loop

    If negative_cases < negative_target Then
        Debug.Print "AVISO: apenas " & negative_cases & " casos com decomposição negativa confirmados em " & attempts & " tentativas"
    End If

    Debug.Print "Resultado GLV vs referência: " & passed & " / " & total & " confirmados"

    Dim regression_total As Long
    Dim regression_passed As Long
    regression_total = 24

    For i = 1 To regression_total
        scalar = random_scalar_mod_n(ctx)

        Dim dec_k1 As BIGNUM_TYPE
        Dim dec_k2 As BIGNUM_TYPE
        dec_k1 = BN_new()
        dec_k2 = BN_new()

        If Not glv_decompose_scalar_for_tests(dec_k1, dec_k2, scalar, ctx) Then
            Debug.Print "FALHOU: decomposição GLV falhou (regressão " & i & ")"
            GoTo NextRegression
        End If

        Dim sqrt_bound As BIGNUM_TYPE
        sqrt_bound = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF")

        Dim k1_abs As BIGNUM_TYPE
        Dim k2_abs As BIGNUM_TYPE
        k1_abs = BN_new()
        k2_abs = BN_new()
        Call BN_copy(k1_abs, dec_k1)
        Call BN_copy(k2_abs, dec_k2)
        k1_abs.neg = False
        k2_abs.neg = False

        If BN_cmp(k1_abs, sqrt_bound) > 0 Or BN_cmp(k2_abs, sqrt_bound) > 0 Then
            Debug.Print "FALHOU: decomposição fora do intervalo √n (regressão " & i & ")"
            GoTo NextRegression
        End If

        base_scalar = random_scalar_mod_n(ctx)
        base_point = ec_point_new()
        reference = ec_point_new()

        If Not ec_point_mul(base_point, base_scalar, ctx.g, ctx) Then
            Debug.Print "FALHOU: geração do ponto base (regressão " & i & ")"
            GoTo NextRegression
        End If

        If Not ec_point_mul(reference, scalar, base_point, ctx) Then
            Debug.Print "FALHOU: multiplicação de referência (regressão " & i & ")"
            GoTo NextRegression
        End If

        Dim beta_point As EC_POINT
        beta_point = apply_endomorphism_for_tests(base_point, ctx)

        Dim k1_point As EC_POINT
        Dim k2_point As EC_POINT
        k1_point = ec_point_new()
        k2_point = ec_point_new()

        If Not ec_point_mul(k1_point, k1_abs, base_point, ctx) Then
            Debug.Print "FALHOU: k1*P falhou (regressão " & i & ")"
            GoTo NextRegression
        End If

        If dec_k1.neg Then
            If Not ec_point_negate(k1_point, k1_point, ctx) Then
                Debug.Print "FALHOU: negação de k1*P (regressão " & i & ")"
                GoTo NextRegression
            End If
        End If

        If Not ec_point_mul(k2_point, k2_abs, beta_point, ctx) Then
            Debug.Print "FALHOU: k2*βP falhou (regressão " & i & ")"
            GoTo NextRegression
        End If

        If dec_k2.neg Then
            If Not ec_point_negate(k2_point, k2_point, ctx) Then
                Debug.Print "FALHOU: negação de k2*βP (regressão " & i & ")"
                GoTo NextRegression
            End If
        End If

        Dim recomposed As EC_POINT
        recomposed = ec_point_new()

        If Not ec_point_add(recomposed, k1_point, k2_point, ctx) Then
            Debug.Print "FALHOU: recomposição de pontos (regressão " & i & ")"
            GoTo NextRegression
        End If

        If ec_point_cmp(reference, recomposed, ctx) = 0 Then
            regression_passed = regression_passed + 1
        Else
            Debug.Print "FALHOU: k1*P + k2*βP ≠ k*P (regressão " & i & ")"
        End If

NextRegression:
    Next i

    Debug.Print "Regressão decomposição GLV: " & regression_passed & " / " & regression_total & " confirmados"
End Sub

Private Function random_scalar_mod_n(ByRef ctx As SECP256K1_CTX) As BIGNUM_TYPE
    Dim hex_str As String
    Dim j As Long

    hex_str = ""
    For j = 1 To 32
        Dim byte_val As Long
        byte_val = Int(Rnd() * 256)
        hex_str = hex_str & Right$("0" & Hex$(byte_val), 2)
    Next j

    Dim scalar As BIGNUM_TYPE
    scalar = BN_hex2bn(hex_str)

    If Not BN_mod(scalar, scalar, ctx.n) Then
        Call BN_zero(scalar)
    End If

    random_scalar_mod_n = scalar
End Function

Private Function apply_endomorphism_for_tests(ByRef point As EC_POINT, ByRef ctx As SECP256K1_CTX) As EC_POINT
    Dim result As EC_POINT
    Dim beta As BIGNUM_TYPE

    result = ec_point_new()
    beta = BN_hex2bn("7AE96A2B657C07106E64479EAC3434E99CF0497512F58995C1396C28719501EE")

    Call BN_mod_mul(result.x, point.x, beta, ctx.p)
    Call BN_copy(result.y, point.y)
    Call BN_set_word(result.z, 1)
    result.infinity = point.infinity

    apply_endomorphism_for_tests = result
End Function
