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
