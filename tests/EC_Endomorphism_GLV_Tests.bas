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

    Dim lib_vectors() As String
    lib_vectors = Split( _
        "CAE2456B168C3AFA8AB6156551D591BAF8CFA4604FD73128A4D97803D3EC1A43|FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE4C1DCE0F5BE6E65205CDF1383143BA68,FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE4C57DED1A5209BD28218D5FD375E6C4A;" & _
        "AC7E89A4DADBCB22AEBBCC8B0A222A4894AF4D43D41D9D865AFE64347AF684B7|FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEABA85EB6932FA760AA2CD2E1D05C43C4,000000000000000000000000000000003E634FFDEC7E72FE849B08AAB4AD5A2A;" & _
        "9C6D5BF2DA9A15C80889EE9F6A68FBC6C655F2337D0978B69EF3E66079E944EC|FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE8861F5737F77C977B0EB8ADEA5204D2C,FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE995DFB26B7B94546753269D18AF0B8CC;" & _
        "B97B47FB6FF8FAF8C93D54089F5C1929938A83D6BFAE1F4AECB17E57A27C6251|FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEAB74C50EEE5A346DF30E1513636D2319,FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE7DD71245D8C3224C66C9A4459A7021A5;" & _
        "CD4B9E3EF628955AE0B9D88521F73F45435C82ECFA1312F45C997040B61E2077|FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE8AA87D16D63CEAB3E3542304534AA742,FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE4210155AD69104CDD939487914AB3E89;" & _
        "3A36CABDC43FF4DEC5FA629D275F3CD48A02D9B082C8D5F4A36A2AC4D7F2B748|000000000000000000000000000000004D485E9F6BF02BEBE9BD891E4A4A7CCC,FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEAE0220F703C20C0F8E12DD8DBC418DFA;" & _
        "0E679417792BD41582B9135CB02EE682A62921A6D3A6AFB6C659C7E541EF048C|FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE56CDF5C0B1BDAC679492730002D1B553,FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE41E69206B1175EF9B578FCFCB303C472;" & _
        "EE6433CCDA37DFB0DDAE0B2B1BD0BCAFB4BBCB1F8B26A1339D7F6291D2A6687B|000000000000000000000000000000002088B6F120DF404CB020194D3C612AAC,000000000000000000000000000000001EDA8C482A2896C9AB7ED069B6899A39;" & _
        "C61014F1EFB048BA67CAB62015148581141677538684CE8BF9412583FA27188D|000000000000000000000000000000003B26B3D7753A579D43826BC983E24263,FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE5B1516EB168A27CB26476FB0355EDD4C;" & _
        "DB8431A983564EA4381B22B79B155B6603418B3444F0B9A1D29B7498F80D6630|FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE7E3FCFB3D5768EFD72624F21DD062636,000000000000000000000000000000004A74F97E591028DB1AA668EE8334C4BB;" & _
        "338930D07E915EE493949F550BD2585233A11A5F52D13D62CDA7DB16000E8D67|0000000000000000000000000000000005FE5020E012E4D4B27A6F7AA5B269FE,FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE5B58F1A2F42B5F14CC6AF447C94A18CF;" & _
        "99BDA25E46BFD3D4B4CA72BC3C82E09051456136CAE71EAA922E20F751366158|000000000000000000000000000000004E5703DDB10F2C261B9AB7C339E10EE3,FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE795E0C098DD0AA653CAB191B305B223C;" & _
        "9DEFE89BE2DCAE6BC922262079C29CBFB3BE872ADFB15C58F2A4D21E3395030C|FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE561EEB025290024AA2507F7F43C8E97F,FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE9D4517EF3DFAE081E40BBB4B0E27999E;" & _
        "1338A37C736832CDC99CF9213CA9DFF2C9C1DFE891FD8490CA2D8F46321C1C46|FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE52E23B3118C980BD6C4A16BAB5C7B29F,FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE84B291CF4FCB763D49D5AFC3DF4B0312;" & _
        "2741E29153E507946EB18A3EB38056BA51F89F25545E4E2EA26191EA9AA94DFF|FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE3D29194645EFB42AA6D5465CF9FEABAF,0000000000000000000000000000000033A69856B3F752DF7C9046CEAFE719AE;" & _
        "32C26EC9ED1F29AEF71F51FFC692FF3FFF3948160688B7324DD1332A8B60EAA0|0000000000000000000000000000000018667E0F3F49C3DDB0C9C0CF9783D408,FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFE668CA3C9E49105D79480FC1B43CF8366", ";")

    Dim half_n As BIGNUM_TYPE
    half_n = BN_new()
    Call BN_copy(half_n, ctx.n)
    Call BN_rshift(half_n, half_n, 1)

    Dim sqrt_bound As BIGNUM_TYPE
    sqrt_bound = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF")

    Dim regression_passed As Long
    Dim regression_total As Long
    regression_total = UBound(lib_vectors)

    For i = 0 To regression_total
        Dim parts() As String
        parts = Split(lib_vectors(i), "|")

        If UBound(parts) <> 1 Then GoTo NextRegression

        Dim expected_parts() As String
        expected_parts = Split(parts(1), ",")
        If UBound(expected_parts) <> 1 Then GoTo NextRegression

        scalar = BN_hex2bn(parts(0))

        Dim expected_k1 As BIGNUM_TYPE
        Dim expected_k2 As BIGNUM_TYPE
        expected_k1 = BN_hex2bn(expected_parts(0))
        expected_k2 = BN_hex2bn(expected_parts(1))

        Call reduce_signed(expected_k1, ctx.n, half_n)
        Call reduce_signed(expected_k2, ctx.n, half_n)

        Dim dec_k1 As BIGNUM_TYPE
        Dim dec_k2 As BIGNUM_TYPE
        dec_k1 = BN_new()
        dec_k2 = BN_new()

        If Not glv_decompose_scalar_for_tests(dec_k1, dec_k2, scalar, ctx) Then
            Debug.Print "FALHOU: decomposição GLV falhou (regressão " & i + 1 & ")"
            GoTo NextRegression
        End If

        If BN_cmp(dec_k1, expected_k1) <> 0 Or BN_cmp(dec_k2, expected_k2) <> 0 Then
            Debug.Print "FALHOU: divergência vs libsecp256k1 (regressão " & i + 1 & ")"
            GoTo NextRegression
        End If

        Dim k1_abs As BIGNUM_TYPE
        Dim k2_abs As BIGNUM_TYPE
        k1_abs = BN_new()
        k2_abs = BN_new()
        Call BN_copy(k1_abs, dec_k1)
        Call BN_copy(k2_abs, dec_k2)
        k1_abs.neg = False
        k2_abs.neg = False

        If BN_cmp(k1_abs, sqrt_bound) > 0 Or BN_cmp(k2_abs, sqrt_bound) > 0 Then
            Debug.Print "FALHOU: decomposição fora do intervalo √n (regressão " & i + 1 & ")"
            GoTo NextRegression
        End If

        base_scalar = random_scalar_mod_n(ctx)
        base_point = ec_point_new()
        reference = ec_point_new()

        If Not ec_point_mul(base_point, base_scalar, ctx.g, ctx) Then
            Debug.Print "FALHOU: geração do ponto base (regressão " & i + 1 & ")"
            GoTo NextRegression
        End If

        If Not ec_point_mul(reference, scalar, base_point, ctx) Then
            Debug.Print "FALHOU: multiplicação de referência (regressão " & i + 1 & ")"
            GoTo NextRegression
        End If

        Dim beta_point As EC_POINT
        beta_point = apply_endomorphism_for_tests(base_point, ctx)

        Dim k1_point As EC_POINT
        Dim k2_point As EC_POINT
        k1_point = ec_point_new()
        k2_point = ec_point_new()

        If Not ec_point_mul(k1_point, k1_abs, base_point, ctx) Then
            Debug.Print "FALHOU: k1*P falhou (regressão " & i + 1 & ")"
            GoTo NextRegression
        End If

        If dec_k1.neg Then
            If Not ec_point_negate(k1_point, k1_point, ctx) Then
                Debug.Print "FALHOU: negação de k1*P (regressão " & i + 1 & ")"
                GoTo NextRegression
            End If
        End If

        If Not ec_point_mul(k2_point, k2_abs, beta_point, ctx) Then
            Debug.Print "FALHOU: k2*βP falhou (regressão " & i + 1 & ")"
            GoTo NextRegression
        End If

        If dec_k2.neg Then
            If Not ec_point_negate(k2_point, k2_point, ctx) Then
                Debug.Print "FALHOU: negação de k2*βP (regressão " & i + 1 & ")"
                GoTo NextRegression
            End If
        End If

        Dim recomposed As EC_POINT
        recomposed = ec_point_new()

        If Not ec_point_add(recomposed, k1_point, k2_point, ctx) Then
            Debug.Print "FALHOU: recomposição de pontos (regressão " & i + 1 & ")"
            GoTo NextRegression
        End If

        If ec_point_cmp(reference, recomposed, ctx) = 0 Then
            regression_passed = regression_passed + 1
        Else
            Debug.Print "FALHOU: k1*P + k2*βP ≠ k*P (regressão " & i + 1 & ")"
        End If

NextRegression:
    Next i

    Debug.Print "Regressão decomposição GLV: " & regression_passed & " / " & (regression_total + 1) & " confirmados"
End Sub

Private Sub reduce_signed(ByRef value As BIGNUM_TYPE, ByRef modulus As BIGNUM_TYPE, ByRef half_mod As BIGNUM_TYPE)
    If BN_cmp(value, half_mod) > 0 Then
        If Not BN_sub(value, value, modulus) Then
            Call BN_zero(value)
        End If
    End If
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
