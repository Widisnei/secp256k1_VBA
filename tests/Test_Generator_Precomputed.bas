Attribute VB_Name = "Test_Generator_Precomputed"
Option Explicit
Option Compare Binary
Option Base 0

' =============================================================================
' TESTES DE REGRESSÃO PARA MULTIPLICAÇÃO DO GERADOR PRÉ-COMPUTADA
' =============================================================================
'
' OBJETIVO:
'   Validar que ec_generator_mul_precomputed_correct produz os mesmos
'   resultados da multiplicação genérica ec_point_mul.
'
' COBERTURA DE TESTES:
'   • Escalares pequenos (1, 2, 3) comparados a G, 2G e 3G
'   • Escalares aleatórios de 256 bits comparados com ec_point_mul
'   • Vetor conhecido utilizado em testes ECDSA
'   • Escalar igual à ordem n para verificar redução modular
'
Public Sub test_generator_precomputed_regression()
    Debug.Print "=== TESTE REGRESSÃO PRECOMPUTED × GENÉRICO ==="

    Call secp256k1_init
    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()

    Dim passed As Long
    Dim total As Long

    Dim scalar As BIGNUM_TYPE
    Dim result_pre As EC_POINT
    Dim result_ref As EC_POINT

    scalar = BN_new()
    result_pre = ec_point_new()
    result_ref = ec_point_new()

    ' ---------------------------
    ' Testes com escalares 1, 2, 3
    ' ---------------------------
    Call BN_set_word(scalar, 1)
    Call ec_generator_mul_precomputed_correct(result_pre, scalar, ctx)
    total = total + 1
    If ec_point_cmp(result_pre, ctx.g, ctx) = 0 Then
        passed = passed + 1
        Debug.Print "✓ k=1 corresponde a G"
    Else
        Debug.Print "✗ k=1 não corresponde a G"
    End If

    Dim double_g As EC_POINT
    double_g = ec_point_new()
    Call ec_point_double(double_g, ctx.g, ctx)

    Call BN_set_word(scalar, 2)
    Call ec_generator_mul_precomputed_correct(result_pre, scalar, ctx)
    total = total + 1
    If ec_point_cmp(result_pre, double_g, ctx) = 0 Then
        passed = passed + 1
        Debug.Print "✓ k=2 corresponde a 2G"
    Else
        Debug.Print "✗ k=2 não corresponde a 2G"
    End If

    Dim triple_g As EC_POINT
    triple_g = ec_point_new()
    Call ec_point_double(triple_g, ctx.g, ctx)
    Call ec_point_add(triple_g, triple_g, ctx.g, ctx)

    Call BN_set_word(scalar, 3)
    Call ec_generator_mul_precomputed_correct(result_pre, scalar, ctx)
    total = total + 1
    If ec_point_cmp(result_pre, triple_g, ctx) = 0 Then
        passed = passed + 1
        Debug.Print "✓ k=3 corresponde a 3G"
    Else
        Debug.Print "✗ k=3 não corresponde a 3G"
    End If

    ' ---------------------------
    ' Escalar zero: deve coincidir com caminho genérico
    ' ---------------------------
    Call BN_set_word(scalar, 0)
    Call ec_generator_mul_precomputed_correct(result_pre, scalar, ctx)
    Call ec_point_mul(result_ref, scalar, ctx.g, ctx)

    total = total + 1
    If result_pre.infinity And ec_point_cmp(result_pre, result_ref, ctx) = 0 Then
        passed = passed + 1
        Debug.Print "✓ k=0 corresponde ao ponto no infinito"
    Else
        Debug.Print "✗ k=0 não corresponde ao ponto no infinito"
    End If

    ' ---------------------------
    ' Escalares aleatórios e vetor conhecido
    ' ---------------------------
    Dim random_scalars() As String
    random_scalars = Array( _
        "1F4C3B2A19080706050403020100FFEEDDCCBBAA99887766554433221100FF", _
        "A1B2C3D4E5F6071827364554637281900A1B2C3D4E5F6A7B8C9D0E1F2233445", _
        "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721", _
        "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEBAAEDCE6AF48A03BBFD25E8CD0364130", _
        "0123456789ABCDEFFEDCBA98765432100123456789ABCDEFFEDCBA9876543210", _
        "FFFFFFFF00000000FFFFFFFFFFFFFFFFBCE6FAADA7179E84F3B9CAC2FC632551", _
        "7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", _
        "B7E08E29CBB3530EBEE24F514941CA9FE0C5E1E5E3E0FA1F1B4D5C7A1E53EAB4")

    Dim idx As Long
    For idx = LBound(random_scalars) To UBound(random_scalars)
        Dim scalar_hex As BIGNUM_TYPE
        scalar_hex = BN_hex2bn(random_scalars(idx))

        Call ec_generator_mul_precomputed_correct(result_pre, scalar_hex, ctx)
        Call ec_point_mul(result_ref, scalar_hex, ctx.g, ctx)

        total = total + 1
        If ec_point_cmp(result_pre, result_ref, ctx) = 0 Then
            passed = passed + 1
            Debug.Print "✓ Escalar " & random_scalars(idx)
        Else
            Debug.Print "✗ Divergência para escalar " & random_scalars(idx)
        End If
    Next idx

    ' ---------------------------
    ' Escalar igual à ordem: deve resultar no ponto no infinito
    ' ---------------------------
    Call ec_point_set_infinity(result_pre)
    Call ec_point_mul(result_ref, ctx.n, ctx.g, ctx)
    Call ec_generator_mul_precomputed_correct(result_pre, ctx.n, ctx)
    total = total + 1
    If result_pre.infinity And ec_point_cmp(result_pre, result_ref, ctx) = 0 Then
        passed = passed + 1
        Debug.Print "✓ k = n retorna ponto no infinito"
    Else
        Debug.Print "✗ k = n não retornou infinito"
    End If

    Debug.Print "--- RESUMO: " & passed & " / " & total & " testes aprovados ---"
End Sub
