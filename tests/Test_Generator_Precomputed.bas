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
    Dim result_core As EC_POINT
    Dim result_ref As EC_POINT

    scalar = BN_new()
    result_pre = ec_point_new()
    result_core = ec_point_new()
    result_ref = ec_point_new()

    ' ---------------------------
    ' Testes com escalares 1, 2, 3
    ' ---------------------------
    Call BN_set_word(scalar, 1)
    Call ec_point_mul(result_ref, scalar, ctx.g, ctx)
    Call ec_generator_mul_precomputed_correct(result_pre, scalar, ctx)
    Call ec_generator_mul_bitcoin_core(result_core, scalar, ctx)
    total = total + 1
    If ec_point_cmp(result_pre, result_ref, ctx) = 0 Then
        passed = passed + 1
        Debug.Print "✓ precomputed_correct k=1 corresponde a G"
    Else
        Debug.Print "✗ precomputed_correct k=1 difere de G"
    End If
    total = total + 1
    If ec_point_cmp(result_core, result_ref, ctx) = 0 Then
        passed = passed + 1
        Debug.Print "✓ bitcoin_core k=1 corresponde a G"
    Else
        Debug.Print "✗ bitcoin_core k=1 difere de G"
    End If

    Call BN_set_word(scalar, 2)
    Call ec_point_mul(result_ref, scalar, ctx.g, ctx)
    Call ec_generator_mul_precomputed_correct(result_pre, scalar, ctx)
    Call ec_generator_mul_bitcoin_core(result_core, scalar, ctx)
    total = total + 1
    If ec_point_cmp(result_pre, result_ref, ctx) = 0 Then
        passed = passed + 1
        Debug.Print "✓ precomputed_correct k=2 corresponde a 2G"
    Else
        Debug.Print "✗ precomputed_correct k=2 difere de 2G"
    End If
    total = total + 1
    If ec_point_cmp(result_core, result_ref, ctx) = 0 Then
        passed = passed + 1
        Debug.Print "✓ bitcoin_core k=2 corresponde a 2G"
    Else
        Debug.Print "✗ bitcoin_core k=2 difere de 2G"
    End If

    Call BN_set_word(scalar, 3)
    Call ec_point_mul(result_ref, scalar, ctx.g, ctx)
    Call ec_generator_mul_precomputed_correct(result_pre, scalar, ctx)
    Call ec_generator_mul_bitcoin_core(result_core, scalar, ctx)
    total = total + 1
    If ec_point_cmp(result_pre, result_ref, ctx) = 0 Then
        passed = passed + 1
        Debug.Print "✓ precomputed_correct k=3 corresponde a 3G"
    Else
        Debug.Print "✗ precomputed_correct k=3 difere de 3G"
    End If
    total = total + 1
    If ec_point_cmp(result_core, result_ref, ctx) = 0 Then
        passed = passed + 1
        Debug.Print "✓ bitcoin_core k=3 corresponde a 3G"
    Else
        Debug.Print "✗ bitcoin_core k=3 difere de 3G"
    End If

    ' ---------------------------
    ' Escalar zero: deve coincidir com caminho genérico
    ' ---------------------------
    Call BN_set_word(scalar, 0)
    Call ec_point_mul(result_ref, scalar, ctx.g, ctx)
    Call ec_generator_mul_precomputed_correct(result_pre, scalar, ctx)
    Call ec_generator_mul_bitcoin_core(result_core, scalar, ctx)

    total = total + 1
    If result_pre.infinity And ec_point_cmp(result_pre, result_ref, ctx) = 0 Then
        passed = passed + 1
        Debug.Print "✓ precomputed_correct k=0 corresponde ao ponto no infinito"
    Else
        Debug.Print "✗ precomputed_correct k=0 não corresponde ao ponto no infinito"
    End If
    total = total + 1
    If result_core.infinity And ec_point_cmp(result_core, result_ref, ctx) = 0 Then
        passed = passed + 1
        Debug.Print "✓ bitcoin_core k=0 corresponde ao ponto no infinito"
    Else
        Debug.Print "✗ bitcoin_core k=0 não corresponde ao ponto no infinito"
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

        Call ec_point_mul(result_ref, scalar_hex, ctx.g, ctx)
        Call ec_generator_mul_precomputed_correct(result_pre, scalar_hex, ctx)
        Call ec_generator_mul_bitcoin_core(result_core, scalar_hex, ctx)

        total = total + 1
        If ec_point_cmp(result_pre, result_ref, ctx) = 0 Then
            passed = passed + 1
            Debug.Print "✓ precomputed_correct escalar " & random_scalars(idx)
        Else
            Debug.Print "✗ precomputed_correct divergência para escalar " & random_scalars(idx)
        End If

        total = total + 1
        If ec_point_cmp(result_core, result_ref, ctx) = 0 Then
            passed = passed + 1
            Debug.Print "✓ bitcoin_core escalar " & random_scalars(idx)
        Else
            Debug.Print "✗ bitcoin_core divergência para escalar " & random_scalars(idx)
        End If
    Next idx

    ' ---------------------------
    ' Escalar igual à ordem: deve resultar no ponto no infinito
    ' ---------------------------
    Call ec_point_set_infinity(result_pre)
    Call ec_point_mul(result_ref, ctx.n, ctx.g, ctx)
    Call ec_generator_mul_precomputed_correct(result_pre, ctx.n, ctx)
    Call ec_generator_mul_bitcoin_core(result_core, ctx.n, ctx)
    total = total + 1
    If result_pre.infinity And ec_point_cmp(result_pre, result_ref, ctx) = 0 Then
        passed = passed + 1
        Debug.Print "✓ precomputed_correct k = n retorna ponto no infinito"
    Else
        Debug.Print "✗ precomputed_correct k = n não retornou infinito"
    End If
    total = total + 1
    If result_core.infinity And ec_point_cmp(result_core, result_ref, ctx) = 0 Then
        passed = passed + 1
        Debug.Print "✓ bitcoin_core k = n retorna ponto no infinito"
    Else
        Debug.Print "✗ bitcoin_core k = n não retornou infinito"
    End If

    Debug.Print "--- RESUMO: " & passed & " / " & total & " testes aprovados ---"
End Sub

Public Sub test_bitcoin_core_generator_entry_conversion()
    ' -------------------------------------------------------------------------
    ' PROPÓSITO:
    '   Validar a correção da conversão little-endian → big-endian aplicada
    '   às entradas da tabela do Bitcoin Core, verificando especificamente se
    '   as coordenadas conhecidas do gerador resultam na forma SEC comprimida
    '   esperada após a normalização.
    ' -------------------------------------------------------------------------

    Debug.Print "=== TESTE CONVERSÃO ENDIANNESS PONTO GERADOR ==="

    Call secp256k1_init
    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()

    ' Coordenadas do gerador G em little-endian por palavras de 32 bits
    Dim generator_entry As String
    generator_entry = _
        "16F81798,59F2815B,2DCE28D9,029BFCDB,CE870B07,55A06295,F9DCBBAC,79BE667E," & _
        "FB10D4B8,9C47D08F,A6855419,FD17B448,0E1108A8,5DA4FBFC,26A3C465,483ADA77"

    Dim point As EC_POINT
    point = ec_point_new()

    If EC_Precomputed_Integration.convert_bitcoin_core_point(generator_entry, point, ctx) Then
        Dim on_curve As Boolean
        on_curve = ec_point_is_on_curve(point, ctx)

        Dim x_hex As String, y_hex As String
        x_hex = BN_bn2hex(point.x)
        y_hex = BN_bn2hex(point.y)

        Dim prefix As String
        If BN_is_odd(point.y) Then
            prefix = "03"
        Else
            prefix = "02"
        End If

        Dim compressed As String
        compressed = prefix & x_hex

        Const EXPECTED_COMPRESSED As String = _
            "0279BE667EF9DCBBAC55A06295CE870B07029BFCDB2DCE28D959F2815B16F81798"

        If on_curve And compressed = EXPECTED_COMPRESSED Then
            Debug.Print "✓ Conversão reproduz SEC comprimido do gerador"
        Else
            Debug.Print "✗ Conversão do gerador falhou"
            Debug.Print "  on_curve = " & on_curve
            Debug.Print "  X = " & x_hex
            Debug.Print "  Y = " & y_hex
            Debug.Print "  Compressed = " & compressed
        End If
    Else
        Debug.Print "✗ Falha ao converter coordenadas little-endian do gerador"
    End If
End Sub
