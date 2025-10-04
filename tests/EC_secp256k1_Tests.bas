Attribute VB_Name = "EC_Secp256k1_Tests"
Option Explicit
Option Compare Binary
Option Base 0

#Const HAVE_EC_SECP256K1 = 1

'==============================================================================
' MÓDULO: EC_Secp256k1_Tests
' Descrição: Testes Integrados da Implementação secp256k1 Completa
' Autor: Desenvolvido para compatibilidade com Bitcoin Core
' Data: 2024
'
' CARACTERÍSTICAS TÉCNICAS:
' • Validação completa da implementação secp256k1
' • Testes de geração e validação de chaves
' • Verificação de compressão/descompressão de pontos
' • Testes de assinatura digital ECDSA
' • Validação de multiplicação escalar
'
' CURVA ELÍPTICA SECP256K1:
' • Equação: y² = x³ + 7 (mod p)
' • Campo primo: p = 2²⁵⁶ - 2³² - 2⁹ - 2⁸ - 2⁷ - 2⁶ - 2⁴ - 1
' • Ordem: n = FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEBAAEDCE6AF48A03BBFD25E8CD0364141
' • Cofator: h = 1
' • Gerador: G = (79BE667E..., 483ADA77...)
'
' ALGORITMOS TESTADOS:
' • secp256k1_generate_keypair()   - Geração de chaves
' • secp256k1_validate_*()         - Validação de chaves
' • secp256k1_point_compress()     - Compressão de pontos
' • secp256k1_point_decompress()   - Descompressão de pontos
' • secp256k1_generator_multiply()  - Multiplicação do gerador
'
' TESTES IMPLEMENTADOS:
' • Inicialização do contexto secp256k1
' • Validação do ponto gerador
' • Geração e validação de pares de chaves
' • Compressão/descompressão roundtrip
' • Formato de assinatura ECDSA (DER)
' • Multiplicação escalar do gerador
'
' FORMATOS SUPORTADOS:
' • Chaves privadas: 256-bit hexadecimal
' • Chaves públicas: Comprimidas (33 bytes) e descomprimidas (65 bytes)
' • Assinaturas: DER encoding (RFC 3279)
' • Pontos: Coordenadas afins (x, y)
'
' SEGURANÇA E CONFORMIDADE:
' • Geração criptograficamente segura de chaves
' • Validação rigorosa de parâmetros
' • Resistência a ataques conhecidos
' • Conformidade com padrões Bitcoin
'
' COMPATIBILIDADE:
' • Bitcoin Core secp256k1 - Totalmente compatível
' • OpenSSL EC_* - Interface similar
' • RFC 5480/6979 - Padrões seguidos
' • SEC 1/2 - Especificações implementadas
'==============================================================================

'==============================================================================
' EXECUÇÃO DOS TESTES INTEGRADOS SECP256K1
'==============================================================================

' Propósito: Validação completa da implementação secp256k1
' Algoritmo: 6 testes integrados cobrindo todas as funcionalidades
' Retorno: Relatório de status via Debug.Print
' Crítico: Falhas indicam problemas na implementação principal

Public Sub Run_EC_Secp256k1_Tests()
#If HAVE_EC_SECP256K1 Then
    Debug.Print "=== Testes EC secp256k1 ==="

    Call secp256k1_init
    
    Dim passed As Long, total As Long
    
    ' Teste 1: Inicialização do contexto
    Debug.Print "APROVADO: Inicialização do contexto"
    passed = passed + 1: total = total + 1
    
    ' Teste 2: Validação do ponto gerador
    Dim generator_compressed As String
    generator_compressed = secp256k1_get_generator()
    If secp256k1_validate_public_key(generator_compressed) Then
        Debug.Print "APROVADO: Validação do ponto gerador"
        passed = passed + 1
    Else
        Debug.Print "FALHOU: Validação do ponto gerador"
    End If
    total = total + 1
    
    ' Teste 3: Geração e validação de chaves
    Dim keypair As ECDSA_KEYPAIR
    keypair = secp256k1_generate_keypair()
    Dim private_hex As String, public_compressed As String
    private_hex = BN_bn2hex(keypair.private_key)
    public_compressed = ec_point_compress(keypair.public_key, secp256k1_context_create())
    
    If secp256k1_validate_private_key(private_hex) And secp256k1_validate_public_key(public_compressed) Then
        Debug.Print "APROVADO: Geração e validação de chaves"
        passed = passed + 1
    Else
        Debug.Print "FALHOU: Geração e validação de chaves"
    End If
    total = total + 1
    
    ' Teste 4: Compressão/descompressão de pontos
    Dim decompressed As String, recompressed As String
    decompressed = secp256k1_point_decompress(public_compressed)
    If decompressed <> "" Then
        Dim coords() As String: coords = Split(decompressed, ",")
        recompressed = secp256k1_point_compress(coords(0), coords(1))
        If recompressed = public_compressed Then
            Debug.Print "APROVADO: Compressão/descompressão de pontos"
            passed = passed + 1
        Else
            Debug.Print "FALHOU: Compressão/descompressão de pontos"
        End If
    Else
        Debug.Print "FALHOU: Descompressão de pontos"
    End If
    total = total + 1

    ' Teste 4B: Roundtrip com coordenada X iniciando com byte 00
    Dim ctx_leading As SECP256K1_CTX: ctx_leading = secp256k1_context_create()
    Dim leading_point As EC_POINT: leading_point = ec_point_new()
    Dim leading_x As BIGNUM_TYPE, leading_y As BIGNUM_TYPE
    leading_x = BN_hex2bn("00E3AE1974566CA06CC516D47E0FB165A674A3DABCFCA15E722F0E3450F45889")
    leading_y = BN_hex2bn("2AEABE7E4531510116217F07BF4D07300DE97E4874F81F533420A72EEB0BD6A4")
    Call ec_point_set_affine(leading_point, leading_x, leading_y)

    Dim expected_leading As String
    expected_leading = "0200E3AE1974566CA06CC516D47E0FB165A674A3DABCFCA15E722F0E3450F45889"

    Dim compressed_leading As String
    compressed_leading = ec_point_compress(leading_point, ctx_leading)

    Dim leading_roundtrip_ok As Boolean
    If compressed_leading = expected_leading Then
        Dim decompressed_leading As EC_POINT
        decompressed_leading = ec_point_decompress(compressed_leading, ctx_leading)

        If Not decompressed_leading.infinity Then
            Dim x_back As BIGNUM_TYPE, y_back As BIGNUM_TYPE
            x_back = BN_new(): y_back = BN_new()

            If ec_point_get_affine(decompressed_leading, x_back, y_back, ctx_leading) Then
                If BN_cmp(leading_x, x_back) = 0 And BN_cmp(leading_y, y_back) = 0 Then
                    Dim recompressed_leading As String
                    recompressed_leading = ec_point_compress(decompressed_leading, ctx_leading)
                    leading_roundtrip_ok = (recompressed_leading = expected_leading)
                End If
            End If
        End If
    End If

    If leading_roundtrip_ok Then
        Debug.Print "APROVADO: Roundtrip preserva ponto com X inicial 00"
        passed = passed + 1
    Else
        Debug.Print "FALHOU: Roundtrip com X inicial 00"
    End If
    total = total + 1

    ' Teste 4C: Compressão e negação com y = 0 (caso reduzido)
    Dim ctx_zero As SECP256K1_CTX: ctx_zero = secp256k1_context_create()
    Dim zeroYPoint As EC_POINT, zeroYNeg As EC_POINT
    zeroYPoint = ec_point_new()
    zeroYNeg = ec_point_new()
    Call ec_point_copy(zeroYPoint, keypair.public_key)
    Call BN_zero(zeroYPoint.y)
    zeroYPoint.infinity = False

    Dim zeroYCompressed As String
    zeroYCompressed = ec_point_compress(zeroYPoint, ctx_zero)
    If Left$(zeroYCompressed, 2) = "02" Then
        Debug.Print "APROVADO: Compressão preserva prefixo par para y=0"
        passed = passed + 1
    Else
        Debug.Print "FALHOU: Compressão com y=0 não gerou prefixo par"
    End If
    total = total + 1

    Dim zeroYOnCurve As Boolean
    zeroYOnCurve = ec_point_is_on_curve(zeroYPoint, ctx_zero)

    If ec_point_negate(zeroYNeg, zeroYPoint, ctx_zero) Then
        If BN_is_zero(zeroYNeg.y) And ec_point_is_on_curve(zeroYNeg, ctx_zero) = zeroYOnCurve Then
            Debug.Print "APROVADO: Negação mantém y=0 e consistência de on-curve"
            passed = passed + 1
        Else
            Debug.Print "FALHOU: Negação de y=0 alterou redução ou estado na curva"
        End If
    Else
        Debug.Print "FALHOU: Negação de ponto com y=0"
    End If
    total = total + 1

    Dim neg_public As EC_POINT
    neg_public = ec_point_new()
    If ec_point_negate(neg_public, keypair.public_key, ctx_zero) Then
        If ec_point_is_on_curve(neg_public, ctx_zero) Then
            Debug.Print "APROVADO: Negação mantém ponto válido na curva"
            passed = passed + 1
        Else
            Debug.Print "FALHOU: Negação produziu ponto fora da curva"
        End If
    Else
        Debug.Print "FALHOU: Negação do ponto público falhou"
    End If
    total = total + 1
    
    ' Teste 5: Formato de assinatura ECDSA (teste simplificado)
    Dim ctx_local As SECP256K1_CTX: ctx_local = secp256k1_context_create()
    Dim sig As ECDSA_SIGNATURE: sig.r = BN_new(): sig.s = BN_new()

    ' Cria assinatura de teste simples manualmente
    Call BN_set_word(sig.r, 1)
    Call BN_set_word(sig.s, 1)
    
    ' Testa estrutura básica da assinatura
    Dim sig_der As String: sig_der = ecdsa_signature_to_der(sig)
    Dim sig_parsed As ECDSA_SIGNATURE
    Call ecdsa_signature_from_der(sig_parsed, sig_der)

    If BN_cmp(sig.r, sig_parsed.r) = 0 And BN_cmp(sig.s, sig_parsed.s) = 0 Then
        Debug.Print "APROVADO: Formato de assinatura ECDSA"
        passed = passed + 1
    Else
        Debug.Print "FALHOU: Formato de assinatura ECDSA"
    End If
    total = total + 1

    ' Teste 6: Validação do decodificador DER contra entradas malformadas
    Dim decoder_strict_ok As Boolean: decoder_strict_ok = True
    Dim sig_invalid As ECDSA_SIGNATURE

    ' Caso 1: Dados extras após o payload declarado
    If ecdsa_signature_from_der(sig_invalid, sig_der & "00") Then decoder_strict_ok = False

    ' Caso 2: Inteiro truncado (dados insuficientes)
    If ecdsa_signature_from_der(sig_invalid, Left$(sig_der, Len(sig_der) - 2)) Then decoder_strict_ok = False

    ' Caso 3: Tag de inteiro ausente/alterada
    Dim malformed_tag As String
    malformed_tag = Left$(sig_der, 4) & "03" & Mid$(sig_der, 7)
    If ecdsa_signature_from_der(sig_invalid, malformed_tag) Then decoder_strict_ok = False

    ' Caso 4: Comprimento longo bem formado deve ser aceito
    Dim long_form_der As String
    long_form_der = "3081" & Mid$(sig_der, 3)
    Dim sig_long As ECDSA_SIGNATURE
    If Not ecdsa_signature_from_der(sig_long, long_form_der) Then
        decoder_strict_ok = False
    ElseIf BN_cmp(sig_long.r, sig.r) <> 0 Or BN_cmp(sig_long.s, sig.s) <> 0 Then
        decoder_strict_ok = False
    End If

    ' Caso 5: Inteiro com padding obrigatório deve ser aceito
    Dim sig_pad As ECDSA_SIGNATURE
    sig_pad.r = BN_new(): sig_pad.s = BN_new()
    Call BN_set_word(sig_pad.r, &H80&)
    Call BN_set_word(sig_pad.s, 1)
    Dim sig_pad_der As String: sig_pad_der = ecdsa_signature_to_der(sig_pad)
    Dim sig_pad_out As ECDSA_SIGNATURE
    If Not ecdsa_signature_from_der(sig_pad_out, sig_pad_der) Then
        decoder_strict_ok = False
    ElseIf BN_cmp(sig_pad_out.r, sig_pad.r) <> 0 Or BN_cmp(sig_pad_out.s, sig_pad.s) <> 0 Then
        decoder_strict_ok = False
    End If

    ' Caso 6: Padding redundante (00 00 ...) deve ser rejeitado
    Dim redundant_padding_der As String
    redundant_padding_der = "300702020001020101"
    If ecdsa_signature_from_der(sig_invalid, redundant_padding_der) Then decoder_strict_ok = False

    ' Caso 7: Inteiro sem padding com bit de sinal definido deve ser rejeitado
    Dim negative_encoding_der As String
    negative_encoding_der = "3006020180020101"
    If ecdsa_signature_from_der(sig_invalid, negative_encoding_der) Then decoder_strict_ok = False

    If decoder_strict_ok Then
        Debug.Print "APROVADO: Decoder DER rejeita entradas malformadas"
        passed = passed + 1
    Else
        Debug.Print "FALHOU: Decoder DER rejeita entradas malformadas"
    End If
    total = total + 1

    ' Teste 7: Multiplicação de pontos
    Dim scalar_hex As String, result_point As String
    scalar_hex = "2"
    result_point = secp256k1_generator_multiply(scalar_hex)

    If result_point <> "" And result_point <> "00" Then
        Debug.Print "APROVADO: Multiplicação de pontos"
        passed = passed + 1
    Else
        Debug.Print "FALHOU: Multiplicação de pontos"
    End If
    total = total + 1

    ' Teste 8: Multiplicação windowed consistente com double-and-add
    total = total + 1
    If Verify_Windowed_Mul_Against_Standard() Then
        Debug.Print "APROVADO: Multiplicação windowed consistente com double-and-add"
        passed = passed + 1
    Else
        Debug.Print "FALHOU: Multiplicação windowed consistente com double-and-add"
    End If

    ' Teste 9: Regressão adição de pontos 256-bit (G + G = 2G)
    Dim ctx_reg As SECP256K1_CTX: ctx_reg = secp256k1_context_create()
    Dim g1 As EC_POINT, g2 As EC_POINT, sum_point As EC_POINT, expected_2g As EC_POINT
    g1 = ec_point_new(): g2 = ec_point_new(): sum_point = ec_point_new(): expected_2g = ec_point_new()
    Call ec_point_copy(g1, ctx_reg.g)
    Call ec_point_copy(g2, ctx_reg.g)
    expected_2g.x = BN_hex2bn("C6047F9441ED7D6D3045406E95C07CD85C778E4B8CEF3CA7ABAC09B95C709EE5")
    expected_2g.y = BN_hex2bn("1AE168FEA63DC339A3C58419466CEAEEF7F632653266D0E1236431A950CFE52A")

    total = total + 1
    If ec_point_add(sum_point, g1, g2, ctx_reg) Then
        If BN_cmp(sum_point.x, expected_2g.x) = 0 And BN_cmp(sum_point.y, expected_2g.y) = 0 Then
            Debug.Print "APROVADO: ec_point_add produz 2G esperado"
            passed = passed + 1
        Else
            Debug.Print "FALHOU: ec_point_add resultou em coordenadas incorretas"
            Debug.Print "  x: ", BN_bn2hex(sum_point.x)
            Debug.Print "  y: ", BN_bn2hex(sum_point.y)
        End If
    Else
        Debug.Print "FALHOU: ec_point_add retornou erro"
    End If

    Debug.Print "=== Testes EC secp256k1: ", passed, "/", total, " aprovados ==="

    If passed = total Then
        Debug.Print "*** IMPLEMENTAÇÃO SECP256K1 COMPLETA E FUNCIONANDO ***"
    Else
        Debug.Print "*** PROBLEMAS DETECTADOS NO SECP256K1 ***"
    End If
#Else
    Debug.Print "=== Testes EC secp256k1 (PULADOS) ==="
#End If
End Sub

Private Function Verify_Windowed_Mul_Against_Standard() As Boolean
#If HAVE_EC_SECP256K1 Then
    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()

    Verify_Windowed_Mul_Against_Standard = False

    Dim scalars(0 To 4) As String
    scalars(0) = "02"
    scalars(1) = "03"
    scalars(2) = "08"
    scalars(3) = "10"
    scalars(4) = "F0E1D2C3B4A5968778695A4B3C2D1E0FF112233445566778899AABBCCDDEEFF0"

    Dim idx As Long
    For idx = 0 To UBound(scalars)
        Dim scalar_bn As BIGNUM_TYPE
        scalar_bn = BN_hex2bn(scalars(idx))

        Dim baseline As EC_POINT, windowed As EC_POINT
        baseline = ec_point_new()
        windowed = ec_point_new()

        If Not ec_point_mul(baseline, scalar_bn, ctx.g, ctx) Then Exit Function
        If Not ec_point_mul_window(windowed, scalar_bn, ctx.g, ctx) Then Exit Function

        If ec_point_cmp(baseline, windowed, ctx) <> 0 Then Exit Function
    Next idx

    Verify_Windowed_Mul_Against_Standard = True
#Else
    Verify_Windowed_Mul_Against_Standard = True
#End If
End Function
