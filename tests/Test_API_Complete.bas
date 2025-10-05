Attribute VB_Name = "Test_API_Complete"
Option Explicit
Option Compare Binary
Option Base 0

'==============================================================================
' MÓDULO: Test_API_Complete
' Descrição: Testes Completos da API Pública secp256k1
' Autor: Desenvolvido para compatibilidade com Bitcoin Core
' Data: 2024
'
' CARACTERÍSTICAS TÉCNICAS:
' • Validação completa da API pública secp256k1_API.bas
' • Testes de integração com todos os módulos
' • Validação de compatibilidade Bitcoin Core
' • Testes de casos extremos e tratamento de erros
' • Verificação de performance e segurança
'
' ALGORITMOS TESTADOS:
' • secp256k1_init()                - Inicialização do contexto
' • secp256k1_generate_keypair()    - Geração segura de chaves
' • secp256k1_sign()                - Assinatura ECDSA determinística
' • secp256k1_verify()              - Verificação de assinatura
' • secp256k1_validate_*()          - Validação de chaves
' • secp256k1_point_*()             - Operações com pontos EC
'
' TESTES IMPLEMENTADOS:
' • Inicialização e configuração do sistema
' • Geração e validação de pares de chaves
' • Assinatura e verificação de mensagens
' • Operações com pontos da curva elíptica
' • Compressão e descompressão de chaves públicas
' • Tratamento de erros e casos extremos
' • Compatibilidade com Bitcoin Core
'
' VALIDAÇÕES DE SEGURANÇA:
' • Chaves privadas no range [1, n-1]
' • Chaves públicas válidas na curva secp256k1
' • Assinaturas determinísticas (RFC 6979)
' • Resistência a casos extremos
' • Validação de entrada rigorosa
'
' COMPATIBILIDADE:
' • Bitcoin Core secp256k1 - API equivalente
' • OpenSSL - Comportamento compatível
' • RFC 6979 - Assinaturas determinísticas
' • BIP 32/44 - Derivação de chaves
'==============================================================================

'==============================================================================
' EXECUÇÃO DOS TESTES COMPLETOS DA API
'==============================================================================

' Propósito: Valida completamente a API pública secp256k1
' Algoritmo: 8 suítes de teste cobrindo todas as funcionalidades
' Retorno: Relatório detalhado via Debug.Print
' Crítico: Deve passar 100% para garantir compatibilidade Bitcoin Core

Public Sub Run_API_Complete_Tests()
    Debug.Print "=== TESTES COMPLETOS DA API SECP256K1 ==="
    
    Dim passed As Long, total As Long
    
    ' Teste 1: Inicialização do sistema
    Call Test_API_Initialization(passed, total)
    
    ' Teste 2: Geração de chaves
    Call Test_API_Key_Generation(passed, total)
    
    ' Teste 3: Validação de chaves
    Call Test_API_Key_Validation(passed, total)
    
    ' Teste 4: Assinatura e verificação
    Call Test_API_Sign_Verify(passed, total)
    
    ' Teste 5: Operações com pontos
    Call Test_API_Point_Operations(passed, total)
    
    ' Teste 6: Compressão de chaves
    Call Test_API_Key_Compression(passed, total)
    
    ' Teste 7: Tratamento de erros
    Call Test_API_Error_Handling(passed, total)
    
    ' Teste 8: Compatibilidade Bitcoin Core
    Call Test_API_Bitcoin_Compatibility(passed, total)
    
    Debug.Print "=== TESTES API: ", passed, "/", total, " APROVADOS ==="
    If passed = total Then
        Debug.Print "*** API SECP256K1 COMPLETAMENTE VALIDADA ***"
    Else
        Debug.Print "*** PROBLEMAS DETECTADOS NA API ***"
    End If
End Sub

' Testa inicialização do sistema
Private Sub Test_API_Initialization(ByRef passed As Long, ByRef total As Long)
    Debug.Print "Testando inicialização da API..."
    
    ' Teste inicialização básica
    If secp256k1_init() Then
        passed = passed + 1
        Debug.Print "APROVADO: Inicialização do contexto"
    Else
        Debug.Print "FALHOU: Inicialização do contexto"
    End If
    total = total + 1
    
    ' Teste parâmetros da curva
    Dim field_prime As String, curve_order As String, generator As String
    field_prime = secp256k1_get_field_prime()
    curve_order = secp256k1_get_curve_order()
    generator = secp256k1_get_generator()
    
    If Len(field_prime) = 64 And Len(curve_order) = 64 And Len(generator) = 66 Then
        passed = passed + 1
        Debug.Print "APROVADO: Parâmetros da curva válidos"
    Else
        Debug.Print "FALHOU: Parâmetros da curva inválidos"
    End If
    total = total + 1
End Sub

' Testa geração de chaves
Private Sub Test_API_Key_Generation(ByRef passed As Long, ByRef total As Long)
    Debug.Print "Testando geração de chaves..."
    
    ' Teste geração básica
    Dim keypair As ECDSA_KEYPAIR
    keypair = secp256k1_generate_keypair()
    
    If keypair.private_key.top > 0 And Not keypair.public_key.infinity Then
        passed = passed + 1
        Debug.Print "APROVADO: Geração de par de chaves"
    Else
        Debug.Print "FALHOU: Geração de par de chaves"
    End If
    total = total + 1
    
    ' Teste derivação de chave pública
    Dim private_hex As String, public_compressed As String
    private_hex = BN_bn2hex(keypair.private_key)
    
    ' Garantir que private_hex tenha 64 caracteres
    Do While Len(private_hex) < 64
        private_hex = "0" & private_hex
    Loop
    
    public_compressed = secp256k1_public_key_from_private(private_hex, True)
    
    If Len(public_compressed) = 66 Then
        passed = passed + 1
        Debug.Print "APROVADO: Derivação de chave pública"
    Else
        Debug.Print "FALHOU: Derivação de chave pública (len=" & Len(public_compressed) & ")"
    End If
    total = total + 1
    
    ' Teste unicidade de chaves
    Dim keypair2 As ECDSA_KEYPAIR
    keypair2 = secp256k1_generate_keypair()
    
    If BN_cmp(keypair.private_key, keypair2.private_key) <> 0 Then
        passed = passed + 1
        Debug.Print "APROVADO: Unicidade de chaves"
    Else
        Debug.Print "FALHOU: Chaves duplicadas geradas"
    End If
    total = total + 1
End Sub

' Testa validação de chaves
Private Sub Test_API_Key_Validation(ByRef passed As Long, ByRef total As Long)
    Debug.Print "Testando validação de chaves..."
    
    ' Teste chave privada válida
    Dim valid_private As String
    valid_private = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"
    
    If secp256k1_validate_private_key(valid_private) Then
        passed = passed + 1
        Debug.Print "APROVADO: Validação chave privada válida"
    Else
        Debug.Print "FALHOU: Validação chave privada válida"
    End If
    total = total + 1
    
    ' Teste chave privada inválida (zero)
    If Not secp256k1_validate_private_key(String(64, "0")) Then
        passed = passed + 1
        Debug.Print "APROVADO: Rejeição chave privada zero"
    Else
        Debug.Print "FALHOU: Aceitou chave privada zero"
    End If
    total = total + 1
    
    ' Teste chave pública válida
    Dim valid_public As String
    valid_public = secp256k1_public_key_from_private(valid_private, True)
    
    If secp256k1_validate_public_key(valid_public) Then
        passed = passed + 1
        Debug.Print "APROVADO: Validação chave pública válida"
    Else
        Debug.Print "FALHOU: Validação chave pública válida"
    End If
    total = total + 1
    
    ' Teste chave pública inválida (formato incorreto)
    If Not secp256k1_validate_public_key("04" & String(64, "0")) Then
        passed = passed + 1
        Debug.Print "APROVADO: Rejeição chave pública inválida"
    Else
        Debug.Print "FALHOU: Aceitou chave pública inválida"
    End If
    total = total + 1

    ' Testes de importação de chaves privadas nos limites
    Dim imported As ECDSA_KEYPAIR
    Dim err_code As SECP256K1_ERROR
    Dim expected As BIGNUM_TYPE

    Dim min_valid As String
    min_valid = String(63, "0") & "1"
    imported = secp256k1_private_key_from_hex(min_valid)
    err_code = secp256k1_get_last_error()
    expected = BN_hex2bn(min_valid)
    If err_code = SECP256K1_OK And BN_cmp(imported.private_key, expected) = 0 Then
        passed = passed + 1
        Debug.Print "APROVADO: Importação de chave privada mínima válida"
    Else
        Debug.Print "FALHOU: Importação de chave privada mínima válida"
    End If
    total = total + 1

    Dim zero_key As String
    zero_key = String(64, "0")
    imported = secp256k1_private_key_from_hex(zero_key)
    err_code = secp256k1_get_last_error()
    If err_code = SECP256K1_ERROR_INVALID_PRIVATE_KEY And BN_is_zero(imported.private_key) Then
        passed = passed + 1
        Debug.Print "APROVADO: Rejeição de importação para chave privada zero"
    Else
        Debug.Print "FALHOU: Rejeição de importação para chave privada zero"
    End If
    total = total + 1

    Dim curve_order_hex As String
    curve_order_hex = "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEBAAEDCE6AF48A03BBFD25E8CD0364141"
    imported = secp256k1_private_key_from_hex(curve_order_hex)
    err_code = secp256k1_get_last_error()
    If err_code = SECP256K1_ERROR_INVALID_PRIVATE_KEY And BN_is_zero(imported.private_key) Then
        passed = passed + 1
        Debug.Print "APROVADO: Rejeição de importação para chave privada igual a n"
    Else
        Debug.Print "FALHOU: Rejeição de importação para chave privada igual a n"
    End If
    total = total + 1

    Dim curve_order_plus_one As String
    curve_order_plus_one = "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEBAAEDCE6AF48A03BBFD25E8CD0364142"
    imported = secp256k1_private_key_from_hex(curve_order_plus_one)
    err_code = secp256k1_get_last_error()
    If err_code = SECP256K1_ERROR_INVALID_PRIVATE_KEY And BN_is_zero(imported.private_key) Then
        passed = passed + 1
        Debug.Print "APROVADO: Rejeição de importação para chave privada maior que n"
    Else
        Debug.Print "FALHOU: Rejeição de importação para chave privada maior que n"
    End If
    total = total + 1

    Dim max_valid As String
    max_valid = "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEBAAEDCE6AF48A03BBFD25E8CD0364140"
    imported = secp256k1_private_key_from_hex(max_valid)
    err_code = secp256k1_get_last_error()
    expected = BN_hex2bn(max_valid)
    If err_code = SECP256K1_OK And BN_cmp(imported.private_key, expected) = 0 Then
        passed = passed + 1
        Debug.Print "APROVADO: Importação de chave privada máxima válida"
    Else
        Debug.Print "FALHOU: Importação de chave privada máxima válida"
    End If
    total = total + 1
End Sub

' Testa assinatura e verificação
Private Sub Test_API_Sign_Verify(ByRef passed As Long, ByRef total As Long)
    Debug.Print "Testando assinatura e verificação..."
    
    Dim private_key As String, public_key As String
    Dim message_hash As String, signature As String
    Dim err_code As Long
    
    private_key = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"
    public_key = secp256k1_public_key_from_private(private_key, True)
    message_hash = secp256k1_hash_sha256("Teste de mensagem")
    
    ' Teste assinatura
    signature = secp256k1_sign(message_hash, private_key)
    
    If Len(signature) > 8 Then
        passed = passed + 1
        Debug.Print "APROVADO: Geração de assinatura"
    Else
        Debug.Print "FALHOU: Geração de assinatura"
    End If
    total = total + 1
    
    ' Teste verificação positiva
    If secp256k1_verify(message_hash, signature, public_key) Then
        passed = passed + 1
        Debug.Print "APROVADO: Verificação de assinatura válida"
    Else
        Debug.Print "FALHOU: Verificação de assinatura válida"
    End If
    total = total + 1
    
    ' Teste verificação negativa (mensagem alterada)
    Dim wrong_hash As String
    wrong_hash = secp256k1_hash_sha256("Mensagem alterada")
    
    If Not secp256k1_verify(wrong_hash, signature, public_key) Then
        passed = passed + 1
        Debug.Print "APROVADO: Rejeição de assinatura inválida"
    Else
        Debug.Print "FALHOU: Aceitou assinatura inválida"
    End If
    total = total + 1

    ' Teste rejeição de hash inválido com caractere fora do hexadecimal
    Dim invalid_hash As String
    invalid_hash = Left$(message_hash, 63) & "G"

    If secp256k1_verify(invalid_hash, signature, public_key) Then
        Debug.Print "FALHOU: Aceitou hash inválido com caractere G"
    Else
        err_code = secp256k1_get_last_error()
        If err_code = SECP256K1_ERROR_INVALID_HASH Then
            passed = passed + 1
            Debug.Print "APROVADO: Rejeição de hash com caractere G"
        Else
            Debug.Print "FALHOU: Código de erro inesperado para hash com G"
        End If
    End If
    total = total + 1

    ' Teste rejeição de hash inválido com caractere fora do hexadecimal (Z)
    Dim invalid_hash2 As String
    invalid_hash2 = Left$(message_hash, 63) & "Z"

    If secp256k1_verify(invalid_hash2, signature, public_key) Then
        Debug.Print "FALHOU: Aceitou hash inválido com caractere Z"
    Else
        err_code = secp256k1_get_last_error()
        If err_code = SECP256K1_ERROR_INVALID_HASH Then
            passed = passed + 1
            Debug.Print "APROVADO: Rejeição de hash com caractere Z"
        Else
            Debug.Print "FALHOU: Código de erro inesperado para hash com Z"
        End If
    End If
    total = total + 1

    ' Teste determinismo (mesma entrada = mesma assinatura)
    Dim signature2 As String
    signature2 = secp256k1_sign(message_hash, private_key)
    
    If signature = signature2 Then
        passed = passed + 1
        Debug.Print "APROVADO: Determinismo de assinatura"
    Else
        Debug.Print "FALHOU: Assinatura não determinística"
    End If
    total = total + 1
End Sub

' Testa operações com pontos
Private Sub Test_API_Point_Operations(ByRef passed As Long, ByRef total As Long)
    Debug.Print "Testando operações com pontos..."
    
    ' Teste multiplicação do gerador
    Dim scalar As String, point As String
    scalar = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"
    point = secp256k1_generator_multiply(scalar)
    
    If Len(point) = 66 Then
        passed = passed + 1
        Debug.Print "APROVADO: Multiplicação do gerador"
    Else
        Debug.Print "FALHOU: Multiplicação do gerador"
    End If
    total = total + 1
    
    ' Teste multiplicação de ponto
    Dim point2 As String
    point2 = secp256k1_point_multiply("2", point)
    
    If Len(point2) = 66 Then
        passed = passed + 1
        Debug.Print "APROVADO: Multiplicação de ponto"
    Else
        Debug.Print "FALHOU: Multiplicação de ponto"
    End If
    total = total + 1
    
    ' Teste adição de pontos
    Dim point_sum As String
    point_sum = secp256k1_point_add(point, point)
    
    If point_sum = point2 Then
        passed = passed + 1
        Debug.Print "APROVADO: Adição de pontos (P + P = 2P)"
    Else
        Debug.Print "FALHOU: Adição de pontos"
    End If
    total = total + 1
End Sub

' Testa compressão de chaves
Private Sub Test_API_Key_Compression(ByRef passed As Long, ByRef total As Long)
    Debug.Print "Testando compressão de chaves..."
    
    Dim private_key As String
    private_key = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"
    
    ' Teste formato comprimido
    Dim compressed As String
    compressed = secp256k1_public_key_from_private(private_key, True)
    
    If Len(compressed) = 66 And (Left(compressed, 2) = "02" Or Left(compressed, 2) = "03") Then
        passed = passed + 1
        Debug.Print "APROVADO: Formato comprimido"
    Else
        Debug.Print "FALHOU: Formato comprimido"
    End If
    total = total + 1
    
    ' Teste formato descomprimido
    Dim uncompressed As String
    uncompressed = secp256k1_public_key_from_private(private_key, False)
    
    If Len(uncompressed) = 130 And Left(uncompressed, 2) = "04" Then
        passed = passed + 1
        Debug.Print "APROVADO: Formato descomprimido"
    Else
        Debug.Print "FALHOU: Formato descomprimido"
    End If
    total = total + 1
    
    ' Teste conversão comprimido -> descomprimido
    Dim converted As String
    converted = secp256k1_uncompress_public_key(compressed)
    
    If converted = uncompressed Then
        passed = passed + 1
        Debug.Print "APROVADO: Conversão para descomprimido"
    Else
        Debug.Print "FALHOU: Conversão para descomprimido"
    End If
    total = total + 1
    
    ' Teste conversão descomprimido -> comprimido
    Dim recompressed As String
    recompressed = secp256k1_compress_public_key(uncompressed)
    
    If recompressed = compressed Then
        passed = passed + 1
        Debug.Print "APROVADO: Conversão para comprimido"
    Else
        Debug.Print "FALHOU: Conversão para comprimido"
    End If
    total = total + 1
End Sub

' Testa tratamento de erros
Private Sub Test_API_Error_Handling(ByRef passed As Long, ByRef total As Long)
    Debug.Print "Testando tratamento de erros..."
    
    ' Teste entrada inválida para assinatura
    Dim invalid_signature As String
    invalid_signature = secp256k1_sign("hash_curto", "chave_invalida")
    
    If invalid_signature = "" Then
        passed = passed + 1
        Debug.Print "APROVADO: Rejeição entrada inválida (sign)"
    Else
        Debug.Print "FALHOU: Aceitou entrada inválida (sign)"
    End If
    total = total + 1
    
    ' Teste entrada inválida para verificação
    If Not secp256k1_verify("hash_curto", "sig_invalida", "key_invalida") Then
        passed = passed + 1
        Debug.Print "APROVADO: Rejeição entrada inválida (verify)"
    Else
        Debug.Print "FALHOU: Aceitou entrada inválida (verify)"
    End If
    total = total + 1
    
    ' Teste chave pública inválida
    If secp256k1_public_key_from_private(String(64, "F"), True) = "" Then
        passed = passed + 1
        Debug.Print "APROVADO: Rejeição chave privada inválida"
    Else
        Debug.Print "FALHOU: Aceitou chave privada inválida"
    End If
    total = total + 1

    ' Teste falha forçada na multiplicação escalar
    Dim original_force_failure As Boolean
    original_force_failure = EC_Multiplication_Dispatch.ec_point_mul_ultimate_force_failure
    EC_Multiplication_Dispatch.ec_point_mul_ultimate_force_failure = True

    Dim forced_private As String
    forced_private = String(63, "0") & "1"

    Dim forced_public As String
    forced_public = secp256k1_public_key_from_private(forced_private, True)

    Dim forced_error As SECP256K1_ERROR
    forced_error = secp256k1_get_last_error()

    If forced_public = "" And forced_error = SECP256K1_ERROR_COMPUTATION_FAILED Then
        passed = passed + 1
        Debug.Print "APROVADO: Falha na derivação de chave pública propagou erro explícito"
    Else
        Debug.Print "FALHOU: Falha na derivação de chave pública não gerou erro esperado"
        Debug.Print "  Retorno: " & forced_public
        Debug.Print "  last_error: " & forced_error
    End If
    total = total + 1

    Dim generator_failure As String
    generator_failure = secp256k1_generator_multiply(forced_private)

    forced_error = secp256k1_get_last_error()

    If generator_failure = "" And forced_error = SECP256K1_ERROR_COMPUTATION_FAILED Then
        passed = passed + 1
        Debug.Print "APROVADO: Falha na multiplicação do gerador retorna erro explícito"
    Else
        Debug.Print "FALHOU: Multiplicação do gerador não propagou falha esperada"
        Debug.Print "  Retorno: " & generator_failure
        Debug.Print "  last_error: " & forced_error
    End If
    total = total + 1

    Dim generator As String
    generator = secp256k1_get_generator()

    Dim point_mul_failure As String
    point_mul_failure = secp256k1_point_multiply(forced_private, generator)

    forced_error = secp256k1_get_last_error()

    If point_mul_failure = "" And forced_error = SECP256K1_ERROR_COMPUTATION_FAILED Then
        passed = passed + 1
        Debug.Print "APROVADO: Falha na multiplicação de ponto retorna erro explícito"
    Else
        Debug.Print "FALHOU: Multiplicação de ponto não propagou falha esperada"
        Debug.Print "  Retorno: " & point_mul_failure
        Debug.Print "  last_error: " & forced_error
    End If
    total = total + 1

    EC_Multiplication_Dispatch.ec_point_mul_ultimate_force_failure = original_force_failure
End Sub

' Testa compatibilidade Bitcoin Core
Private Sub Test_API_Bitcoin_Compatibility(ByRef passed As Long, ByRef total As Long)
    Debug.Print "Testando compatibilidade Bitcoin Core..."
    
    ' Teste com chave conhecida do Bitcoin Core
    Dim btc_private As String, btc_public As String, expected_public As String
    btc_private = "18E14A7B6A307F426A94F8114701E7C8E774E7F9A47E2C2035DB29A206321725"
    expected_public = "0250863AD64A87AE8A2FE83C1AF1A8403CB53F53E486D8511DAD8A04887E5B2352"
    
    btc_public = secp256k1_public_key_from_private(btc_private, True)
    
    If btc_public = expected_public Then
        passed = passed + 1
        Debug.Print "APROVADO: Compatibilidade chave Bitcoin Core"
    Else
        Debug.Print "FALHOU: Incompatibilidade chave Bitcoin Core"
        Debug.Print "  Esperado: ", expected_public
        Debug.Print "  Obtido:   ", btc_public
    End If
    total = total + 1
    
    ' Teste parâmetros da curva secp256k1
    Dim field_p As String, order_n As String
    field_p = secp256k1_get_field_prime()
    order_n = secp256k1_get_curve_order()
    
    Dim expected_p As String, expected_n As String
    expected_p = "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F"
    expected_n = "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEBAAEDCE6AF48A03BBFD25E8CD0364141"
    
    If field_p = expected_p And order_n = expected_n Then
        passed = passed + 1
        Debug.Print "APROVADO: Parâmetros secp256k1 Bitcoin Core"
    Else
        Debug.Print "FALHOU: Parâmetros secp256k1 incorretos"
    End If
    total = total + 1
    
    ' Teste gerador da curva
    Dim generator As String, expected_gen As String
    generator = secp256k1_get_generator()
    expected_gen = "0279BE667EF9DCBBAC55A06295CE870B07029BFCDB2DCE28D959F2815B16F81798"
    
    If generator = expected_gen Then
        passed = passed + 1
        Debug.Print "APROVADO: Gerador secp256k1 Bitcoin Core"
    Else
        Debug.Print "FALHOU: Gerador secp256k1 incorreto"
        Debug.Print "  Esperado: ", expected_gen
        Debug.Print "  Obtido:   ", generator
    End If
    total = total + 1
End Sub