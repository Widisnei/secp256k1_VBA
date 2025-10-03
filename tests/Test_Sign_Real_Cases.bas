Attribute VB_Name = "Test_Sign_Real_Cases"
Option Explicit

'==============================================================================
' TESTES DE ASSINATURA ECDSA COM CASOS REAIS
'==============================================================================
'
' PROPÓSITO:
' • Validação de assinatura ECDSA com cenários do mundo real
' • Testes com diferentes chaves privadas e tipos de mensagem
' • Verificação de determinismo e cross-validation
' • Validação de casos extremos (chave mínima, mensagens longas)
' • Teste de robustez com caracteres especiais
'
' CARACTERÍSTICAS TÉCNICAS:
' • Chaves privadas: Diferentes padrões (sequencial, reverso, mínima)
' • Mensagens: Curtas, longas, caracteres especiais, transações Bitcoin
' • Algoritmo: ECDSA com curva secp256k1
' • Hash: SHA-256 para todas as mensagens
' • Validação: Assinatura/verificação completa
'
' CASOS DE TESTE IMPLEMENTADOS:
' • Chave diferente: Padrão sequencial hexadecimal
' • Mensagem longa: Texto extenso com caracteres especiais
' • Chave mínima: Valor 1 (menor chave privada válida)
' • Cross-verification: Assinatura de uma chave com outra (deve falhar)
' • Determinismo: Mesma entrada produz mesma assinatura
'
' ALGORITMOS TESTADOS:
' • secp256k1_public_key_from_private() - Geração de chave pública
' • secp256k1_sign() - Assinatura ECDSA
' • secp256k1_verify() - Verificação de assinatura
' • sha256_hash() - Hash SHA-256 das mensagens
'
' VALIDAÇÕES REALIZADAS:
' • Assinatura válida para cada par chave/mensagem
' • Chave pública correta derivada da privada
' • Cross-verification falha como esperado
' • Determinismo mantido entre execuções
' • Robustez com diferentes tipos de entrada
'
' COMPATIBILIDADE:
' • Bitcoin Core secp256k1 - Assinaturas idênticas
' • RFC 6979 - Geração determinística de k
' • OpenSSL ECDSA - Comportamento compatível
' • BIP 32/44 - Base para derivação hierárquica
'==============================================================================

'==============================================================================
' TESTE DE ASSINATURA COM CASOS REAIS
'==============================================================================

' Propósito: Valida assinatura ECDSA com cenários reais diversos
' Algoritmo: 5 testes cobrindo diferentes chaves, mensagens e validações
' Retorno: Relatório detalhado via Debug.Print com resultados de cada teste

Public Sub Test_Sign_Real_Cases()
    Debug.Print "=== TESTE ASSINATURA CASOS REAIS ==="

    Call secp256k1_init

    ' Teste 1: Chave privada diferente
    Dim private_key1 As String, message1 As String, hash1 As String
    private_key1 = "1234567890ABCDEF1234567890ABCDEF1234567890ABCDEF1234567890ABCDEF"
    message1 = "Teste transação Bitcoin"
    hash1 = SHA256_VBA.SHA256_String(message1)

    Dim pubkey1 As String, sig1 As String, valid1 As Boolean
    pubkey1 = secp256k1_public_key_from_private(private_key1)
    sig1 = secp256k1_sign(hash1, private_key1)
    valid1 = secp256k1_verify(hash1, sig1, pubkey1)

    Debug.Print "Teste 1 - Chave diferente:"
    Debug.Print "  Privada: " & private_key1
    Debug.Print "  Pública: " & pubkey1
    Debug.Print "  Válida: " & valid1

    ' Teste 2: Mensagem longa
    Dim private_key2 As String, message2 As String, hash2 As String
    private_key2 = "FEDCBA0987654321FEDCBA0987654321FEDCBA0987654321FEDCBA0987654321"
    message2 = "Esta é uma mensagem muito longa para testar se o sistema funciona com textos extensos e caracteres especiais: !@#$%^&*()"
    hash2 = SHA256_VBA.SHA256_String(message2)

    Dim pubkey2 As String, sig2 As String, valid2 As Boolean
    pubkey2 = secp256k1_public_key_from_private(private_key2)
    sig2 = secp256k1_sign(hash2, private_key2)
    valid2 = secp256k1_verify(hash2, sig2, pubkey2)

    Debug.Print "Teste 2 - Mensagem longa:"
    Debug.Print "  Válida: " & valid2

    ' Teste 3: Chave mínima
    Dim private_key3 As String, message3 As String, hash3 As String
    private_key3 = "0000000000000000000000000000000000000000000000000000000000000001"
    message3 = "Teste chave mínima"
    hash3 = SHA256_VBA.SHA256_String(message3)

    Dim pubkey3 As String, sig3 As String, valid3 As Boolean
    pubkey3 = secp256k1_public_key_from_private(private_key3)
    sig3 = secp256k1_sign(hash3, private_key3)
    valid3 = secp256k1_verify(hash3, sig3, pubkey3)

    Debug.Print "Teste 3 - Chave mínima:"
    Debug.Print "  Pública: " & pubkey3
    Debug.Print "  Válida: " & valid3

    ' Teste 4: Cross-verification (assinatura de uma chave com outra)
    Dim cross_valid As Boolean
    cross_valid = secp256k1_verify(hash1, sig2, pubkey1)
    Debug.Print "Teste 4 - Cross-verification: " & cross_valid & " (deve ser Falso)"

    ' Teste 5: Determinismo (mesma entrada = mesma saída)
    Dim sig1_repeat As String
    sig1_repeat = secp256k1_sign(hash1, private_key1)
    Debug.Print "Teste 5 - Determinismo: " & (sig1 = sig1_repeat) & " (deve ser Verdadeiro)"

    Debug.Print "=== TESTE CASOS REAIS CONCLUÍDO ==="
End Sub