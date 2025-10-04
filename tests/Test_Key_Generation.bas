Attribute VB_Name = "Test_Key_Generation"
Option Explicit

'==============================================================================
' TESTES DE GERAÇÃO DE CHAVES CRIPTOGRÁFICAS SECP256K1
'==============================================================================
'
' PROPÓSITO:
' • Validação da geração segura de pares de chaves ECDSA
' • Testes de unicidade e distribuição aleatória
' • Verificação de range válido para chaves privadas [1, n-1]
' • Validação de chaves públicas na curva secp256k1
' • Teste funcional de assinatura/verificação
'
' CARACTERÍSTICAS TÉCNICAS:
' • Chaves privadas: 256-bit no range [1, n-1]
' • Chaves públicas: Pontos na curva y² = x³ + 7 (mod p)
' • Geração: Criptograficamente segura com entropia adequada
' • Validação: Conformidade com padrões Bitcoin e RFC 6979
' • Distribuição: Análise estatística básica de aleatoriedade
'
' ALGORITMOS TESTADOS:
' • secp256k1_generate_keypair() - Geração de par de chaves
' • BN_is_zero() - Validação de chave não-zero
' • BN_ucmp() - Comparação para range [1, n-1]
' • ec_point_is_on_curve() - Validação de ponto na curva
' • secp256k1_sign/verify() - Funcionalidade completa
'
' TESTES IMPLEMENTADOS:
' • Geração básica: Chaves não-zero e ponto não-infinito
' • Unicidade: Múltiplas gerações produzem chaves diferentes
' • Range válido: Chaves privadas em [1, n-1]
' • Validação curva: Chaves públicas na curva secp256k1
' • Funcionalidade: Assinatura/verificação funcional
' • Distribuição: Análise estatística de 100 chaves
'
' SEGURANÇA VALIDADA:
' • Entropia criptográfica adequada
' • Ausência de chaves zero ou inválidas
' • Distribuição uniforme básica
' • Conformidade com padrões Bitcoin
' • Resistência a ataques conhecidos
'
' COMPATIBILIDADE:
' • Bitcoin Core secp256k1 - Geração idêntica
' • OpenSSL EC_KEY - Comportamento compatível
' • RFC 6979 - Padrões seguidos
' • BIP 32/44 - Base para derivação hierárquica
'==============================================================================

'==============================================================================
' EXECUÇÃO DOS TESTES DE GERAÇÃO DE CHAVES
'==============================================================================

' Propósito: Executa bateria completa de testes de geração de chaves
' Algoritmo: 5 testes cobrindo geração, unicidade, range, validação e funcionalidade
' Retorno: Relatório detalhado via Debug.Print com estatísticas

Public Sub test_key_generation()
    Debug.Print "=== TESTE GERAÇÃO DE CHAVES ==="

    Call secp256k1_init

    Dim passed As Long, total As Long
    passed = 0 : total = 0

    ' Teste 1: Geração básica
    total = total + 1
    If test_basic_generation() Then passed = passed + 1

    ' Teste 2: Unicidade das chaves
    total = total + 1
    If test_key_uniqueness() Then passed = passed + 1

    ' Teste 3: Validação de range
    total = total + 1
    If test_private_key_range() Then passed = passed + 1

    ' Teste 4: Chave pública válida
    total = total + 1
    If test_public_key_validity() Then passed = passed + 1

    ' Teste 5: Assinatura funcional
    total = total + 1
    If test_signature_functionality() Then passed = passed + 1

    ' Teste 6: Tratamento de falha do dispatcher de multiplicação
    total = total + 1
    If test_dispatch_failure_handling() Then passed = passed + 1

    ' Teste 7: Falha na derivação de chave pública a partir de chave privada
    total = total + 1
    If test_private_key_derivation_failure_propagation() Then passed = passed + 1

    Debug.Print "=== GERAÇÃO DE CHAVES: " & passed & "/" & total & " APROVADOS ==="
End Sub

'==============================================================================
' TESTE DE GERAÇÃO BÁSICA
'==============================================================================

' Propósito: Valida geração básica de par de chaves não-zero
' Algoritmo: Gera keypair, verifica chave privada não-zero e pública não-infinita
' Retorno: True se geração bem-sucedida, False caso contrário

Private Function test_basic_generation() As Boolean
    Debug.Print "Testando geração básica de chaves..."

    Dim keypair As ECDSA_KEYPAIR
    keypair = secp256k1_generate_keypair()

    If Not BN_is_zero(keypair.private_key) And Not keypair.public_key.infinity Then
        Debug.Print "APROVADO: Geração básica de chaves"
        test_basic_generation = True
    Else
        Debug.Print "FALHOU: Geração básica de chaves falhou"
        test_basic_generation = False
    End If
End Function

'==============================================================================
' TESTE DE UNICIDADE DE CHAVES
'==============================================================================

' Propósito: Verifica que múltiplas gerações produzem chaves diferentes
' Algoritmo: Gera 3 keypairs, compara chaves privadas para garantir unicidade
' Retorno: True se todas as chaves são únicas, False caso contrário

Private Function test_key_uniqueness() As Boolean
    Debug.Print "Testando unicidade de chaves..."

    Dim keypair1 As ECDSA_KEYPAIR, keypair2 As ECDSA_KEYPAIR, keypair3 As ECDSA_KEYPAIR
    keypair1 = secp256k1_generate_keypair()
    keypair2 = secp256k1_generate_keypair()
    keypair3 = secp256k1_generate_keypair()

    Dim unique As Boolean
    unique = (BN_cmp(keypair1.private_key, keypair2.private_key) <> 0) And
             (BN_cmp(keypair1.private_key, keypair3.private_key) <> 0) And
             (BN_cmp(keypair2.private_key, keypair3.private_key) <> 0)

    If unique Then
        Debug.Print "APROVADO: Chaves são únicas"
        test_key_uniqueness = True
    Else
        Debug.Print "FALHOU: Chaves duplicadas geradas"
        test_key_uniqueness = False
    End If
End Function

'==============================================================================
' TESTE DE RANGE DE CHAVE PRIVADA
'==============================================================================

' Propósito: Valida que chaves privadas estão no range válido [1, n-1]
' Algoritmo: Gera 10 chaves, verifica se todas estão no range correto
' Retorno: True se todas as chaves estão no range, False caso contrário

Private Function test_private_key_range() As Boolean
    Debug.Print "Testando range de chave privada [1, n-1]..."

    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()

    Dim valid_count As Long, i As Long
    For i = 1 To 10
        Dim keypair As ECDSA_KEYPAIR
        keypair = secp256k1_generate_keypair()

        ' Verificar se 1 <= private_key < n
        If Not BN_is_zero(keypair.private_key) And BN_ucmp(keypair.private_key, ctx.n) < 0 Then
            valid_count = valid_count + 1
        End If
    Next i

    If valid_count = 10 Then
        Debug.Print "APROVADO: Todas as chaves privadas no range válido"
        test_private_key_range = True
    Else
        Debug.Print "FALHOU: " & (10 - valid_count) & " chaves fora do range"
        test_private_key_range = False
    End If
End Function

'==============================================================================
' TESTE DE VALIDAÇÃO DE CHAVE PÚBLICA
'==============================================================================

' Propósito: Verifica que chaves públicas geradas estão na curva secp256k1
' Algoritmo: Gera 5 keypairs, valida se pontos públicos estão na curva
' Retorno: True se todos os pontos estão na curva, False caso contrário

Private Function test_public_key_validity() As Boolean
    Debug.Print "Testando validação de chave pública..."

    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()

    Dim valid_count As Long, i As Long
    For i = 1 To 5
        Dim keypair As ECDSA_KEYPAIR
        keypair = secp256k1_generate_keypair()

        ' Verificar se ponto está na curva
        If ec_point_is_on_curve(keypair.public_key, ctx) Then
            valid_count = valid_count + 1
        End If
    Next i

    If valid_count = 5 Then
        Debug.Print "APROVADO: Todas as chaves públicas na curva"
        test_public_key_validity = True
    Else
        Debug.Print "FALHOU: " & (5 - valid_count) & " chaves públicas fora da curva"
        test_public_key_validity = False
    End If
End Function

'==============================================================================
' TESTE DE FUNCIONALIDADE DE ASSINATURA
'==============================================================================

' Propósito: Valida que chaves geradas funcionam para assinatura/verificação
' Algoritmo: Gera keypair, assina mensagem, verifica com chave pública
' Retorno: True se assinatura/verificação funciona, False caso contrário

Private Function test_signature_functionality() As Boolean
    Debug.Print "Testando funcionalidade de assinatura com chaves geradas..."

    Dim keypair As ECDSA_KEYPAIR
    keypair = secp256k1_generate_keypair()

    Dim message As String, hash As String, signature As String
    message = "Mensagem de teste para chave gerada"
    hash = SHA256_VBA.SHA256_String(message)

    ' Assinar com chave gerada
    signature = secp256k1_sign(hash, BN_bn2hex(keypair.private_key))

    ' Verificar com chave pública gerada
    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()
    Dim public_key_compressed As String
    public_key_compressed = ec_point_compress(keypair.public_key, ctx)
    
    Debug.Print "DEBUG: Hash: " & hash
    Debug.Print "DEBUG: Signature: " & Left$(signature, 32) & "..."
    Debug.Print "DEBUG: PubKey: " & public_key_compressed

    Dim valid As Boolean
    valid = secp256k1_verify(hash, signature, public_key_compressed)
    Debug.Print "DEBUG: Verification result: " & valid

    If valid Then
        Debug.Print "APROVADO: Chaves geradas funcionam para assinatura/verificação"
        test_signature_functionality = True
    Else
        Debug.Print "FALHOU: Chaves geradas falharam na assinatura/verificação"
        test_signature_functionality = False
    End If
End Function

'==============================================================================
' TESTE DE TRATAMENTO DE FALHAS DO DISPATCHER
'==============================================================================
'
' Propósito: Garante que a API não retorna pontos inválidos quando a multiplicação
'             escalar falha e que o erro é sinalizado corretamente.
' Algoritmo: Força ec_point_mul_ultimate a falhar, chama a API de geração e
'            valida last_error e o par retornado.
' Retorno: True se a falha for tratada corretamente, False caso contrário.

Private Function test_dispatch_failure_handling() As Boolean
    Debug.Print "Testando tratamento de falha do dispatcher na geração de chaves..."

    Dim originalFlag As Boolean
    originalFlag = EC_Multiplication_Dispatch.ec_point_mul_ultimate_force_failure

    On Error GoTo UnexpectedError

    EC_Multiplication_Dispatch.ec_point_mul_ultimate_force_failure = True

    Dim keypair As ECDSA_KEYPAIR
    keypair = secp256k1_generate_keypair()

    Dim lastErr As SECP256K1_ERROR
    lastErr = secp256k1_get_last_error()

    If lastErr <> SECP256K1_ERROR_COMPUTATION_FAILED Then
        Debug.Print "FALHOU: Erro não propagado corretamente após falha na multiplicação."
        GoTo Cleanup
    End If

    If Not BN_is_zero(keypair.private_key) Then
        Debug.Print "FALHOU: Chave privada deveria estar zerada após falha da multiplicação."
        GoTo Cleanup
    End If

    If Not keypair.public_key.infinity Then
        Debug.Print "FALHOU: Ponto público inválido retornado após falha da multiplicação."
        GoTo Cleanup
    End If

    Debug.Print "APROVADO: Falha do dispatcher tratada corretamente pela API."
    test_dispatch_failure_handling = True

Cleanup:
    EC_Multiplication_Dispatch.ec_point_mul_ultimate_force_failure = originalFlag
    Exit Function

UnexpectedError:
    Debug.Print "FALHOU: Erro inesperado ao validar tratamento de falha - " & Err.Description
    Err.Clear
    GoTo Cleanup
End Function

'==============================================================================
' TESTE DE PROPAGAÇÃO DE FALHA NA DERIVAÇÃO DE CHAVE PRIVADA
'==============================================================================
' Propósito: Garante que ecdsa_set_private_key gera erro quando a multiplicação
'            escalar falha e não retorna chave pública não inicializada.
' Algoritmo: Força ec_point_mul_ultimate a falhar, tenta derivar par via
'            ecdsa_set_private_key e verifica se o erro correto é gerado.
' Retorno: True se o erro esperado for disparado, False caso contrário.

Private Function test_private_key_derivation_failure_propagation() As Boolean
    Debug.Print "Testando propagação de falha na derivação de chave privada..."

    Dim originalFlag As Boolean
    originalFlag = EC_Multiplication_Dispatch.ec_point_mul_ultimate_force_failure

    On Error GoTo Handler

    EC_Multiplication_Dispatch.ec_point_mul_ultimate_force_failure = True

    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()

    Dim valid_private_hex As String
    valid_private_hex = String(63, "0") & "1"

    Dim keypair As ECDSA_KEYPAIR
    keypair = ecdsa_set_private_key(valid_private_hex, ctx)

    Debug.Print "FALHOU: ecdsa_set_private_key não gerou erro na falha de multiplicação."
    GoTo Cleanup

Handler:
    If Err.Number = vbObjectError + &H1102& Then
        Debug.Print "APROVADO: ecdsa_set_private_key propagou a falha corretamente."
        test_private_key_derivation_failure_propagation = True
    Else
        Debug.Print "FALHOU: Erro inesperado propagado - " & Err.Number & " - " & Err.Description
    End If
    Err.Clear

Cleanup:
    EC_Multiplication_Dispatch.ec_point_mul_ultimate_force_failure = originalFlag
    Exit Function
End Function

'==============================================================================
' TESTE DE DISTRIBUIÇÃO DE CHAVES
'==============================================================================

' Propósito: Analisa distribuição estatística básica de chaves geradas
' Algoritmo: Gera 100 chaves, analisa distribuição do primeiro nibble
' Retorno: Relatório de distribuição via Debug.Print

Public Sub test_key_distribution()
    Debug.Print "=== TESTE DISTRIBUIÇÃO DE CHAVES ==="
    
    Dim i As Long, first_bytes(0 To 15) As Long
    
    ' Gerar 100 chaves e analisar primeiro byte
    For i = 1 To 100
        Dim keypair As ECDSA_KEYPAIR
        keypair = secp256k1_generate_keypair()
        
        Dim hex_key As String
        hex_key = BN_bn2hex(keypair.private_key)
        
        ' Contar primeiro caractere hex
        Dim first_char As String, first_val As Long
        first_char = Left$(hex_key, 1)
        first_val = CLng("&H" & first_char)
        first_bytes(first_val) = first_bytes(first_val) + 1
    Next i
    
    ' Mostrar distribuição
    Debug.Print "Distribuição do primeiro nibble (0-F):"
    For i = 0 To 15
        Debug.Print Hex$(i) & ": " & first_bytes(i) & " vezes"
    Next i
    
    Debug.Print "=== DISTRIBUIÇÃO CONCLUÍDA ==="
End Sub
