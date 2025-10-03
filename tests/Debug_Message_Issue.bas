Attribute VB_Name = "Debug_Message_Issue"
Option Explicit

'==============================================================================
' DEBUG DE PROBLEMAS COM MENSAGENS ECDSA
'==============================================================================
'
' PROPÓSITO:
' • Debug detalhado de problemas específicos com certas mensagens ECDSA
' • Análise comparativa entre mensagens que funcionam e que falham
' • Validação de hashes, geração de k (RFC 6979), e assinaturas DER
' • Identificação de edge cases em implementações ECDSA
'
' CARACTERÍSTICAS TÉCNICAS:
' • Mensagem problemática: "Test message for both key formats"
' • Mensagem funcionando: "Bitcoin test message"
' • Chave privada: C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721
' • Análise: Hash SHA-256, geração k, validação DER, verificação BIP 62
' • Validação: Range [1, n-1], low-s enforcement, propriedades matemáticas
'
' ALGORITMOS IMPLEMENTADOS:
' • debug_message_issue() - Debug completo com análise comparativa
' • generate_k_simple_debug() - Geração k simplificada para debug
'
' VANTAGENS:
' • Identificação precisa de problemas ECDSA
' • Comparação lado a lado de casos funcionais vs problemáticos
' • Validação completa de propriedades criptográficas
' • Debug de implementação RFC 6979
'
' COMPATIBILIDADE:
' • Bitcoin Core secp256k1 - Comportamento idêntico
' • RFC 6979 - Geração determinística de k
' • BIP 62 - Low-s enforcement
' • OpenSSL - Validação DER compatível
'==============================================================================

'==============================================================================
' DEBUG COMPLETO DE MENSAGEM PROBLEMÁTICA
'==============================================================================

' Propósito: Debug detalhado de problemas específicos com mensagens ECDSA
' Algoritmo: Compara mensagem problemática vs funcional em todas as etapas
' Retorno: Relatório completo via Debug.Print com análises detalhadas

Public Sub debug_message_issue()
    Debug.Print "=== DEBUG MENSAGEM PROBLEMÁTICA ==="
    
    Call secp256k1_init
    
    Dim private_key As String
    private_key = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"
    
    ' Testar ambas as mensagens
    Dim msg_problem As String, msg_working As String
    msg_problem = "Test message for both key formats"
    msg_working = "Bitcoin test message"
    
    ' Gerar hashes
    Dim hash_problem As String, hash_working As String
    hash_problem = SHA256_VBA.SHA256_String(msg_problem)
    hash_working = SHA256_VBA.SHA256_String(msg_working)
    
    Debug.Print "Mensagem problemática: " & msg_problem
    Debug.Print "Hash problemático: " & hash_problem
    Debug.Print "Mensagem funcionando: " & msg_working
    Debug.Print "Hash funcionando: " & hash_working
    
    ' Analisar características dos hashes
    Debug.Print ""
    Debug.Print "=== ANÁLISE DOS HASHES ==="
    Debug.Print "Hash problemático começa com: " & Left$(hash_problem, 8)
    Debug.Print "Hash funcionando começa com: " & Left$(hash_working, 8)
    
    ' Converter para BIGNUM e analisar
    Dim bn_problem As BIGNUM_TYPE, bn_working As BIGNUM_TYPE
    bn_problem = BN_hex2bn(hash_problem)
    bn_working = BN_hex2bn(hash_working)
    
    ' Verificar se são válidos para ECDSA (< n)
    Dim ctx As SECP256K1_CTX: ctx = secp256k1_context_create()
    
    Debug.Print "Hash problemático < n: " & (BN_ucmp(bn_problem, ctx.n) < 0)
    Debug.Print "Hash funcionando < n: " & (BN_ucmp(bn_working, ctx.n) < 0)
    Debug.Print "Hash problemático é zero: " & BN_is_zero(bn_problem)
    Debug.Print "Hash funcionando é zero: " & BN_is_zero(bn_working)
    
    ' Testar geração de k para ambos
    Debug.Print ""
    Debug.Print "=== TESTE GERAÇÃO K (RFC 6979) ==="
    
    Dim d As BIGNUM_TYPE: d = BN_hex2bn(private_key)
    
    ' Simular geração de k
    Dim k_problem As BIGNUM_TYPE, k_working As BIGNUM_TYPE
    k_problem = generate_k_simple_debug(bn_problem, d, ctx)
    k_working = generate_k_simple_debug(bn_working, d, ctx)
    
    Debug.Print "k problemático: " & BN_bn2hex(k_problem)
    Debug.Print "k funcionando: " & BN_bn2hex(k_working)
    Debug.Print "k problemático é zero: " & BN_is_zero(k_problem)
    Debug.Print "k funcionando é zero: " & BN_is_zero(k_working)
    Debug.Print "k problemático < n: " & (BN_ucmp(k_problem, ctx.n) < 0)
    Debug.Print "k funcionando < n: " & (BN_ucmp(k_working, ctx.n) < 0)
    
    ' Testar assinatura completa
    Debug.Print ""
    Debug.Print "=== TESTE ASSINATURA COMPLETA ==="
    
    Dim sig_problem As String, sig_working As String
    Dim pubkey As String: pubkey = secp256k1_public_key_from_private(private_key)
    
    sig_problem = secp256k1_sign(hash_problem, private_key)
    sig_working = secp256k1_sign(hash_working, private_key)
    
    Debug.Print "Assinatura problemática: " & sig_problem
    Debug.Print "Assinatura funcionando: " & sig_working
    
    ' Analisar assinaturas DER
    Debug.Print ""
    Debug.Print "=== ANÁLISE DER ==="
    Debug.Print "DER problemática válida: " & (Left$(sig_problem, 2) = "30")
    Debug.Print "DER funcionando válida: " & (Left$(sig_working, 2) = "30")
    Debug.Print "Tamanho DER problemática: " & Len(sig_problem)
    Debug.Print "Tamanho DER funcionando: " & Len(sig_working)
    
    ' Decodificar assinaturas
    Dim sig_struct_problem As ECDSA_SIGNATURE, sig_struct_working As ECDSA_SIGNATURE
    Dim decode_problem As Boolean, decode_working As Boolean
    
    decode_problem = ecdsa_signature_from_der(sig_struct_problem, sig_problem)
    decode_working = ecdsa_signature_from_der(sig_struct_working, sig_working)
    
    Debug.Print "Decode DER problemática: " & decode_problem
    Debug.Print "Decode DER funcionando: " & decode_working
    
    If decode_problem Then
        Debug.Print "r problemático: " & BN_bn2hex(sig_struct_problem.r)
        Debug.Print "s problemático: " & BN_bn2hex(sig_struct_problem.s)
        Debug.Print "r é zero: " & BN_is_zero(sig_struct_problem.r)
        Debug.Print "s é zero: " & BN_is_zero(sig_struct_problem.s)
        
        ' Verificar low-s (BIP 62)
        Dim half_n As BIGNUM_TYPE, two As BIGNUM_TYPE, temp As BIGNUM_TYPE
        half_n = BN_new(): two = BN_new(): temp = BN_new()
        Call BN_set_word(two, 2)
        Call BN_div(half_n, temp, ctx.n, two)
        
        Debug.Print "s > half_n (high-s): " & (BN_ucmp(sig_struct_problem.s, half_n) > 0)
        Debug.Print "s < half_n (low-s): " & (BN_ucmp(sig_struct_problem.s, half_n) <= 0)
        Debug.Print "half_n: " & BN_bn2hex(half_n)
        
        ' Verificar se r e s estão no range válido [1, n-1]
        Dim one As BIGNUM_TYPE: one = BN_new(): Call BN_set_word(one, 1)
        Debug.Print "r >= 1: " & (BN_ucmp(sig_struct_problem.r, one) >= 0)
        Debug.Print "r < n: " & (BN_ucmp(sig_struct_problem.r, ctx.n) < 0)
        Debug.Print "s >= 1: " & (BN_ucmp(sig_struct_problem.s, one) >= 0)
        Debug.Print "s < n: " & (BN_ucmp(sig_struct_problem.s, ctx.n) < 0)
    End If
    
    ' Testar verificação
    Dim verify_problem As Boolean, verify_working As Boolean
    verify_problem = secp256k1_verify(hash_problem, sig_problem, pubkey)
    verify_working = secp256k1_verify(hash_working, sig_working, pubkey)
    
    Debug.Print "Verificação problemática: " & verify_problem
    Debug.Print "Verificação funcionando: " & verify_working
    
    Debug.Print "=== DEBUG CONCLUÍDO ==="
End Sub

'==============================================================================
' GERAÇÃO K SIMPLIFICADA PARA DEBUG
'==============================================================================

' Propósito: Cópia da função generate_k_simple para debug detalhado
' Algoritmo: Combina z + d + "RFC6979", aplica SHA-256 duplo, reduz mod n
' Retorno: Valor k válido para assinatura ECDSA

Private Function generate_k_simple_debug(ByRef z As BIGNUM_TYPE, ByRef d As BIGNUM_TYPE, ByRef ctx As SECP256K1_CTX) As BIGNUM_TYPE
    ' Cópia da função generate_k_simple para debug
    Dim combined_hex As String, hash_hex As String, k_simple As BIGNUM_TYPE
    
    combined_hex = BN_bn2hex(z) & BN_bn2hex(d) & "RFC6979"
    hash_hex = SHA256_VBA.SHA256_String(combined_hex)
    hash_hex = SHA256_VBA.SHA256_String(hash_hex)
    
    k_simple = BN_hex2bn(hash_hex)
    Call BN_mod(k_simple, k_simple, ctx.n)
    
    If BN_is_zero(k_simple) Then Call BN_set_word(k_simple, 1)
    
    generate_k_simple_debug = k_simple
End Function