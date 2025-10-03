Attribute VB_Name = "Test_ECDSA_Direct_Fix"
Option Explicit

'==============================================================================
' TESTE ECDSA DIRETO SEM CONVERSÃO DER
'==============================================================================
'
' PROPÓSITO:
' • Teste de assinatura e verificação ECDSA sem conversão DER
' • Validação direta usando estruturas ECDSA_SIGNATURE nativas
' • Verificação de integridade com hash correto e incorreto
' • Debug de implementação Bitcoin Core sem overhead DER
'
' CARACTERÍSTICAS TÉCNICAS:
' • Chave privada: C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721
' • Mensagem: "Hello, secp256k1!"
' • Algoritmo: ecdsa_sign_bitcoin_core() e ecdsa_verify_bitcoin_core()
' • Estrutura: ECDSA_SIGNATURE com r e s como BIGNUM_TYPE
' • Validação: Teste positivo e negativo (hash errado)
'
' ALGORITMOS IMPLEMENTADOS:
' • test_ecdsa_direct_fix() - Teste ECDSA direto completo
'
' VANTAGENS:
' • Performance superior (sem overhead DER)
' • Debug direto da implementação core
' • Validação de integridade matemática
' • Teste de robustez com casos negativos
'
' COMPATIBILIDADE:
' • Bitcoin Core secp256k1 - Algoritmo idêntico
' • RFC 6979 - Geração determinística de k
' • FIPS 186-4 - Padrão ECDSA
' • OpenSSL - Comportamento compatível
'==============================================================================

'==============================================================================
' TESTE ECDSA DIRETO
'==============================================================================

' Propósito: Teste de assinatura e verificação ECDSA sem conversão DER
' Algoritmo: Assina com ecdsa_sign_bitcoin_core, verifica com ecdsa_verify_bitcoin_core
' Retorno: Relatório completo via Debug.Print com validação positiva e negativa

Public Sub test_ecdsa_direct_fix()
    Debug.Print "=== TESTE ECDSA DIRETO (SEM DER) ==="
    
    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()
    
    Dim private_key As String, message As String, hash As String
    private_key = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"
    message = "Hello, secp256k1!"
    hash = SHA256_VBA.SHA256_String(message)
    
    ' Assinar usando Bitcoin Core
    Dim sig As ECDSA_SIGNATURE
    sig = ecdsa_sign_bitcoin_core(hash, private_key, ctx)
    
    ' Gerar chave pública
    Dim private_bn As BIGNUM_TYPE, public_key As EC_POINT
    private_bn = BN_hex2bn(private_key)
    Call ec_point_mul_generator(public_key, private_bn, ctx)
    
    Debug.Print "r: " & BN_bn2hex(sig.r)
    Debug.Print "s: " & BN_bn2hex(sig.s)
    
    ' Verificar diretamente
    Dim valid As Boolean
    valid = ecdsa_verify_bitcoin_core(hash, sig, public_key, ctx)
    Debug.Print "Verificação direta: " & valid
    
    ' Teste com hash errado
    Dim wrong_hash As String
    wrong_hash = SHA256_VBA.SHA256_String("Wrong message")
    Dim wrong_valid As Boolean
    wrong_valid = ecdsa_verify_bitcoin_core(wrong_hash, sig, public_key, ctx)
    Debug.Print "Hash errado: " & wrong_valid
    
    Debug.Print "=== TESTE DIRETO CONCLUÍDO ==="
End Sub