Attribute VB_Name = "Test_ECDSA_Fix"
Option Explicit

'==============================================================================
' TESTE DE CORREÇÃO ECDSA COM CONVERSÃO DER
'==============================================================================
'
' PROPÓSITO:
' • Teste completo de assinatura e verificação ECDSA com conversão DER
' • Validação da API de alto nível secp256k1_sign() e secp256k1_verify()
' • Verificação de integridade com mensagem correta e incorreta
' • Teste determinístico com chave privada fixa
'
' CARACTERÍSTICAS TÉCNICAS:
' • Chave privada: C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721
' • Mensagem: "Hello, secp256k1!"
' • API: secp256k1_sign() ? DER, secp256k1_verify() ? DER
' • Formato: Assinatura DER completa com encoding ASN.1
' • Validação: Teste positivo e negativo (mensagem diferente)
'
' ALGORITMOS IMPLEMENTADOS:
' • test_ecdsa_fix() - Teste ECDSA completo com DER
'
' VANTAGENS:
' • Teste da API completa de alto nível
' • Validação de conversão DER automática
' • Teste determinístico reproduzível
' • Verificação de robustez com casos negativos
'
' COMPATIBILIDADE:
' • Bitcoin Core secp256k1 - API idêntica
' • BIP 66 - Strict DER encoding
' • RFC 3279 - ASN.1 para criptografia
' • OpenSSL - Formato DER compatível
'==============================================================================

'==============================================================================
' TESTE DE CORREÇÃO ECDSA
'==============================================================================

' Propósito: Teste completo de assinatura e verificação ECDSA com conversão DER
' Algoritmo: Usa API de alto nível com conversão DER automática
' Retorno: Relatório completo via Debug.Print com validação positiva e negativa

Public Sub test_ecdsa_fix()
    Debug.Print "=== TESTE CORREÇÃO ECDSA ==="

    Call secp256k1_init

    ' Usar chave fixa para teste determinístico
    Dim private_key As String
    private_key = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"

    Dim public_key As String
    public_key = secp256k1_public_key_from_private(private_key)
    Debug.Print "Chave pública: " & public_key

    ' Testar com mensagem simples
    Dim message As String, hash As String
    message = "Hello, secp256k1!"
    hash = SHA256_VBA.SHA256_String(message)
    Debug.Print "Hash: " & hash

    ' Assinar
    Dim signature As String
    signature = secp256k1_sign(hash, private_key)
    Debug.Print "Assinatura: " & signature

    ' Verificar
    Dim is_valid As Boolean
    is_valid = secp256k1_verify(hash, signature, public_key)
    Debug.Print "Válida: " & is_valid

    ' Teste com hash diferente (deve falhar)
    Dim wrong_hash As String
    wrong_hash = SHA256_VBA.SHA256_String("Different message")
    Dim wrong_valid As Boolean
    wrong_valid = secp256k1_verify(wrong_hash, signature, public_key)
    Debug.Print "Hash errado válido: " & wrong_valid

    Debug.Print "=== TESTE CONCLUÍDO ==="
End Sub