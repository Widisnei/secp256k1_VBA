Attribute VB_Name = "Debug_ECDSA_Verify"
Option Explicit

Public Sub debug_ecdsa_verification()
    Debug.Print "=== DEBUG VERIFICAÇÃO ECDSA ==="
    
    Call secp256k1_init
    
    Dim private_key As String, public_key As String
    Dim message_hash As String, signature As String
    
    private_key = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"
    public_key = secp256k1_public_key_from_private(private_key, True)
    message_hash = secp256k1_hash_sha256("Teste de mensagem")
    
    Debug.Print "Chave privada: " & private_key
    Debug.Print "Chave pública: " & public_key
    Debug.Print "Hash mensagem: " & message_hash
    
    ' Gerar assinatura
    signature = secp256k1_sign(message_hash, private_key)
    Debug.Print "Assinatura: " & signature
    
    ' Verificar assinatura
    Dim is_valid As Boolean
    is_valid = secp256k1_verify(message_hash, signature, public_key)
    Debug.Print "Verificação: " & is_valid
    
    If Not is_valid Then
        Debug.Print "*** PROBLEMA NA VERIFICAÇÃO - INVESTIGANDO ***"
        Call debug_verify_steps(message_hash, signature, public_key)
    End If
End Sub

Private Sub debug_verify_steps(ByVal msg_hash As String, ByVal sig_der As String, ByVal pub_key As String)
    Debug.Print "--- DEBUG PASSOS VERIFICAÇÃO ---"
    
    ' Verificar se chave pública é válida
    If secp256k1_validate_public_key(pub_key) Then
        Debug.Print "[OK] Chave pública válida"
    Else
        Debug.Print "[ERRO] Chave pública inválida"
        Exit Sub
    End If
    
    ' Verificar formato da assinatura
    If Len(sig_der) > 8 Then
        Debug.Print "[OK] Formato assinatura válido"
    Else
        Debug.Print "[ERRO] Formato assinatura inválido"
        Exit Sub
    End If
    
    Debug.Print "Problema pode estar na implementação da verificação ECDSA"
End Sub