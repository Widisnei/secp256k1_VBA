Attribute VB_Name = "Test_RFC6979_Rejection"
Option Explicit

'==============================================================================
' TESTE DE REGRESSÃO RFC6979 COM REJEIÇÃO FORÇADA
'==============================================================================
' Objetivo:
'   • Validar que generate_k_rfc6979 reinicia o acumulador de bytes "T" após uma
'     rejeição de candidato e gera um novo k determinístico válido.
' Metodologia:
'   • Usa gancho de teste (RFC6979_Test_RejectNextCandidates) para forçar a
'     rejeição do primeiro candidato k.
'   • Verifica que um segundo candidato é gerado e que uma assinatura ECDSA é
'     produzida com sucesso.
'
Public Sub test_rfc6979_rejection()
    Debug.Print "=== TESTE RFC6979 REJECTION ==="

    Call secp256k1_init

    ' Forçar a rejeição do primeiro candidato k
    EC_secp256k1_ECDSA.RFC6979_Test_RejectNextCandidates = 1

    Dim message As String
    message = "Regression test for RFC6979 rejection"

    Dim message_hash As String
    message_hash = SHA256_VBA.SHA256_String(message)

    Dim private_key As String
    private_key = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"

    Dim signature As String
    signature = secp256k1_sign(message_hash, private_key)

    Debug.Print "Rejeições forçadas: " & EC_secp256k1_ECDSA.RFC6979_Test_Rejections
    Debug.Print "Assinatura gerada: " & (signature <> "")

    If EC_secp256k1_ECDSA.RFC6979_Test_Rejections < 1 Then
        Err.Raise vbObjectError + &H2000&, "test_rfc6979_rejection", _
                  "Nenhum candidato foi rejeitado durante o teste."
    End If

    If signature = "" Then
        Err.Raise vbObjectError + &H2001&, "test_rfc6979_rejection", _
                  "Assinatura vazia após rejeição forçada de k."
    End If

    ' Garantir que o gancho não permaneça ativo após o teste
    EC_secp256k1_ECDSA.RFC6979_Test_RejectNextCandidates = 0
    EC_secp256k1_ECDSA.RFC6979_Test_Rejections = 0

    Debug.Print "=== TESTE RFC6979 CONCLUÍDO ==="
End Sub
