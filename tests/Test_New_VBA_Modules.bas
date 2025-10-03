Attribute VB_Name = "Test_New_VBA_Modules"
Option Explicit

Public Sub test_new_vba_modules()
    Debug.Print "=== TESTE NOVOS MÓDULOS VBA ==="
    
    ' Teste SHA-256 VBA
    Dim result1 As String
    result1 = SHA256_VBA.SHA256_String("")
    Debug.Print "SHA-256(''): " & result1
    Debug.Print "Esperado: E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855"
    Debug.Print "Match: " & (result1 = "E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855")
    Debug.Print ""
    
    Dim result2 As String
    result2 = SHA256_VBA.SHA256_String("abc")
    Debug.Print "SHA-256('abc'): " & result2
    Debug.Print "Esperado: BA7816BF8F01CFEA414140DE5DAE2223B00361A396177A9CB410FF61F20015AD"
    Debug.Print "Match: " & (result2 = "BA7816BF8F01CFEA414140DE5DAE2223B00361A396177A9CB410FF61F20015AD")
    Debug.Print ""
    
    ' Teste endereço Bitcoin usando novos módulos
    Dim pubkey As String
    pubkey = "025F5A8004612C4B16402796A8A043773B60FCEA2FC58983DA36ED28CD0B12C7AF"
    
    ' Usar Hash160_VBA diretamente
    Dim hash160_result As String
    hash160_result = Hash160_VBA.Hash160_Hex(pubkey)
    Debug.Print "Hash160: " & hash160_result

    ' Gerar endereço usando Base58_VBA
    Dim hash160_bytes() As Byte
    Dim i As Long
    ReDim hash160_bytes(0 To 19)
    For i = 0 To 19
        hash160_bytes(i) = CByte("&H" & Mid(hash160_result, i * 2 + 1, 2))
    Next

    Dim address As String
    address = Base58_VBA.Base58Check_Encode(0, hash160_bytes)
    Debug.Print "Endereço: " & address
    Debug.Print "Esperado: 1Pk7DQ4VGRExSxg7pwyRrF2479WB5JBsGH"
    Debug.Print "Match: " & (address = "1Pk7DQ4VGRExSxg7pwyRrF2479WB5JBsGH")
    
End Sub