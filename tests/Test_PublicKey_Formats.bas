Attribute VB_Name = "Test_PublicKey_Formats"
Option Explicit

'==============================================================================
' TESTES DE FORMATOS DE CHAVE PÚBLICA SECP256K1
'==============================================================================
'
' PROPÓSITO:
' • Validação de formatos comprimido e descomprimido de chaves públicas
' • Testes de conversão entre formatos (compressão/descompressão)
' • Verificação de geração de endereços Bitcoin diferentes
' • Validação de assinatura/verificação com ambos os formatos
' • Teste de compatibilidade com mensagens problemáticas
'
' CARACTERÍSTICAS TÉCNICAS:
' • Formato comprimido: 33 bytes (02/03 + coordenada X)
' • Formato descomprimido: 65 bytes (04 + coordenada X + coordenada Y)
' • Hash160: SHA-256 seguido de RIPEMD-160
' • Endereços: Diferentes para formatos comprimido/descomprimido
' • Assinaturas: Compatíveis com ambos os formatos
'
' ALGORITMOS TESTADOS:
' • secp256k1_public_key_from_private() - Geração com formato específico
' • secp256k1_compress_public_key() - Compressão de chave
' • secp256k1_uncompress_public_key() - Descompressão de chave
' • bitcoin_hash160() - Hash160 para endereços
' • secp256k1_sign/verify() - Assinatura com ambos formatos
'
' TESTES IMPLEMENTADOS:
' • Geração de chaves em ambos os formatos
' • Validação de comprimento e prefixos corretos
' • Conversão bidirecional entre formatos
' • Geração de Hash160 e endereços diferentes
' • Assinatura/verificação com formatos mistos
' • Teste com mensagem anteriormente problemática
'
' FORMATOS VALIDADOS:
' • Comprimido: 66 caracteres hex, prefixo 02/03
' • Descomprimido: 130 caracteres hex, prefixo 04
' • Hash160: 40 caracteres hex para cada formato
' • Endereços: Base58 diferentes para cada formato
' • Assinaturas: DER compatível com ambos
'
' COMPATIBILIDADE:
' • Bitcoin Core secp256k1 - Formatos idênticos
' • BIP 32/44 - Padrões de derivação
' • RFC 5480 - Compressão de pontos EC
' • Base58Check - Codificação de endereços
'==============================================================================

'==============================================================================
' TESTE DE FORMATOS DE CHAVE PÚBLICA
'==============================================================================

' Propósito: Valida geração e conversão entre formatos comprimido/descomprimido
' Algoritmo: Gera chaves em ambos formatos, testa conversões e endereços
' Retorno: Relatório detalhado via Debug.Print com validações de formato

Public Sub test_public_key_formats()
    Debug.Print "=== TESTE FORMATOS CHAVE PÚBLICA ==="

    Call secp256k1_init

    ' Chave privada conhecida
    Dim private_key As String
    private_key = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"

    ' Gerar chaves públicas em ambos os formatos
    Dim compressed As String, uncompressed As String
    compressed = secp256k1_public_key_from_private(private_key, True)
    uncompressed = secp256k1_public_key_from_private(private_key, False)

    Debug.Print "Chave Privada: " & private_key
    Debug.Print "Comprimida (33 bytes): " & compressed & " (" & Len(compressed) & " chars)"
    Debug.Print "Descomprimida (65 bytes): " & uncompressed & " (" & Len(uncompressed) & " chars)"

    ' Verificar formatos
    Debug.Print "Formato comprimido OK: " & (Len(compressed) = 66 And (Left$(compressed, 2) = "02" Or Left$(compressed, 2) = "03"))
    Debug.Print "Formato descomprimido OK: " & (Len(uncompressed) = 130 And Left$(uncompressed, 2) = "04")

    ' Testar conversões
    Dim compressed_from_uncompressed As String, uncompressed_from_compressed As String
    compressed_from_uncompressed = secp256k1_compress_public_key(uncompressed)
    uncompressed_from_compressed = secp256k1_uncompress_public_key(compressed)

    Debug.Print "Conversão compressão OK: " & (compressed = compressed_from_uncompressed)
    Debug.Print "Conversão descompressão OK: " & (uncompressed = uncompressed_from_compressed)

    ' Testar endereços com ambos os formatos
    Dim hash160_compressed As String, hash160_uncompressed As String
    hash160_compressed = Hash160_VBA.Hash160_Hex(compressed)
    hash160_uncompressed = Hash160_VBA.Hash160_Hex(uncompressed)

    Debug.Print "Hash160 comprimida: " & hash160_compressed
    Debug.Print "Hash160 descomprimida: " & hash160_uncompressed
    Debug.Print "Hash160 diferentes: " & (hash160_compressed <> hash160_uncompressed)

    ' Gerar endereços Legacy para ambos
    Dim addr_compressed As String, addr_uncompressed As String
    ' Converter hash160 para bytes e usar Base58_VBA
    Dim hash160_bytes_comp() As Byte, hash160_bytes_uncomp() As Byte
    Dim i As Long
    
    ReDim hash160_bytes_comp(0 To 19)
    For i = 0 To 19
        hash160_bytes_comp(i) = CByte("&H" & Mid(hash160_compressed, i * 2 + 1, 2))
    Next
    
    ReDim hash160_bytes_uncomp(0 To 19)
    For i = 0 To 19
        hash160_bytes_uncomp(i) = CByte("&H" & Mid(hash160_uncompressed, i * 2 + 1, 2))
    Next
    
    addr_compressed = Base58_VBA.Base58Check_Encode(0, hash160_bytes_comp)
    addr_uncompressed = Base58_VBA.Base58Check_Encode(0, hash160_bytes_uncomp)

    Debug.Print "Endereço comprimido: " & addr_compressed
    Debug.Print "Endereço descomprimido: " & addr_uncompressed
    Debug.Print "Endereços diferentes: " & (addr_compressed <> addr_uncompressed)

    Debug.Print "=== TESTE FORMATOS CONCLUÍDO ==="
End Sub

'==============================================================================
' TESTE DE ASSINATURA COM AMBOS OS FORMATOS
'==============================================================================

' Propósito: Valida assinatura/verificação com chaves em ambos os formatos
' Algoritmo: Testa mensagem problemática, verifica com formatos mistos
' Retorno: Relatório detalhado via Debug.Print com validações de compatibilidade

Public Sub test_signature_with_both_formats()
    Debug.Print "=== TESTE ASSINATURA COM AMBOS FORMATOS ==="

    Call secp256k1_init

    ' Testar com a mensagem que estava falhando antes da correção
    Dim private_key As String, message As String, hash As String
    private_key = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"
    message = "Test message for both key formats"  ' Mensagem que estava falhando
    hash = SHA256_VBA.SHA256_String(message)

    Debug.Print "Testando com mensagem anteriormente problemática..."

    ' Gerar assinatura
    Dim signature As String
    signature = secp256k1_sign(hash, private_key)

    ' Testar verificação com chave comprimida
    Dim compressed As String, verify_compressed As Boolean
    compressed = secp256k1_public_key_from_private(private_key, True)
    verify_compressed = secp256k1_verify(hash, signature, compressed)

    ' Testar verificação com chave descomprimida (convertida para comprimida internamente)
    Dim uncompressed As String, compressed_converted As String, verify_uncompressed As Boolean
    uncompressed = secp256k1_public_key_from_private(private_key, False)
    compressed_converted = secp256k1_compress_public_key(uncompressed)
    verify_uncompressed = secp256k1_verify(hash, signature, compressed_converted)

    Debug.Print "Mensagem: " & message
    Debug.Print "Hash: " & hash
    Debug.Print "Assinatura: " & Left$(signature, 32) & "..."
    Debug.Print "Chave comprimida: " & compressed
    Debug.Print "Chave convertida: " & compressed_converted
    Debug.Print "Chaves coincidem: " & (compressed = compressed_converted)
    Debug.Print "Verificação com comprimida: " & verify_compressed
    Debug.Print "Verificação com descomprimida (convertida): " & verify_uncompressed
    Debug.Print "Ambas verificações coincidem: " & (verify_compressed = verify_uncompressed)

    ' Teste simples de verificação
    Dim test_verify As Boolean
    test_verify = secp256k1_verify(hash, signature, compressed)
    Debug.Print "Teste verificação direta: " & test_verify

    Debug.Print "=== TESTE ASSINATURA CONCLUÍDO ==="
End Sub

'==============================================================================
' EXECUÇÃO DE TODOS OS TESTES DE FORMATOS
'==============================================================================

' Propósito: Executa bateria completa de testes de formatos de chave pública
' Algoritmo: Chama test_public_key_formats() e test_signature_with_both_formats()
' Retorno: Relatório consolidado via Debug.Print

Public Sub test_all_public_key_formats()
    Debug.Print "=== EXECUTANDO TODOS OS TESTES FORMATOS ==="

    Call test_public_key_formats()
    Debug.Print ""
    Call test_signature_with_both_formats()

    Debug.Print "=== TODOS OS TESTES FORMATOS CONCLUÍDOS ==="
End Sub