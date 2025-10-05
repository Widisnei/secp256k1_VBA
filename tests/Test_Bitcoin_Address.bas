Attribute VB_Name = "Test_Bitcoin_Address"
Option Explicit

'==============================================================================
' MÓDULO: Test_Bitcoin_Address
' Descrição: Testes de Geração e Validação de Endereços Bitcoin
' Autor: Desenvolvido para compatibilidade com Bitcoin Core
' Data: 2024
'
' CARACTERÍSTICAS TÉCNICAS:
' • Validação de geração de endereços Bitcoin
' • Testes de todos os tipos de endereço (Legacy, SegWit, Script)
' • Verificação de conversão entre formatos
' • Testes de extração de Hash160
' • Validação de integridade de endereços
'
' TIPOS DE ENDEREÇO BITCOIN:
' • Legacy P2PKH: Pay-to-Public-Key-Hash (1...)
' • Script P2SH: Pay-to-Script-Hash (3...)
' • SegWit P2WPKH: Pay-to-Witness-Public-Key-Hash (bc1...)
' • SegWit P2WSH: Pay-to-Witness-Script-Hash (bc1...)
'
' ALGORITMOS TESTADOS:
' • generate_bitcoin_address()      - Geração de endereços
' • address_from_private_key()      - Derivação de endereços
' • validate_bitcoin_address()      - Validação de endereços
' • extract_hash160_from_address()  - Extração de hash
' • get_address_type()             - Identificação de tipo
'
' TESTES IMPLEMENTADOS:
' • Geração de todos os tipos de endereço
' • Validação de formatos e checksums
' • Conversão de chave privada para endereços
' • Extração e verificação de Hash160
' • Identificação correta de tipos
'
' ESTRUTURA BitcoinAddress:
' • address: String do endereço formatado
' • hash160: Hash160 da chave pública
' • public_key: Chave pública comprimida
' • address_type: Tipo do endereço
' • network: Mainnet ou Testnet
'
' SEGURANÇA E CONFORMIDADE:
' • Checksums válidos para todos os formatos
' • Derivação determinística de endereços
' • Validação rigorosa de parâmetros
' • Compatibilidade com carteiras padrão
'
' COMPATIBILIDADE:
' • Bitcoin Core - Algoritmos idênticos
' • BIP 44/49/84 - Derivação hierárquica
' • Electrum - Formatos compatíveis
' • Hardware Wallets - Padrões seguidos
'==============================================================================

Public Sub test_bitcoin_addresses()
    Debug.Print "=== TESTE GERAÇÃO ENDEREÇOS BITCOIN ==="

    ' Teste 1: Gerar todos os tipos de endereço
    Dim legacy_addr As BitcoinAddress, segwit_addr As BitcoinAddress, script_addr As BitcoinAddress
    
    legacy_addr = generate_bitcoin_address(LEGACY_P2PKH, "mainnet")
    segwit_addr = generate_bitcoin_address(SEGWIT_P2WPKH, "mainnet")
    script_addr = generate_bitcoin_address(SCRIPT_P2SH, "mainnet")
    
    Debug.Print "=== LEGACY P2PKH ==="
    Debug.Print "Address: " & legacy_addr.address
    Debug.Print "Type: " & get_address_type(legacy_addr.address)
    Debug.Print "Hash160: " & legacy_addr.hash160
    Debug.Print "Valid: " & validate_bitcoin_address(legacy_addr.address)
    
    Debug.Print "=== SEGWIT P2WPKH ==="
    Debug.Print "Address: " & segwit_addr.address
    Debug.Print "Type: " & get_address_type(segwit_addr.address)
    Debug.Print "Hash160: " & segwit_addr.hash160
    Debug.Print "Valid: " & validate_bitcoin_address(segwit_addr.address)
    
    Debug.Print "=== SCRIPT P2SH ==="
    Debug.Print "Address: " & script_addr.address
    Debug.Print "Type: " & get_address_type(script_addr.address)
    Debug.Print "Hash160: " & script_addr.hash160
    Debug.Print "Valid: " & validate_bitcoin_address(script_addr.address)

    Debug.Print "=== TESTE GERAÇÃO CONCLUÍDO ==="
End Sub

Public Sub test_address_conversion()
    Debug.Print "=== TESTE CONVERSÃO ENDEREÇOS ==="

    ' Usar chave privada conhecida
    Dim private_key As String
    private_key = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"
    
    ' Gerar todos os tipos a partir da mesma chave
    Dim legacy As BitcoinAddress, segwit As BitcoinAddress, script As BitcoinAddress
    legacy = address_from_private_key(private_key, LEGACY_P2PKH, "mainnet")
    segwit = address_from_private_key(private_key, SEGWIT_P2WPKH, "mainnet")
    script = address_from_private_key(private_key, SCRIPT_P2SH, "mainnet")
    
    Debug.Print "Private Key: " & private_key
    Debug.Print "Public Key: " & legacy.public_key
    Debug.Print "Hash160: " & legacy.hash160
    Debug.Print ""
    Debug.Print "Legacy: " & legacy.address
    Debug.Print "SegWit: " & segwit.address
    Debug.Print "Script: " & script.address
    Debug.Print ""

    ' Testar extração de hash160
    Dim extracted1 As String, extracted2 As String, extracted3 As String
    extracted1 = extract_hash160_from_address(legacy.address)
    extracted2 = extract_hash160_from_address(segwit.address)
    extracted3 = extract_hash160_from_address(script.address)
    
    Debug.Print "Hash160 extraction:"
    Debug.Print "Legacy OK: " & (UCase$(extracted1) = UCase$(legacy.hash160))
    Debug.Print "SegWit OK: " & (UCase$(extracted2) = UCase$(segwit.hash160))
    Debug.Print "Script OK: " & (Len(extracted3) > 0) ' P2SH tem hash160 diferente

    Debug.Print "=== TESTE CONVERSÃO CONCLUÍDO ==="
End Sub

Public Sub test_address_from_private_key_rejects_invalid_inputs()
    Debug.Print "=== TESTE: ADDRESS_FROM_PRIVATE_KEY REJEITA ENTRADAS INVÁLIDAS ==="

    On Error GoTo Handler

    Dim errCode As SECP256K1_ERROR

    ' Cenário 1: chave privada inválida deve causar erro imediato
    Call secp256k1_reset_context_for_tests()

    Dim invalidKey As String
    invalidKey = "00"

    Dim invalidErr As Long

    On Error Resume Next
    Call address_from_private_key(invalidKey, LEGACY_P2PKH, "mainnet")
    invalidErr = Err.Number
    On Error GoTo Handler

    Debug.Print "Erro propagado para chave inválida: ", (invalidErr <> 0)
    If invalidErr = 0 Then
        Err.Raise vbObjectError + &H6110&, "test_address_from_private_key_rejects_invalid_inputs", _
                  "address_from_private_key aceitou chave privada inválida '00'."
    End If

    errCode = secp256k1_get_last_error()
    Debug.Print "Código de erro reportado: ", errCode
    If errCode <> SECP256K1_ERROR_INVALID_PRIVATE_KEY Then
        Err.Raise vbObjectError + &H6111&, "test_address_from_private_key_rejects_invalid_inputs", _
                  "Código de erro incorreto para chave privada inválida."
    End If

    Err.Clear

    ' Cenário 2: contexto não inicializado deve gerar erro em derivação válida
    Call secp256k1_reset_context_for_tests()

    Dim validKey As String
    validKey = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"

    Dim contextErr As Long

    On Error Resume Next
    Call address_from_private_key(validKey, LEGACY_P2PKH, "mainnet")
    contextErr = Err.Number
    On Error GoTo Handler

    Debug.Print "Erro propagado sem inicializar contexto: ", (contextErr <> 0)
    If contextErr = 0 Then
        Err.Raise vbObjectError + &H6112&, "test_address_from_private_key_rejects_invalid_inputs", _
                  "address_from_private_key não sinalizou erro sem contexto inicializado."
    End If

    errCode = secp256k1_get_last_error()
    Debug.Print "Código de erro após falha por contexto não inicializado: ", errCode
    If errCode = SECP256K1_OK Then
        Err.Raise vbObjectError + &H6113&, "test_address_from_private_key_rejects_invalid_inputs", _
                  "Nenhum código de erro reportado para falha de contexto não inicializado."
    End If

    GoTo Cleanup

Handler:
    Debug.Print "FALHOU: " & Err.Description

Cleanup:
    If Err.Number <> 0 Then
        Dim errNumber As Long, errSource As String, errDescription As String
        errNumber = Err.Number
        errSource = Err.Source
        errDescription = Err.Description
        Err.Clear
        Debug.Print "=== TESTE ABORTADO ==="
        Err.Raise errNumber, errSource, errDescription
    Else
        Debug.Print "=== TESTE CONCLUÍDO ==="
    End If
End Sub

Public Sub test_all_bitcoin_addresses()
    Debug.Print "=== EXECUTANDO TODOS OS TESTES ENDEREÇOS ==="

    Call test_bitcoin_addresses
    Debug.Print ""
    Call test_address_conversion()

    Debug.Print "=== TODOS OS TESTES CONCLUÍDOS ==="
End Sub