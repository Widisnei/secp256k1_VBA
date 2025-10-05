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

Public Sub test_validate_bitcoin_address_strictness()
    Debug.Print "=== TESTE: VALIDACAO RIGOROSA DE ENDERECOS ==="

    Dim deterministicKey As String
    deterministicKey = "C9AFA9D845BA75166B5C215767B1D6934E50C3DB36E89B127B8A622B120F6721"

    Dim legacyAddr As BitcoinAddress
    Dim segwitAddr As BitcoinAddress
    Dim scriptAddr As BitcoinAddress

    legacyAddr = address_from_private_key(deterministicKey, LEGACY_P2PKH, "mainnet")
    segwitAddr = address_from_private_key(deterministicKey, SEGWIT_P2WPKH, "mainnet")
    scriptAddr = address_from_private_key(deterministicKey, SCRIPT_P2SH, "mainnet")

    Debug.Print "Legacy válido aceito: ", validate_bitcoin_address(legacyAddr.address)
    If Not validate_bitcoin_address(legacyAddr.address) Then
        Err.Raise vbObjectError + &H6120&, "test_validate_bitcoin_address_strictness", _
                  "Endereço Legacy válido foi rejeitado."
    End If

    Debug.Print "SegWit válido aceito: ", validate_bitcoin_address(segwitAddr.address)
    If Not validate_bitcoin_address(segwitAddr.address) Then
        Err.Raise vbObjectError + &H6121&, "test_validate_bitcoin_address_strictness", _
                  "Endereço SegWit P2WPKH válido foi rejeitado."
    End If

    Debug.Print "P2SH válido aceito: ", validate_bitcoin_address(scriptAddr.address)
    If Not validate_bitcoin_address(scriptAddr.address) Then
        Err.Raise vbObjectError + &H6122&, "test_validate_bitcoin_address_strictness", _
                  "Endereço P2SH válido foi rejeitado."
    End If

    ' Preparar dados inválidos
    Dim wifPayload(0 To 31) As Byte
    Dim i As Long
    For i = 0 To 31
        wifPayload(i) = i
    Next i
    Dim wifKey As String
    wifKey = Base58_VBA.Base58Check_Encode(&H80, wifPayload) ' WIF (versão 0x80)

    Debug.Print "WIF rejeitado: ", Not validate_bitcoin_address(wifKey)
    If validate_bitcoin_address(wifKey) Then
        Err.Raise vbObjectError + &H6123&, "test_validate_bitcoin_address_strictness", _
                  "Chave WIF passou na validação de endereço."
    End If

    Dim randomPayload(0 To 19) As Byte
    For i = 0 To 19
        randomPayload(i) = (255 - i) And &HFF
    Next i
    Dim arbitraryBase58 As String
    arbitraryBase58 = Base58_VBA.Base58Check_Encode(&H23, randomPayload) ' Versão desconhecida

    Debug.Print "Base58Check arbitrário rejeitado: ", Not validate_bitcoin_address(arbitraryBase58)
    If validate_bitcoin_address(arbitraryBase58) Then
        Err.Raise vbObjectError + &H6124&, "test_validate_bitcoin_address_strictness", _
                  "Base58Check arbitrário foi aceito como endereço."
    End If

    Dim mixedCaseAddress As String
    mixedCaseAddress = "bc1QW508d6qejxtdg4y5r3zarvary0c5xw7kg3g4ty" ' mistura maiúscula/minúscula

    Debug.Print "Bech32 com mistura de maiúsculas/minúsculas rejeitado: ", Not validate_bitcoin_address(mixedCaseAddress)
    If validate_bitcoin_address(mixedCaseAddress) Then
        Err.Raise vbObjectError + &H6125&, "test_validate_bitcoin_address_strictness", _
                  "Bech32 com mistura de maiúsculas/minúsculas foi aceito."
    End If

    Dim witness32(0 To 31) As Byte
    For i = 0 To 31
        witness32(i) = (i * 7) And &HFF
    Next i
    Dim longBech32 As String
    longBech32 = Bech32_VBA.Bech32_SegwitEncode("bc", 0, witness32) ' Programa de 32 bytes (P2WSH)

    Debug.Print "Bech32 com programa de 32 bytes rejeitado: ", Not validate_bitcoin_address(longBech32)
    If validate_bitcoin_address(longBech32) Then
        Err.Raise vbObjectError + &H6126&, "test_validate_bitcoin_address_strictness", _
                  "Bech32 com programa de 32 bytes foi aceito, deveria ser rejeitado."
    End If

    Debug.Print "=== TESTE CONCLUÍDO ==="
End Sub

Public Sub test_address_from_public_key_rejects_invalid_inputs()
    Debug.Print "=== TESTE: ADDRESS_FROM_PUBLIC_KEY REJEITA ENTRADAS INVÁLIDAS ==="

    Dim invalid_inputs(1 To 3) As String
    invalid_inputs(1) = ""
    invalid_inputs(2) = "05" & String$(64, "0")
    invalid_inputs(3) = "0200000000000000000000000000000000000000000000000000000000000000001"

    Dim scenario As Long
    For scenario = LBound(invalid_inputs) To UBound(invalid_inputs)
        Dim result As BitcoinAddress
        Dim errNumber As Long
        Dim errDescription As String

        On Error Resume Next
        result = address_from_public_key(invalid_inputs(scenario), LEGACY_P2PKH, "mainnet")
        errNumber = Err.Number
        errDescription = Err.Description
        On Error GoTo HandleError

        Debug.Print "Cenário " & scenario & " - erro propagado: ", (errNumber <> 0)

        If errNumber = 0 Then
            Err.Raise vbObjectError + &H6120& + scenario, _
                      "test_address_from_public_key_rejects_invalid_inputs", _
                      "address_from_public_key aceitou entrada inválida sem erro."
        End If

        Dim expectedError As Long
        expectedError = vbObjectError + &H6100& + SECP256K1_ERROR_INVALID_PUBLIC_KEY
        If errNumber <> expectedError Then
            Err.Raise vbObjectError + &H6124& + scenario, _
                      "test_address_from_public_key_rejects_invalid_inputs", _
                      "Código de erro inesperado retornado: " & errNumber & " - " & errDescription
        End If

        If LenB(result.address) <> 0 Or LenB(result.hash160) <> 0 Or LenB(result.public_key) <> 0 Then
            Err.Raise vbObjectError + &H6128& + scenario, _
                      "test_address_from_public_key_rejects_invalid_inputs", _
                      "address_from_public_key retornou dados para entrada inválida."
        End If

        Err.Clear
    Next scenario

    Debug.Print "=== TESTE CONCLUÍDO ==="
    Exit Sub

HandleError:
    Debug.Print "FALHOU: " & Err.Description
    Dim errNumberOut As Long, errSourceOut As String, errDescriptionOut As String
    errNumberOut = Err.Number
    errSourceOut = Err.Source
    errDescriptionOut = Err.Description
    Err.Clear
    Err.Raise errNumberOut, errSourceOut, errDescriptionOut
End Sub

Public Sub test_all_bitcoin_addresses()
    Debug.Print "=== EXECUTANDO TODOS OS TESTES ENDEREÇOS ==="

    Call test_bitcoin_addresses
    Debug.Print ""
    Call test_address_conversion()

    Debug.Print "=== TODOS OS TESTES CONCLUÍDOS ==="
End Sub