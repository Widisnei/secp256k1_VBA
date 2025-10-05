Attribute VB_Name = "Bitcoin_Address_Generation"
Option Explicit

' =============================================================================
' BITCOIN ADDRESS GENERATION VBA - GERAÇÃO COMPLETA DE ENDEREÇOS
' =============================================================================
' Módulo integrador para geração de todos os tipos de endereços Bitcoin
' Suporte completo: Legacy P2PKH, SegWit P2WPKH, Script Hash P2SH
' Integra secp256k1, ECDSA, SHA-256, RIPEMD-160, Base58 e Bech32
' =============================================================================

' =============================================================================
' TIPOS E ENUMERAÇÕES
' =============================================================================
' Enumeração dos tipos de endereço Bitcoin suportados
Public Enum AddressType
    LEGACY_P2PKH = 0        ' Endereços Legacy P2PKH (começam com "1")
    SEGWIT_P2WPKH = 1       ' Endereços SegWit P2WPKH (começam com "bc1" ou "tb1")
    SCRIPT_P2SH = 2         ' Endereços Script Hash P2SH (começam com "3")
End Enum

' Estrutura completa de um endereço Bitcoin com metadados
Public Type BitcoinAddress
    address As String       ' Endereço Bitcoin final (Base58Check ou Bech32)
    address_type As AddressType ' Tipo do endereço (Legacy, SegWit, P2SH)
    hash160 As String       ' Hash160 da chave pública (40 caracteres hex)
    public_key As String    ' Chave pública comprimida (66 caracteres hex)
    private_key As String   ' Chave privada (64 caracteres hex, pode estar vazia)
    network As String       ' Rede Bitcoin ("mainnet" ou "testnet")
End Type

' =============================================================================
' FUNÇÕES PRINCIPAIS DE GERAÇÃO DE ENDEREÇOS
' =============================================================================
Public Function generate_bitcoin_address(ByVal addr_type As AddressType, Optional ByVal network As String = "mainnet") As BitcoinAddress
    ' Gera endereço Bitcoin completo com novo par de chaves
    ' Parâmetros: addr_type - tipo de endereço desejado, network - rede Bitcoin
    ' Retorna: Estrutura BitcoinAddress com todos os dados
    ' Processo: gerar chaves → calcular hash160 → criar endereço
    
    Dim result As BitcoinAddress
    Dim ctx As SECP256K1_CTX
    
    ' Inicializar contexto criptográfico secp256k1
    Call secp256k1_init
    ctx = secp256k1_context_create()
    
    ' Gerar novo par de chaves ECDSA criptograficamente seguro
    Dim keypair As ECDSA_KEYPAIR
    keypair = ecdsa_generate_keypair(ctx)
    
    ' Converter chaves para formato hexadecimal
    result.private_key = BN_bn2hex(keypair.private_key)           ' Chave privada (256 bits)
    result.public_key = ec_point_compress(keypair.public_key, ctx) ' Chave pública comprimida (33 bytes)
    result.network = network
    result.address_type = addr_type

    ' Calcular Hash160 usando módulos VBA
    result.hash160 = Hash160_VBA.Hash160_Hex(result.public_key)

    ' Gerar endereço final baseado no tipo selecionado
    Select Case addr_type
        Case LEGACY_P2PKH
            result.address = generate_legacy_address(result.hash160, network)   ' Endereço 1xxx
        Case SEGWIT_P2WPKH
            result.address = generate_segwit_address(result.hash160, network)   ' Endereço bc1xxx
        Case SCRIPT_P2SH
            result.address = generate_script_address(result.hash160, network)   ' Endereço 3xxx
    End Select

    generate_bitcoin_address = result
End Function

Public Function address_from_private_key(ByVal private_key_hex As String, ByVal addr_type As AddressType, Optional ByVal network As String = "mainnet") As BitcoinAddress
    ' Gera endereço Bitcoin a partir de chave privada conhecida
    ' Parâmetros: private_key_hex - chave privada em hex, addr_type - tipo de endereço
    ' Retorna: Estrutura BitcoinAddress completa
    ' Útil para restaurar endereços de chaves existentes

    Dim result As BitcoinAddress

    ' Derivar chave pública da chave privada usando secp256k1
    result.private_key = private_key_hex
    result.public_key = secp256k1_public_key_from_private(private_key_hex)
    result.network = network
    result.address_type = addr_type

    Dim errCode As SECP256K1_ERROR
    errCode = secp256k1_get_last_error()

    If (result.public_key = "") Or errCode <> SECP256K1_OK Then
        result.public_key = ""
        result.hash160 = ""
        result.address = ""

        If errCode = SECP256K1_OK Then
            errCode = SECP256K1_ERROR_INVALID_PRIVATE_KEY
        End If

        Err.Raise vbObjectError + &H6100& + errCode, _
                  "address_from_private_key", _
                  "Falha ao derivar chave pública: " & secp256k1_error_string(errCode)
    End If

    ' Calcular Hash160 usando módulos VBA
    result.hash160 = Hash160_VBA.Hash160_Hex(result.public_key)

    ' Gerar endereço
    Select Case addr_type
        Case LEGACY_P2PKH
            result.address = generate_legacy_address(result.hash160, network)
        Case SEGWIT_P2WPKH
            result.address = generate_segwit_address(result.hash160, network)
        Case SCRIPT_P2SH
            result.address = generate_script_address(result.hash160, network)
    End Select

    address_from_private_key = result
End Function

Public Function address_from_public_key(ByVal public_key_hex As String, ByVal addr_type As AddressType, Optional ByVal network As String = "mainnet") As BitcoinAddress
    ' Gera endereço Bitcoin a partir de chave pública conhecida
    ' Parâmetros: public_key_hex - chave pública em hex, addr_type - tipo de endereço
    ' Retorna: Estrutura BitcoinAddress (sem chave privada)
    ' Útil para endereços watch-only ou validação

    Dim result As BitcoinAddress

    ' Configurar dados (chave privada não disponível)
    result.private_key = ""           ' Não disponível para segurança
    result.public_key = public_key_hex
    result.network = network
    result.address_type = addr_type

    ' Calcular Hash160 usando módulos VBA
    result.hash160 = Hash160_VBA.Hash160_Hex(public_key_hex)

    ' Gerar endereço
    Select Case addr_type
        Case LEGACY_P2PKH
            result.address = generate_legacy_address(result.hash160, network)
        Case SEGWIT_P2WPKH
            result.address = generate_segwit_address(result.hash160, network)
        Case SCRIPT_P2SH
            result.address = generate_script_address(result.hash160, network)
    End Select

    address_from_public_key = result
End Function

' =============================================================================
' GERADORES ESPECÍFICOS POR TIPO DE ENDEREÇO
' =============================================================================

Private Function generate_legacy_address(ByVal hash160 As String, ByVal network As String) As String
    ' Gera endereço Legacy P2PKH (Pay-to-Public-Key-Hash)
    ' Formato: Base58Check(version_byte + hash160 + checksum)
    ' Endereços resultantes começam com "1" (mainnet) ou "m"/"n" (testnet)

    Dim version As Byte
    ' Selecionar byte de versão baseado na rede
    If network = "testnet" Then
        version = &H6F ' 0x6F = Testnet P2PKH
    Else
        version = &H0  ' 0x00 = Mainnet P2PKH
    End If

    ' Converter hash160 para bytes
    Dim hash160_bytes() As Byte
    Dim i As Long
    ReDim hash160_bytes(0 To 19)
    For i = 0 To 19
        hash160_bytes(i) = CByte("&H" & Mid(hash160, i * 2 + 1, 2))
    Next

    generate_legacy_address = Base58_VBA.Base58Check_Encode(version, hash160_bytes)
End Function

Private Function generate_segwit_address(ByVal hash160 As String, ByVal network As String) As String
    ' Gera endereço SegWit P2WPKH (Pay-to-Witness-Public-Key-Hash)
    ' Formato: Bech32(hrp + witness_version + hash160 + checksum)
    ' Endereços resultantes começam com "bc1" (mainnet) ou "tb1" (testnet)

    ' Converter hash160 para bytes
    Dim hash160_bytes() As Byte
    Dim i As Long
    ReDim hash160_bytes(0 To 19)
    For i = 0 To 19
        hash160_bytes(i) = CByte("&H" & Mid(hash160, i * 2 + 1, 2))
    Next

    Dim hrp As String
    If network = "testnet" Then
        hrp = "tb"
    Else
        hrp = "bc"
    End If

    generate_segwit_address = Bech32_VBA.Bech32_SegwitEncode(hrp, 0, hash160_bytes)
End Function

Private Function generate_script_address(ByVal hash160 As String, ByVal network As String) As String
    ' Gera endereço P2SH-P2WPKH (Pay-to-Script-Hash wrapping Pay-to-Witness-Public-Key-Hash)
    ' Processo: criar script P2WPKH → calcular hash160 do script → Base58Check
    ' Endereços resultantes começam com "3" (mainnet) ou "2" (testnet)
    
    ' Construir script P2WPKH: OP_0 <20-byte-pubkey-hash>
    Dim script_hex As String
    script_hex = "0014" & hash160 ' OP_0 (0x00) + PUSH20 (0x14) + hash160

    ' Calcular Hash160 do script usando módulos VBA
    Dim script_hash160 As String
    script_hash160 = Hash160_VBA.Hash160_Hex(script_hex)

    ' Selecionar byte de versão P2SH baseado na rede
    Dim version As Byte
    If network = "testnet" Then
        version = &HC4 ' 0xC4 = Testnet P2SH
    Else
        version = &H5  ' 0x05 = Mainnet P2SH
    End If

    ' Converter script_hash160 para bytes
    Dim script_hash_bytes() As Byte
    Dim i As Long
    ReDim script_hash_bytes(0 To 19)
    For i = 0 To 19
        script_hash_bytes(i) = CByte("&H" & Mid(script_hash160, i * 2 + 1, 2))
    Next

    generate_script_address = Base58_VBA.Base58Check_Encode(version, script_hash_bytes)
End Function

' =============================================================================
' FUNÇÕES DE VALIDAÇÃO E UTILITÁRIOS
' =============================================================================

Public Function validate_bitcoin_address(ByVal address As String) As Boolean
    ' Valida qualquer tipo de endereço Bitcoin (Legacy, SegWit, P2SH)
    ' Parâmetro: address - endereço Bitcoin a ser validado
    ' Retorna: True se formato válido, False caso contrário
    ' Testa todos os formatos suportados automaticamente

    ' Tentar validação usando Base58_VBA
    Dim version As Byte
    Dim payload() As Byte
    If Base58_VBA.Base58Check_Decode(address, version, payload) Then
        validate_bitcoin_address = True
        Exit Function
    End If

    ' Tentar validação usando Bech32_VBA
    Dim hrp As String, witness_version As Byte, witness_program() As Byte
    If Bech32_VBA.Bech32_SegwitDecode(address, hrp, witness_version, witness_program) Then
        validate_bitcoin_address = True
        Exit Function
    End If

    validate_bitcoin_address = False
End Function

Public Function get_address_type(ByVal address As String) As String
    ' Identifica o tipo de endereço Bitcoin baseado no prefixo
    ' Parâmetro: address - endereço Bitcoin
    ' Retorna: Descrição do tipo de endereço
    ' Analisa prefixos para determinar formato e rede

    If Left$(address, 1) = "1" Then
        get_address_type = "Legacy P2PKH (Mainnet)"
    ElseIf Left$(address, 1) = "3" Then
        get_address_type = "Script Hash P2SH (Mainnet)"
    ElseIf Left$(address, 3) = "bc1" Then
        get_address_type = "SegWit P2WPKH (Mainnet)"
    ElseIf Left$(address, 3) = "tb1" Then
        get_address_type = "SegWit P2WPKH (Testnet)"
    ElseIf Left$(address, 1) = "m" Or Left$(address, 1) = "n" Then
        get_address_type = "Legacy P2PKH (Testnet)"
    ElseIf Left$(address, 1) = "2" Then
        get_address_type = "Script Hash P2SH (Testnet)"
    Else
        get_address_type = "Formato Desconhecido"
    End If
End Function

Public Function extract_hash160_from_address(ByVal address As String) As String
    ' Extrai hash160 de qualquer tipo de endereço Bitcoin válido
    ' Parâmetro: address - endereço Bitcoin em qualquer formato
    ' Retorna: Hash160 como string hexadecimal de 40 caracteres, ou vazio se inválido
    ' Função universal que detecta formato automaticamente

    ' Tentar como Base58Check
    Dim version As Byte
    Dim payload() As Byte
    If Base58_VBA.Base58Check_Decode(address, version, payload) Then
        If UBound(payload) = 19 Then ' 20 bytes
            Dim i As Long, result As String
            For i = 0 To 19
                result = result & Right$("0" & Hex$(payload(i)), 2)
            Next
            extract_hash160_from_address = UCase$(result)
            Exit Function
        End If
    End If

    ' Tentar como Bech32
    Dim hrp As String, witness_version As Byte, witness_program() As Byte
    If Bech32_VBA.Bech32_SegwitDecode(address, hrp, witness_version, witness_program) Then
        If UBound(witness_program) = 19 Then ' 20 bytes
            Dim j As Long, result2 As String
            For j = 0 To 19
                result2 = result2 & Right$("0" & Hex$(witness_program(j)), 2)
            Next
            extract_hash160_from_address = UCase$(result2)
            Exit Function
        End If
    End If

    extract_hash160_from_address = ""
End Function