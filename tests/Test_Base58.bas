Attribute VB_Name = "Test_Base58"
Option Explicit

'==============================================================================
' MÓDULO: Test_Base58
' Descrição: Testes de Codificação Base58 e Base58Check para Bitcoin
' Autor: Desenvolvido para compatibilidade com Bitcoin Core
' Data: 2024
'
' CARACTERÍSTICAS TÉCNICAS:
' • Validação de codificação/decodificação Base58
' • Testes de Base58Check com checksum
' • Verificação de endereços Bitcoin Legacy
' • Testes de roundtrip (ida e volta)
' • Validação de integridade de dados
'
' CODIFICAÇÃO BASE58:
' • Alfabeto: 123456789ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz
' • Exclui: 0 (zero), O (maiúsculo), I (maiúsculo), l (minúsculo)
' • Propósito: Evitar confusão visual em endereços
' • Uso: Endereços Bitcoin Legacy, chaves WIF
'
' BASE58CHECK:
' • Formato: Base58(payload + checksum)
' • Checksum: Primeiros 4 bytes de SHA256(SHA256(payload))
' • Detecção: Erros de digitação e transmissão
' • Segurança: Probabilidade de erro não detectado: 1/2³²
'
' ALGORITMOS TESTADOS:
' • base58_encode()              - Codificação Base58 pura
' • base58_decode()              - Decodificação Base58 pura
' • base58check_encode()         - Codificação com checksum
' • base58check_decode()         - Decodificação com verificação
' • hash160_to_bitcoin_address() - Geração de endereços
' • bitcoin_address_to_hash160() - Extração de hash
' • validate_base58_address()    - Validação de endereços
'
' TESTES IMPLEMENTADOS:
' • Encoding/decoding básico com roundtrip
' • Base58Check com verificação de checksum
' • Geração e validação de endereços Bitcoin
' • Extração de Hash160 de endereços
' • Validação de integridade
'
' COMPATIBILIDADE:
' • Bitcoin Core - Algoritmos idênticos
' • BIP 13 - Endereços P2SH
' • Satoshi Client - Formato original
' • Electrum - Compatível
'==============================================================================

'==============================================================================
' TESTE DE CODIFICAÇÃO BASE58
'==============================================================================

' Propósito: Valida sistema de codificação Base58 para endereços Bitcoin
' Algoritmo: 4 testes cobrindo encoding, checksum, endereços e validação
' Retorno: Relatório de funcionalidade via Debug.Print
' Crítico: Essencial para compatibilidade com carteiras Bitcoin

Public Sub Test_Base58()
    Debug.Print "=== TESTE BASE58 ==="

    ' Teste 1: Encoding básico
    Dim hex1 As String, b58_1 As String, decoded1 As String
    hex1 = "00010966776006953D5567439E5E39F86A0D273BEED61967F6"
    Dim bytes1() As Byte
    bytes1 = HexToBytes(hex1)
    b58_1 = Base58_VBA.Base58_Encode(bytes1)
    Dim decoded_bytes() As Byte
    decoded_bytes = Base58_VBA.Base58_Decode(b58_1)
    decoded1 = BytesToHex(decoded_bytes)
    
    Debug.Print "Hex: " & hex1
    Debug.Print "Base58: " & b58_1
    Debug.Print "Decoded: " & decoded1
    Debug.Print "Round-trip OK: " & (UCase$(hex1) = UCase$(decoded1))
    
    ' Teste 2: Base58Check
    Dim hex2 As String, b58check As String, decoded2 As String
    hex2 = "00010966776006953D5567439E5E39F86A0D273BEE"
    Dim bytes2() As Byte, version As Byte
    bytes2 = HexToBytes(Mid$(hex2, 3))
    version = CByte("&H" & Left$(hex2, 2))
    b58check = Base58_VBA.Base58Check_Encode(version, bytes2)
    Dim decoded_version As Byte, decoded_payload() As Byte
    If Base58_VBA.Base58Check_Decode(b58check, decoded_version, decoded_payload) Then
        decoded2 = Right$("0" & Hex$(decoded_version), 2) & BytesToHex(decoded_payload)
    End If
    
    Debug.Print "Base58Check: " & b58check
    Debug.Print "Check OK: " & (UCase$(hex2) = UCase$(decoded2))

    ' Teste 3: Criar e testar endereço Bitcoin
    Dim test_hash160 As String, test_address As String, extracted_hash160 As String
    test_hash160 = "89ABCDEFABBAABBAABBAABBAABBAABBAABBAABBA" ' Hash160 de teste
    Dim hash_bytes() As Byte
    hash_bytes = HexToBytes(test_hash160)
    test_address = Base58_VBA.Base58Check_Encode(0, hash_bytes)
    Dim ext_version As Byte, ext_payload() As Byte
    If Base58_VBA.Base58Check_Decode(test_address, ext_version, ext_payload) Then
        extracted_hash160 = BytesToHex(ext_payload)
    End If
    
    Debug.Print "Test Hash160: " & test_hash160
    Debug.Print "Test Address: " & test_address
    Debug.Print "Extracted Hash160: " & extracted_hash160
    Debug.Print "Hash160 Round-trip OK: " & (UCase$(test_hash160) = UCase$(extracted_hash160))

    ' Teste 4: Validação de endereço
    Dim is_valid As Boolean
    Dim val_version As Byte, val_payload() As Byte
    is_valid = Base58_VBA.Base58Check_Decode(test_address, val_version, val_payload)
    Debug.Print "Address validation: " & is_valid

    Debug.Print "=== TESTE BASE58 CONCLUÍDO ==="
End Sub

Public Sub Test_Base58_HexParsing()
    Debug.Print "=== TESTE BASE58 HEX PARSING ==="

    Dim errNum As Long, errDesc As String

    On Error Resume Next
    Base58_VBA.Base58_SelfTest "123"
    errNum = Err.Number: errDesc = Err.Description
    Debug.Print "Odd-length hex rejected: " & (errNum <> 0)
    Debug.Assert errNum <> 0
    Debug.Print "Erro: " & errDesc
    Err.Clear
    On Error GoTo 0

    On Error Resume Next
    Base58_VBA.Base58_SelfTest "00GG"
    errNum = Err.Number: errDesc = Err.Description
    Debug.Print "Non-hex characters rejected: " & (errNum <> 0)
    Debug.Assert errNum <> 0
    Debug.Print "Erro: " & errDesc
    Err.Clear
    On Error GoTo 0

    On Error GoTo TestFail
    Base58_VBA.Base58_SelfTest
    On Error GoTo 0
    Debug.Print "Valid Base58 self-test completed sem erros de parsing."
    Exit Sub

TestFail:
    Debug.Print "Unexpected Base58 self-test error: " & Err.Description
    Debug.Assert False
    On Error GoTo 0
End Sub

' Funções auxiliares para conversão
Private Function HexToBytes(ByVal hexStr As String) As Byte()
    Dim n As Long, i As Long
    n = Len(hexStr) \ 2
    Dim b() As Byte
    If n > 0 Then
        ReDim b(0 To n - 1)
        For i = 0 To n - 1
            b(i) = CByte("&H" & Mid$(hexStr, 2 * i + 1, 2))
        Next
    End If
    HexToBytes = b
End Function

Private Function BytesToHex(ByRef data() As Byte) As String
    Dim i As Long, s As String
    If (Not Not data) = 0 Then BytesToHex = "": Exit Function
    For i = LBound(data) To UBound(data)
        s = s & Right$("0" & Hex$(data(i)), 2)
    Next
    BytesToHex = s
End Function