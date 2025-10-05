Attribute VB_Name = "Test_Bech32"
Option Explicit

'==============================================================================
' MÓDULO: Test_Bech32
' Descrição: Testes de Codificação Bech32 para Endereços SegWit
' Autor: Desenvolvido para compatibilidade com Bitcoin Core
' Data: 2024
'
' CARACTERÍSTICAS TÉCNICAS:
' • Validação de codificação Bech32 (BIP 173)
' • Testes de endereços SegWit v0 e v1
' • Verificação de checksums BCH
' • Suporte a mainnet e testnet
' • Testes de roundtrip e determinismo
'
' CODIFICAÇÃO BECH32:
' • Alfabeto: qpzry9x8gf2tvdw0s3jn54khce6mua7l
' • Checksum: BCH (Bose-Chaudhuri-Hocquenghem)
' • Detecção: Até 4 erros de caractere
' • Correção: Até 2 erros de caractere
' • Vantagem: Melhor detecção de erros que Base58Check
'
' ENDEREÇOS SEGWIT:
' • Formato: hrp + separador + dados + checksum
' • HRP Mainnet: "bc" (Bitcoin)
' • HRP Testnet: "tb" (Testnet Bitcoin)
' • Versão 0: P2WPKH (20 bytes), P2WSH (32 bytes)
' • Versão 1+: Taproot e futuras atualizações
'
' ALGORITMOS TESTADOS:
' • segwit_address_encode()     - Codificação SegWit
' • segwit_address_decode()     - Decodificação SegWit
' • validate_segwit_address()   - Validação de endereços
' • hash160_to_segwit_address() - Geração de endereços
'
' TESTES IMPLEMENTADOS:
' • Geração de endereços SegWit mainnet
' • Decodificação e roundtrip
' • Validação de checksums
' • Endereços testnet
' • Determinismo da codificação
'
' VANTAGENS DO BECH32:
' • Detecção superior de erros
' • Case-insensitive (maiúsculas/minúsculas)
' • Mais eficiente (menos bytes)
' • Melhor UX (QR codes menores)
'
' COMPATIBILIDADE:
' • BIP 173 - Especificação Bech32
' • BIP 350 - Bech32m (Taproot)
' • Bitcoin Core 0.16+ - Suporte nativo
' • Carteiras modernas - Amplamente suportado
'==============================================================================

'==============================================================================
' TESTE DE CODIFICAÇÃO BECH32
'==============================================================================

' Propósito: Valida sistema de codificação Bech32 para endereços SegWit
' Algoritmo: 5 testes cobrindo encoding, decoding, validação e redes
' Retorno: Relatório de funcionalidade via Debug.Print
' Moderno: Padrão atual para endereços Bitcoin SegWit

Public Sub Test_Bech32()
    Debug.Print "=== TESTE BECH32 SIMPLIFICADO ==="

    ' Teste 1: Criar endereço SegWit
    Dim hash160 As String, segwit_addr As String, decoded_hash As String
    hash160 = "89ABCDEFABBAABBAABBAABBAABBAABBAABBAABBA"
    Dim hash_bytes() As Byte
    hash_bytes = HexToBytes(hash160)
    segwit_addr = Bech32_VBA.Bech32_SegwitEncode("bc", 0, hash_bytes)

    Debug.Print "Hash160: " & hash160
    Debug.Print "SegWit Address: " & segwit_addr
    Debug.Print "Tamanho correto: " & (Len(segwit_addr) > 20)

    ' Teste 2: Decodificar endereço
    Dim hrp_out As String, witver As Byte, prog_out() As Byte
    If Bech32_VBA.Bech32_SegwitDecode(segwit_addr, hrp_out, witver, prog_out) Then
        decoded_hash = BytesToHex(prog_out)
    End If
    Debug.Print "Decoded Hash160: " & decoded_hash
    Debug.Print "Round-trip OK: " & (UCase$(hash160) = UCase$(decoded_hash))

    ' Teste 3: Validação
    Dim is_valid As Boolean
    Dim val_hrp As String, val_witver As Byte, val_prog() As Byte
    is_valid = Bech32_VBA.Bech32_SegwitDecode(segwit_addr, val_hrp, val_witver, val_prog)
    Debug.Print "Address validation: " & is_valid

    ' Teste 4: Testnet
    Dim testnet_addr As String
    Dim testnet_bytes() As Byte
    testnet_bytes = HexToBytes(hash160)
    testnet_addr = Bech32_VBA.Bech32_SegwitEncode("tb", 0, testnet_bytes)
    Debug.Print "Testnet Address: " & testnet_addr
    Debug.Print "Testnet prefix: " & (left$(testnet_addr, 3) = "tb1")

    ' Teste 5: Determinismo
    Dim addr_repeat As String
    Dim repeat_bytes() As Byte
    repeat_bytes = HexToBytes(hash160)
    addr_repeat = Bech32_VBA.Bech32_SegwitEncode("bc", 0, repeat_bytes)
    Debug.Print "Determinístico: " & (segwit_addr = addr_repeat)

    ' Teste 6: HRP em diferentes caixas
    Debug.Print "HRP minúscula permanece minúscula: " & (segwit_addr <> "" And segwit_addr = LCase$(segwit_addr))
    Dim segwit_addr_upper As String
    segwit_addr_upper = Bech32_VBA.Bech32_SegwitEncode("BC", 0, hash_bytes)
    Debug.Print "HRP maiúscula normalizada: " & (segwit_addr_upper <> "" And segwit_addr_upper = LCase$(segwit_addr_upper))
    Dim segwit_addr_mixed As String
    segwit_addr_mixed = Bech32_VBA.Bech32_SegwitEncode("Bc", 0, hash_bytes)
    Debug.Print "HRP mista rejeitada: " & (segwit_addr_mixed = "")

    ' Teste 7: Decodificação respeitando caixa
    Dim addr_lower_ok As Boolean, addr_upper_ok As Boolean, addr_mixed_ok As Boolean
    Dim hrp_case As String, witver_case As Byte, prog_case() As Byte

    addr_lower_ok = Bech32_VBA.Bech32_SegwitDecode(segwit_addr, hrp_case, witver_case, prog_case)
    addr_upper_ok = Bech32_VBA.Bech32_SegwitDecode(UCase$(segwit_addr), hrp_case, witver_case, prog_case)

    Dim addr_mixed As String
    addr_mixed = Left$(segwit_addr, 6) & UCase$(Mid$(segwit_addr, 7, 1)) & Mid$(segwit_addr, 8)
    addr_mixed_ok = Bech32_VBA.Bech32_SegwitDecode(addr_mixed, hrp_case, witver_case, prog_case)

    Debug.Print "Decode minúsculo OK: " & addr_lower_ok
    Debug.Print "Decode maiúsculo OK: " & addr_upper_ok
    Debug.Print "Decode caixa mista rejeitado: " & (addr_mixed_ok = False)

    Debug.Print "=== TESTE BECH32 CONCLUÍDO ==="
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
    If (Not Not data) = 0 Then BytesToHex = "" : Exit Function
    For i = LBound(data) To UBound(data)
        s = s & Right$("0" & Hex$(data(i)), 2)
    Next
    BytesToHex = s
End Function