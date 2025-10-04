Attribute VB_Name = "EC_Precomputed_Integration"
Option Explicit
Option Compare Binary
Option Base 0

Private Const COMB_BLOCKS As Long = 11
Private Const COMB_TEETH As Long = 6
Private Const COMB_SPACING As Long = 4
Private Const COMB_POINTS As Long = 32
Private Const COMB_BITS As Long = COMB_BLOCKS * COMB_TEETH * COMB_SPACING
Private Const GEN_TABLE_BLOCK_OFFSET As Long = 1

' =============================================================================
' MÓDULO EC PRECOMPUTED INTEGRATION - TABELA PRÉ-COMPUTADA SECP256K1
' =============================================================================
'
' DESCRIÇÃO:
' Este módulo integra as tabelas pré-computadas do Bitcoin Core secp256k1
' com o sistema VBA, fornecendo multiplicação escalar otimizada do gerador.
' Implementa conversão correta entre formatos e validação de pontos.
'
' CARACTERÍSTICAS TÉCNICAS:
' - Integração com tabelas pré-computadas do Bitcoin Core
' - Método COMB com janelas de 4 bits para otimização
' - Conversão little-endian para big-endian corrigida
' - Validação rigorosa de pontos na curva secp256k1
'
' FUNCIONALIDADES:
' - Multiplicação escalar k*G otimizada usando tabelas
' - Conversão de formato Bitcoin Core para estruturas VBA
' - Fallback para multiplicação regular quando necessário
' - Testes de validação e diagnóstico
'
' COMPATIBILIDADE:
' - Baseado nas tabelas do Bitcoin Core secp256k1
' - Formato secp256k1_ge_storage (coordenadas afins)
' - Estrutura 11×32 pontos (352 entradas totais)
'
' =============================================================================

' =============================================================================
' MULTIPLICAÇÃO ESCALAR OTIMIZADA DO GERADOR
' =============================================================================

Public Function ec_generator_mul_precomputed_correct(ByRef result As EC_POINT, ByRef scalar As BIGNUM_TYPE, ByRef ctx As SECP256K1_CTX) As Boolean
    ' -------------------------------------------------------------------------
    ' PROPÓSITO:
    '   Realiza multiplicação escalar k*G usando tabelas pré-computadas
    '   do Bitcoin Core com correções de formato e validação
    ' 
    ' PARÂMETROS:
    '   result - Ponto resultante da multiplicação k*G
    '   scalar - Escalar k (256 bits)
    '   ctx - Contexto secp256k1 inicializado
    ' 
    ' ALGORITMO:
    '   - Método COMB com janelas de 4 bits
    '   - Processa 64 janelas (256 bits ÷ 4 = 64)
    '   - Usa tabelas pré-computadas quando disponível
    '   - Fallback para multiplicação regular se necessário
    ' 
    ' RETORNA:
    '   True se multiplicação foi bem-sucedida, False caso contrário
    ' -------------------------------------------------------------------------
    
    Dim scalar_norm As BIGNUM_TYPE
    Dim d As BIGNUM_TYPE
    Dim scalar_offset As BIGNUM_TYPE
    Dim ge_offset As EC_POINT
    Dim pow2 As BIGNUM_TYPE
    Dim one As BIGNUM_TYPE
    Dim diff As BIGNUM_TYPE
    Dim tmp As BIGNUM_TYPE

    scalar_norm = BN_new()
    d = BN_new()
    scalar_offset = BN_new()
    ge_offset = ec_point_new()
    pow2 = BN_new()
    one = BN_new()
    diff = BN_new()
    tmp = BN_new()

    If BN_cmp(scalar, ctx.n) >= 0 Then
        If Not BN_mod(scalar_norm, scalar, ctx.n) Then Exit Function
    Else
        Call BN_copy(scalar_norm, scalar)
    End If

    Call BN_set_word(pow2, 1)
    If Not BN_lshift(pow2, pow2, COMB_BITS) Then Exit Function
    Call BN_set_word(one, 1)
    If Not BN_sub(pow2, pow2, one) Then Exit Function
    If Not BN_rshift(diff, pow2, 1) Then Exit Function
    If Not BN_mod(diff, diff, ctx.n) Then Exit Function
    If Not BN_add(tmp, diff, one) Then Exit Function
    If Not BN_mod(scalar_offset, tmp, ctx.n) Then Exit Function

    Call ec_point_copy(ge_offset, ctx.g)
    If Not ec_point_negate(ge_offset, ge_offset, ctx) Then Exit Function

    If Not BN_mod_add(d, scalar_norm, scalar_offset, ctx.n) Then Exit Function

    Call ec_point_set_infinity(result)

    Dim comb_off As Long, block As Long, tooth As Long
    Dim bits As Long, sign As Long, absVal As Long
    Dim bit_pos As Long
    Dim add_point As EC_POINT

    For comb_off = COMB_SPACING - 1 To 0 Step -1
        For block = 0 To COMB_BLOCKS - 1
            bits = 0
            For tooth = 0 To COMB_TEETH - 1
                bit_pos = (block * COMB_TEETH + tooth) * COMB_SPACING + comb_off
                If BN_is_bit_set(d, bit_pos) Then
                    bits = bits Or (1& << tooth)
                End If
            Next tooth

            sign = (bits >> (COMB_TEETH - 1)) And 1
            absVal = (bits Xor (-sign)) And (COMB_POINTS - 1)

            If Not get_precomputed_point_fixed(block, absVal, add_point, ctx) Then Exit Function

            If sign <> 0 Then
                If Not ec_point_negate(add_point, add_point, ctx) Then Exit Function
            End If

            If result.infinity Then
                If Not ec_point_copy(result, add_point) Then Exit Function
            Else
                If Not ec_point_add(result, result, add_point, ctx) Then Exit Function
            End If
        Next block

        If comb_off > 0 And Not result.infinity Then
            If Not ec_point_double(result, result, ctx) Then Exit Function
        End If
    Next comb_off

    If result.infinity Then
        If Not ec_point_copy(result, ge_offset) Then Exit Function
    Else
        If Not ec_point_add(result, result, ge_offset, ctx) Then Exit Function
    End If

    ec_generator_mul_precomputed_correct = True
End Function

' =============================================================================
' ACESSO ÀS TABELAS PRÉ-COMPUTADAS
' =============================================================================

Private Function get_precomputed_point_fixed(ByVal block As Long, ByVal digit As Long, ByRef point As EC_POINT, ByRef ctx As SECP256K1_CTX) As Boolean
    ' -------------------------------------------------------------------------
    ' PROPÓSITO:
    '   Obtém ponto da tabela pré-computada com mapeamento corrigido
    '   para estrutura 11×32 do Bitcoin Core
    ' 
    ' PARÂMETROS:
    '   block - Índice do bloco COMB (0 a 10)
    '   digit - Índice dentro do bloco (0 a 31)
    '   point - Ponto resultante da tabela
    '   ctx - Contexto secp256k1
    ' 
    ' MAPEAMENTO:
    '   - 11 blocos × 32 pontos do modo COMB
    '   - Offset aplicado para alinhar com a tabela global
    ' 
    ' RETORNA:
    '   True se ponto foi obtido com sucesso, False caso contrário
    ' -------------------------------------------------------------------------
    
    ' -------------------------------------------------------------------------
    ' VALIDAÇÃO: Verificar parâmetros de entrada
    ' -------------------------------------------------------------------------
    If block < 0 Or block >= COMB_BLOCKS Or digit < 0 Or digit >= COMB_POINTS Then
        get_precomputed_point_fixed = False
        Exit Function
    End If
    
    ' Obter entrada da tabela
    Dim entry As String
    entry = get_gen_point(block + GEN_TABLE_BLOCK_OFFSET, digit)
    If entry = "" Then
        get_precomputed_point_fixed = False
        Exit Function
    End If
    
    ' Converter formato Bitcoin Core CORRIGIDO
    If Not convert_bitcoin_core_point(entry, point, ctx) Then
        get_precomputed_point_fixed = False
        Exit Function
    End If
    
    get_precomputed_point_fixed = True
End Function

' =============================================================================
' CONVERSÃO DE FORMATO BITCOIN CORE
' =============================================================================

Public Function convert_bitcoin_core_point(ByVal entry As String, ByRef point As EC_POINT, ByRef ctx As SECP256K1_CTX) As Boolean
    ' -------------------------------------------------------------------------
    ' PROPÓSITO:
    '   Converte entrada da tabela Bitcoin Core para estrutura EC_POINT VBA
    '   com correção de endianness e validação rigorosa
    ' 
    ' PARÂMETROS:
    '   entry - String com coordenadas em formato "x0,x1,...,x7,y0,y1,...,y7"
    '   point - Ponto resultante em coordenadas Jacobianas
    '   ctx - Contexto secp256k1 para validação
    ' 
    ' FORMATO DE ENTRADA:
    '   - 16 valores uint32 separados por vírgula
    '   - Primeiros 8: coordenada X em little-endian
    '   - Últimos 8: coordenada Y em little-endian
    ' 
    ' PROCESSO:
    '   1. Parse da string de entrada
    '   2. Conversão little-endian → big-endian
    '   3. Criação do ponto EC_POINT
    '   4. Validação se ponto está na curva
    ' 
    ' RETORNA:
    '   True se conversão e validação foram bem-sucedidas
    ' -------------------------------------------------------------------------
    
    Dim coords() As String
    coords = Split(entry, ",")

    If UBound(coords) < 15 Then
        convert_bitcoin_core_point = False
        Exit Function
    End If

    ' -------------------------------------------------------------------------
    ' CONVERSÃO: Transformar arrays uint32 em coordenadas hexadecimais
    ' CORREÇÃO: Little-endian → Big-endian para compatibilidade VBA
    ' -------------------------------------------------------------------------
    Dim x_hex As String, y_hex As String
    Dim i As Long, uint32_val As String

    ' X coordinate: primeiros 8 uint32 em ordem reversa (little-endian → big-endian)
    For i = 7 To 0 Step -1
        uint32_val = coords(i)
        ' Garantir 8 caracteres hex
        If Len(uint32_val) < 8 Then
            uint32_val = String(8 - Len(uint32_val), "0") & uint32_val
        End If
        x_hex = x_hex & uint32_val
    Next i

    ' Y coordinate: próximos 8 uint32 em ordem reversa
    For i = 7 To 0 Step -1
        uint32_val = coords(i + 8)
        ' Garantir 8 caracteres hex
        If Len(uint32_val) < 8 Then
            uint32_val = String(8 - Len(uint32_val), "0") & uint32_val
        End If
        y_hex = y_hex & uint32_val
    Next i

    ' Criar ponto
    On Error GoTo conversion_error
    point.x = BN_hex2bn(x_hex)
    point.y = BN_hex2bn(y_hex)
    Call BN_set_word(point.z, 1)
    point.infinity = False
    On Error GoTo 0

    ' -------------------------------------------------------------------------
    ' VALIDAÇÃO CRÍTICA: Verificar se ponto pertence à curva secp256k1
    ' -------------------------------------------------------------------------
    If Not ec_point_is_on_curve(point, ctx) Then
        convert_bitcoin_core_point = False
        Exit Function
    End If

    convert_bitcoin_core_point = True
    Exit Function

conversion_error:
    convert_bitcoin_core_point = False
End Function

' =============================================================================
' FUNÇÕES DE TESTE E VALIDAÇÃO
' =============================================================================

Public Sub test_fixed_conversion()
    ' -------------------------------------------------------------------------
    ' PROPÓSITO:
    '   Testa a conversão corrigida de pontos da tabela Bitcoin Core
    '   para verificar se a integração está funcionando corretamente
    ' 
    ' FUNCIONAMENTO:
    '   1. Inicializa contexto secp256k1
    '   2. Obtém primeira entrada da tabela
    '   3. Testa conversão de formato
    '   4. Valida se ponto está na curva
    '   5. Exibe resultados no Debug
    ' 
    ' SAÍDA:
    '   Relatório de teste no Debug com coordenadas e validação
    ' -------------------------------------------------------------------------
    
    Debug.Print "=== TESTE CONVERSÃO CORRIGIDA ==="
    
    Call secp256k1_init
    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()
    
    ' Testar primeira entrada
    Dim entry As String
    entry = get_gen_point(0, 1)
    Debug.Print "Testando entrada: " & Left(entry, 50) & "..."
    
    Dim point As EC_POINT
    If convert_bitcoin_core_point(entry, point, ctx) Then
        Debug.Print "✓ Conversão bem-sucedida"
        Debug.Print "X: " & BN_bn2hex(point.x)
        Debug.Print "Y: " & BN_bn2hex(point.y)
        Debug.Print "Na curva: " & ec_point_is_on_curve(point, ctx)
    Else
        Debug.Print "✗ Conversão falhou"
    End If
    
    Debug.Print "=== TESTE CONCLUÍDO ==="
End Sub