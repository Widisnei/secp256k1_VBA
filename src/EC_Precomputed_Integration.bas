Attribute VB_Name = "EC_Precomputed_Integration"
Option Explicit
Option Compare Binary
Option Base 0

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
    
    ' -------------------------------------------------------------------------
    ' INICIALIZAÇÃO: Configurar ponto resultado como identidade
    ' -------------------------------------------------------------------------
    result.infinity = True
    Call BN_set_word(result.x, 0)
    Call BN_set_word(result.y, 0)
    Call BN_set_word(result.z, 1)
    
    ' -------------------------------------------------------------------------
    ' PREPARAÇÃO: Extrair bits do escalar para processamento COMB
    ' -------------------------------------------------------------------------
    Dim scalar_bits(255) As Boolean
    Dim i As Long, bit_pos As Long
    
    For i = 0 To 255
        scalar_bits(i) = BN_is_bit_set(scalar, i)
    Next i
    
    ' -------------------------------------------------------------------------
    ' ALGORITMO COMB: Processar janelas de 4 bits para otimização
    ' -------------------------------------------------------------------------
    Dim window As Long, window_val As Long
    Dim temp_point As EC_POINT, add_point As EC_POINT
    
    For window = 0 To 63  ' 256 bits / 4 = 64 janelas
        window_val = 0
        
        ' Construir valor da janela de 4 bits (0-15)
        For i = 0 To 3
            bit_pos = window * 4 + i
            If bit_pos < 256 And scalar_bits(bit_pos) Then
                window_val = window_val + (2 ^ i)
            End If
        Next i
        
        ' Adicionar ponto pré-computado se janela não for zero
        If window_val > 0 Then
            If get_precomputed_point_fixed(window, window_val, add_point, ctx) Then
                If result.infinity Then
                    result = add_point
                Else
                    Call ec_point_add(result, result, add_point, ctx)
                End If
            Else
                ' Fallback: usar multiplicação regular para esta janela
                Dim window_scalar As BIGNUM_TYPE
                Call BN_set_word(window_scalar, window_val)
                Call BN_lshift(window_scalar, window_scalar, window * 4)
                Call ec_point_mul(temp_point, window_scalar, ctx.g, ctx)
                
                If result.infinity Then
                    result = temp_point
                Else
                    Call ec_point_add(result, result, temp_point, ctx)
                End If
            End If
        End If
    Next window
    
    ec_generator_mul_precomputed_correct = True
End Function

' =============================================================================
' ACESSO ÀS TABELAS PRÉ-COMPUTADAS
' =============================================================================

Private Function get_precomputed_point_fixed(ByVal window As Long, ByVal index As Long, ByRef point As EC_POINT, ByRef ctx As SECP256K1_CTX) As Boolean
    ' -------------------------------------------------------------------------
    ' PROPÓSITO:
    '   Obtém ponto da tabela pré-computada com mapeamento corrigido
    '   para estrutura 11×32 do Bitcoin Core
    ' 
    ' PARÂMETROS:
    '   window - Número da janela (0 a 63)
    '   index - Índice dentro da janela (1 a 15)
    '   point - Ponto resultante da tabela
    '   ctx - Contexto secp256k1
    ' 
    ' MAPEAMENTO:
    '   - 64 janelas teóricas → 22 janelas reais (352 pontos)
    '   - 2 janelas por bloco (11 blocos × 32 pontos)
    '   - Distribuição uniforme entre offsets 0-31
    ' 
    ' RETORNA:
    '   True se ponto foi obtido com sucesso, False caso contrário
    ' -------------------------------------------------------------------------
    
    ' -------------------------------------------------------------------------
    ' VALIDAÇÃO: Verificar parâmetros de entrada
    ' -------------------------------------------------------------------------
    If window < 0 Or window > 63 Or index <= 0 Or index > 15 Then
        get_precomputed_point_fixed = False
        Exit Function
    End If
    
    ' -------------------------------------------------------------------------
    ' MAPEAMENTO: Converter coordenadas 2D para estrutura 11×32
    ' -------------------------------------------------------------------------
    ' 64 janelas × 16 valores = 1024 pontos teóricos
    ' 11 blocos × 32 pontos = 352 pontos reais
    ' Usar apenas primeiras 22 janelas (22×16=352)
    
    If window >= 22 Then
        get_precomputed_point_fixed = False
        Exit Function
    End If
    
    Dim block As Long, offset As Long
    block = window \ 2  ' 2 janelas por bloco
    offset = (index - 1) + (window Mod 2) * 16  ' Distribuir entre 0-31
    
    If block > 10 Or offset > 31 Then
        get_precomputed_point_fixed = False
        Exit Function
    End If
    
    ' Obter entrada da tabela
    Dim entry As String
    entry = get_gen_point(block, offset)
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