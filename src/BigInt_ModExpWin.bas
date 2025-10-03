Attribute VB_Name = "BigInt_ModExpWin"
Option Explicit
Option Compare Binary
Option Base 0

' =============================================================================
' BIGINT WINDOWED MODULAR EXPONENTIATION VBA - ALGORITMO SLIDING WINDOW
' =============================================================================
' Implementação de exponenciação modular com janela deslizante (w=4)
' Algoritmo: left-to-right com pré-computação de potências ímpares
' Otimizações: tabela apenas com potências ímpares (8 entradas) e loops eficientes
' Ideal para expoentes densos e grandes (>160 bits com alta densidade)
' =============================================================================

' =============================================================================
' EXPONENCIAÇÃO MODULAR COM JANELA DESLIZANTE DE 4 BITS
' =============================================================================

Public Function BN_mod_exp_win4(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef e As BIGNUM_TYPE, ByRef m As BIGNUM_TYPE) As Boolean
    ' Calcula a^e mod m usando algoritmo de janela deslizante
    ' Parâmetros: r = a^e mod m (resultado), a = base, e = expoente, m = módulo
    ' Algoritmo: pré-computa potências ímpares e processa expoente em janelas de 4 bits
    ' Complexidade: reduz número de multiplicações para expoentes densos
    
    Const w As Long = 4  ' Tamanho da janela (4 bits)
    Dim nbits As Long, i As Long
    Dim acc As BIGNUM_TYPE, base As BIGNUM_TYPE, base2 As BIGNUM_TYPE
    Dim bnOne As BIGNUM_TYPE
    Dim table(0 To 7) As BIGNUM_TYPE   ' Tabela de potências ímpares: a^1, a^3, ..., a^15
    
    ' Validação de entrada: módulo não pode ser zero
    If BN_is_zero(m) Then BN_mod_exp_win4 = False: Exit Function
    
    ' Analisar tamanho do expoente
    nbits = BN_num_bits(e)
    
    ' Caso especial: expoente zero (a^0 = 1 para qualquer a)
    If nbits = 0 Then
        bnOne = BN_new(): Call BN_set_word(bnOne, 1): Call BN_copy(r, bnOne)
        BN_mod_exp_win4 = True: Exit Function
    End If

    ' =============================================================================
    ' PRÉ-COMPUTAÇÃO DE POTÊNCIAS ÍMPARES
    ' =============================================================================

    ' Inicializar variáveis de trabalho
    base = BN_new(): acc = BN_new(): base2 = BN_new()
    
    ' Reduzir base módulo m para eficiência
    If Not BN_mod(base, a, m) Then BN_mod_exp_win4 = False: Exit Function
    
    ' Calcular base^2 mod m (usado para gerar potências ímpares)
    If Not BN_mod_sqr(base2, base, m) Then BN_mod_exp_win4 = False: Exit Function
    
    ' Pré-computar tabela de potências ímpares: a^1, a^3, a^5, ..., a^15
    ' Fórmula: table[0] = a^1, table[i] = table[i-1] * a^2 = a^(2*i+1)
    Dim k As Long
    For k = 0 To 7
        table(k) = BN_new()
    Next k
    
    ' table[0] = a^1
    Call BN_copy(table(0), base)
    
    ' Gerar demais potências ímpares multiplicando por a^2
    For k = 1 To 7
        If Not BN_mod_mul(table(k), table(k - 1), base2, m) Then BN_mod_exp_win4 = False: Exit Function
    Next k

    ' =============================================================================
    ' ALGORITMO SLIDING WINDOW PRINCIPAL
    ' =============================================================================

    ' Inicializar acumulador com 1 (identidade multiplicativa)
    bnOne = BN_new(): Call BN_set_word(bnOne, 1)
    Call BN_copy(acc, bnOne)
    
    ' Processar expoente da esquerda para direita (MSB para LSB)
    i = nbits - 1
    Do While i >= 0
        ' Caso 1: bit atual é 0 - apenas elevar ao quadrado
        If Not BN_is_bit_set(e, i) Then
            If Not BN_mod_sqr(acc, acc, m) Then BN_mod_exp_win4 = False: Exit Function
            i = i - 1
        Else
            ' Caso 2: bit atual é 1 - usar janela deslizante
            ' Algoritmo de janela deslizante: encontrar janela ímpar máxima
            ' Escolher l em [i-3 .. i] tal que bit l = 1 (janela ímpar) e maximizar comprimento
            Dim l As Long, j As Long, s As Long, winval As Long
            ' Definir limite inferior da janela (máximo 4 bits)
            l = i - 3: If l < 0 Then l = 0
            
            ' Mover l para cima até encontrar bit 1 (garantir janela ímpar)
            Do While (l < i) And (Not BN_is_bit_set(e, l))
                l = l + 1
            Loop
            ' Calcular tamanho da janela
            s = i - l + 1
            
            ' Elevar acumulador ao quadrado s vezes (processar janela)
            For j = 1 To s
                If Not BN_mod_sqr(acc, acc, m) Then BN_mod_exp_win4 = False: Exit Function
            Next j
            ' Construir valor da janela a partir dos bits e[i..l]
            winval = 0
            For j = i To l Step -1
                winval = winval * 2
                If BN_is_bit_set(e, j) Then winval = winval + 1
            Next j
            ' Calcular índice na tabela de potências ímpares
            ' Índice = (winval - 1) / 2 (pois winval é ímpar)
            Dim idx As Long: idx = (winval - 1) \ 2
            
            ' Multiplicar pelo valor pré-computado correspondente
            If Not BN_mod_mul(acc, acc, table(idx), m) Then BN_mod_exp_win4 = False: Exit Function
            ' Avançar para próxima posição
            i = l - 1
        End If
    Loop

    ' =============================================================================
    ' FINALIZAÇÃO
    ' =============================================================================

    ' Copiar resultado final do acumulador
    Call BN_copy(r, acc)
    BN_mod_exp_win4 = True
End Function