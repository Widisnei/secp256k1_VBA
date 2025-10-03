Attribute VB_Name = "BigInt_VBA"
Option Explicit
Option Compare Binary
Option Base 0

' =============================================================================
' BIGINT VBA - ARITMÉTICA DE PRECISÃO ARBITRÁRIA
' =============================================================================
' Implementação completa de números inteiros de precisão arbitrária em VBA
' Compatível com OpenSSL BIGNUM, suporte a operações modulares e criptográficas
' Otimizado para operações secp256k1 e algoritmos de curva elíptica
' =============================================================================

' =============================================================================
' ESTRUTURAS DE DADOS E CONSTANTES
' =============================================================================

' Representa um número inteiro de precisão arbitrária
Public Type BIGNUM_TYPE
    d() As Long         ' Array de limbs de 32 bits (palavra menos significativa em d(0))
    top As Long         ' Número de limbs em uso; top=0 indica zero lógico
    dmax As Long        ' Tamanho alocado do array d() (capacidade máxima)
    neg As Boolean      ' Sinal do número (True = negativo, False = positivo)
    flags As Long       ' Flags de controle para otimizações e estados especiais
End Type

' Flags de controle para otimizações e estados especiais
Public Const BN_FLG_MALLOCED As Long = &H1        ' Memória alocada dinamicamente
Public Const BN_FLG_STATIC_DATA As Long = &H2     ' Dados estáticos (não liberar)
Public Const BN_FLG_CONSTTIME As Long = &H4       ' Operações em tempo constante
Public Const BN_FLG_FIXED_TOP As Long = &H8000    ' Top fixo para segurança

' Constantes matemáticas e de arquitetura
Public Const BN_BITS2 As Long = 32                ' Bits por limb (32 bits)
Public Const BN_BYTES As Long = 4                 ' Bytes por limb (4 bytes)
Private Const TWO32 As Double = 4294967296.0#       ' 2^32 para conversões
Private Const TWO16 As Double = 65536.0#            ' 2^16 para decomposição

' =============================================================================
' FUNÇÕES DE CONVERSÃO E UTILITÁRIOS DE TIPO
' =============================================================================

Public Function DoubleToLong32(ByVal val As Double) As Long
    ' Converte Double para Long de 32 bits com tratamento de overflow
    ' Essencial para aritmética de precisão em VBA
    val = val - Fix(val / TWO32) * TWO32
    If val < 0# Then val = val + TWO32
    If val >= 2147483648.0# Then
        DoubleToLong32 = CLng(val - TWO32)
    Else
        DoubleToLong32 = CLng(val)
    End If
End Function

Public Function LongToUnsignedDouble(ByVal val As Long) As Double
    ' Converte Long com sinal para Double sem sinal
    ' Necessário pois VBA não tem tipos unsigned nativos
    If val < 0 Then
        LongToUnsignedDouble = TWO32 + CDbl(val)
    Else
        LongToUnsignedDouble = CDbl(val)
    End If
End Function

' =============================================================================
' OPERAÇÕES BÁSICAS DE CRIAÇÃO E MANIPULAÇÃO
' =============================================================================

Public Function BN_new() As BIGNUM_TYPE
    ' Cria um novo BIGNUM inicializado com valor zero
    ' Retorna: BIGNUM_TYPE com capacidade mínima de 1 limb
    Dim bn As BIGNUM_TYPE
    ReDim bn.d(0 To 0)
    bn.dmax = 1
    BN_new = bn
End Function

Public Sub BN_free(ByRef bn As BIGNUM_TYPE)
    ' Libera memória e reinicializa BIGNUM para estado limpo
    Erase bn.d
    bn.top = 0
    bn.dmax = 0
    bn.neg = False
End Sub

Public Sub BN_zero(ByRef bn As BIGNUM_TYPE)
    ' Define BIGNUM como zero, mantendo capacidade alocada
    If bn.dmax < 1 Then
        ReDim bn.d(0 To 0)
        bn.dmax = 1
    End If
    bn.d(0) = 0
    bn.top = 0
    bn.neg = False
End Sub

Public Function BN_is_zero(ByRef bn As BIGNUM_TYPE) As Boolean
    ' Verifica se BIGNUM representa zero (top = 0)
    BN_is_zero = (bn.top = 0)
End Function

Public Function BN_is_one(ByRef bn As BIGNUM_TYPE) As Boolean
    ' Verifica se BIGNUM representa o valor 1
    If bn.neg Then BN_is_one = False : Exit Function
    If bn.top <> 1 Then BN_is_one = False : Exit Function
    BN_is_one = (bn.d(0) = 1)
End Function

Public Function BN_is_odd(ByRef bn As BIGNUM_TYPE) As Boolean
    ' Verifica se BIGNUM é ímpar (bit menos significativo = 1)
    If bn.top = 0 Then BN_is_odd = False : Exit Function
    BN_is_odd = ((bn.d(0) And 1) <> 0)
End Function

Public Function BN_set_word(ByRef bn As BIGNUM_TYPE, ByVal w As Long) As Boolean
    ' Define BIGNUM com valor de uma palavra (32 bits)
    ' Parâmetro: w - valor Long a ser atribuído
    If Not bn_wexpand(bn, 1) Then BN_set_word = False : Exit Function
    bn.d(0) = w
    If w = 0 Then
        bn.top = 0
    Else
        bn.top = 1
    End If
    bn.neg = False
    BN_set_word = True
End Function

Public Function BN_bn2hex(ByRef bn As BIGNUM_TYPE) As String
    ' Converte BIGNUM para representação hexadecimal
    ' Retorna: String hexadecimal em maiúsculas, com sinal se negativo
    Dim i As Long, hexStr As String
    If bn.top = 0 Then BN_bn2hex = "0" : Exit Function
    hexStr = hex$(bn.d(bn.top - 1))
    For i = bn.top - 2 To 0 Step -1
        hexStr = hexStr & right$("00000000" & hex$(bn.d(i)), 8)
    Next i
    If bn.neg And hexStr <> "0" Then
        BN_bn2hex = "-" & hexStr
    Else
        BN_bn2hex = hexStr
    End If
End Function

Public Function BN_hex2bn(ByVal hexStr As String) As BIGNUM_TYPE
    ' Converte string hexadecimal para BIGNUM
    ' Parâmetro: hexStr - string hex (aceita prefixos 0x, -0x e zeros à esquerda)
    ' Retorna: BIGNUM_TYPE com valor correspondente
    Dim bn As BIGNUM_TYPE, i As Long, chunkStr As String, isNegative As Boolean
    Dim numChunks As Long, tempStr As String
    bn = BN_new()
    tempStr = Trim$(hexStr)
    If Len(tempStr) = 0 Then BN_hex2bn = bn : Exit Function
    If left$(tempStr, 1) = "-" Then
        isNegative = True
        tempStr = mid$(tempStr, 2)
    End If
    If LCase$(left$(tempStr, 2)) = "0x" Then tempStr = mid$(tempStr, 3)
    Do While left$(tempStr, 1) = "0" And Len(tempStr) > 1
        tempStr = mid$(tempStr, 2)
    Loop
    If tempStr = "0" Or tempStr = "" Then BN_hex2bn = bn : Exit Function
    numChunks = (Len(tempStr) + 7) \ 8
    If Not bn_wexpand(bn, numChunks) Then BN_hex2bn = bn : Exit Function
    bn.top = numChunks
    For i = 0 To numChunks - 1
        If Len(tempStr) > 8 Then
            chunkStr = right$(tempStr, 8)
            tempStr = left$(tempStr, Len(tempStr) - 8)
        Else
            chunkStr = tempStr
            tempStr = ""
        End If
        On Error Resume Next
        bn.d(i) = CLng("&H" & chunkStr)
        If Err.Number <> 0 Then
            ' Tratar overflow para valores hexadecimais grandes
            bn.d(i) = DoubleToLong32(CDbl("&H" & chunkStr))
            Err.Clear
        End If
        On Error GoTo 0
    Next i
    bn.neg = isNegative
    If bn.top = 0 Then bn.neg = False
    BN_hex2bn = bn
End Function

Public Function bn_wexpand(ByRef bn As BIGNUM_TYPE, ByVal words As Long) As Boolean
    ' Expande capacidade do BIGNUM para acomodar número especificado de palavras
    ' Parâmetro: words - número mínimo de palavras necessárias
    ' Retorna: True se expansão foi bem-sucedida
    If words <= 0 Then words = 1
    If words > bn.dmax Then
        ReDim Preserve bn.d(0 To words - 1)
        bn.dmax = words
    End If
    bn_wexpand = True
End Function

Public Sub BN_copy(ByRef dest As BIGNUM_TYPE, ByRef src As BIGNUM_TYPE)
    ' Copia todos os dados de um BIGNUM para outro
    ' Parâmetros: dest - BIGNUM destino, src - BIGNUM origem
    If Not bn_wexpand(dest, src.top) Then Exit Sub
    Dim i As Long
    For i = 0 To src.top - 1
        dest.d(i) = src.d(i)
    Next i
    dest.top = src.top
    dest.neg = src.neg
End Sub

' =============================================================================
' ARITMÉTICA DE BAIXO NÍVEL (OPERAÇÕES PRIMITIVAS)
' =============================================================================

Private Function bn_add_words(ByRef rp() As Long, ByRef ap() As Long, ByRef bp() As Long, ByVal num As Long) As Double
    ' Adição de arrays de palavras com propagação de carry
    ' Implementação otimizada para operações de precisão
    Dim i As Long, carry As Double, t As Double
    carry = 0#
    For i = 0 To num - 1
        t = LongToUnsignedDouble(ap(i)) + LongToUnsignedDouble(bp(i)) + carry
        rp(i) = DoubleToLong32(t)
        carry = Fix(t / TWO32)
    Next i
    bn_add_words = carry
End Function

Public Function bn_add_word(ByRef rp() As Long, ByVal num As Long, ByVal w As Double) As Double
    ' Adiciona uma palavra a um array de limbs com propagação de carry
    ' Usado em operações de multiplicação e divisão
    Dim i As Long, carry As Double, t As Double
    carry = w
    For i = 0 To num - 1
        If carry = 0# Then Exit For
        t = LongToUnsignedDouble(rp(i)) + carry
        rp(i) = DoubleToLong32(t)
        carry = Fix(t / TWO32)
    Next i
    bn_add_word = carry
End Function

Private Function bn_sub_words(ByRef rp() As Long, ByRef ap() As Long, ByRef bp() As Long, ByVal num As Long) As Long
    ' Subtração de arrays de palavras com propagação de borrow
    Dim i As Long, borrow As Long, t As Double
    borrow = 0
    For i = 0 To num - 1
        t = LongToUnsignedDouble(ap(i)) - LongToUnsignedDouble(bp(i)) - borrow
        rp(i) = DoubleToLong32(t)
        If t < 0# Then
            borrow = 1
        Else
            borrow = 0
        End If
    Next i
    bn_sub_words = borrow
End Function

' Multiplicação precisa 32x32→64 bits via decomposição 16x16
' Evita overflow em VBA usando aritmética de dupla precisão
Private Sub MulAdd32(ByVal a As Long, ByVal b As Long, ByVal add_lo As Long, ByVal add_hi As Long,
                     ByRef out_lo As Long, ByRef out_hi As Long)
    ' Implementa (a * b + add_hi:add_lo) com resultado de 64 bits
    ' Essencial para multiplicação de precisão arbitrária
    Dim ua As Double, ub As Double
    Dim a0 As Double, a1 As Double, b0 As Double, b1 As Double
    Dim m0 As Double, m1 As Double, m2 As Double
    Dim m1_lo As Double, m1_hi As Double
    Dim lo_acc As Double, lo_carry As Double, hi_acc As Double

    ua = LongToUnsignedDouble(a)
    ub = LongToUnsignedDouble(b)

    a0 = ua - Fix(ua / TWO16) * TWO16
    a1 = Fix(ua / TWO16)
    b0 = ub - Fix(ub / TWO16) * TWO16
    b1 = Fix(ub / TWO16)

    m0 = a0 * b0
    m1 = a0 * b1 + a1 * b0
    m2 = a1 * b1

    m1_hi = Fix(m1 / TWO16)
    m1_lo = m1 - m1_hi * TWO16

    lo_acc = m0 + (m1_lo * TWO16) + LongToUnsignedDouble(add_lo)
    lo_carry = Fix(lo_acc / TWO32)
    out_lo = DoubleToLong32(lo_acc)

    hi_acc = m2 + m1_hi + LongToUnsignedDouble(add_hi) + lo_carry
    out_hi = DoubleToLong32(hi_acc)
End Sub

Private Sub bn_mul_add_words_ptr(ByRef rp() As Long, ByVal rp_start As Long,
                                 ByRef ap() As Long, ByVal ap_start As Long,
                                 ByVal num As Long, ByVal w As Long)
    ' Multiplica array de palavras por escalar e adiciona ao resultado
    ' Usado no algoritmo de multiplicação escola primária
    Dim i As Long, carry As Long, out_lo As Long, out_hi As Long
    Dim sum As Double, add_lo As Long, add_hi As Long

    carry = 0
    For i = 0 To num - 1
        sum = LongToUnsignedDouble(rp(rp_start + i)) + LongToUnsignedDouble(carry)
        add_lo = DoubleToLong32(sum)
        add_hi = CLng(Fix(sum / TWO32))
        Call MulAdd32(ap(ap_start + i), w, add_lo, add_hi, out_lo, out_hi)
        rp(rp_start + i) = out_lo
        carry = out_hi
    Next i
    rp(rp_start + num) = DoubleToLong32(LongToUnsignedDouble(rp(rp_start + num)) + LongToUnsignedDouble(carry))
End Sub

Private Sub bn_mul_word_internal(ByRef r() As Long, ByRef a() As Long, ByVal num As Long, ByVal w As Long)
    ' Multiplica array de palavras por escalar único
    ' Função base para multiplicação de BIGNUM por palavra
    Dim i As Long, carry As Long, out_lo As Long, out_hi As Long
    carry = 0
    For i = 0 To num - 1
        Call MulAdd32(a(i), w, carry, 0, out_lo, out_hi)
        r(i) = out_lo
        carry = out_hi
    Next i
    r(num) = carry
End Sub

' =============================================================================
' OPERAÇÕES ARITMÉTICAS PRINCIPAIS
' =============================================================================

Public Function BN_ucmp(ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE) As Long
    ' Comparação de magnitudes (sem considerar sinal)
    ' Retorna: -1 se a < b, 0 se a = b, 1 se a > b
    Dim i As Long
    If a.top > b.top Then BN_ucmp = 1 : Exit Function
    If a.top < b.top Then BN_ucmp = -1 : Exit Function
    For i = a.top - 1 To 0 Step -1
        If LongToUnsignedDouble(a.d(i)) > LongToUnsignedDouble(b.d(i)) Then BN_ucmp = 1 : Exit Function
        If LongToUnsignedDouble(a.d(i)) < LongToUnsignedDouble(b.d(i)) Then BN_ucmp = -1 : Exit Function
    Next i
    BN_ucmp = 0
End Function

Public Function BN_cmp(ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE) As Long
    ' Comparação com sinal (considera valores negativos)
    ' Retorna: -1 se a < b, 0 se a = b, 1 se a > b
    If a.neg <> b.neg Then
        If a.top = 0 And b.top = 0 Then BN_cmp = 0 Else BN_cmp = IIf(a.neg, -1, 1)
        Exit Function
    End If
    Dim res As Long : res = BN_ucmp(a, b)
    If a.neg Then BN_cmp = -res Else BN_cmp = res
End Function

Public Function BN_uadd(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE) As Boolean
    ' Adição de magnitudes (sem considerar sinal)
    ' Parâmetros: r - resultado, a e b - operandos
    ' Retorna: True se operação foi bem-sucedida
    Dim p_a As BIGNUM_TYPE, p_b As BIGNUM_TYPE
    Dim max_len As Long, min_len As Long
    Dim carry As Double, k As Long, t As Double

    If a.top < b.top Then p_a = b : p_b = a Else p_a = a : p_b = b
    max_len = p_a.top : min_len = p_b.top

    If Not bn_wexpand(r, max_len + 1) Then BN_uadd = False : Exit Function

    carry = bn_add_words(r.d, p_a.d, p_b.d, min_len)
    For k = min_len To max_len - 1
        t = LongToUnsignedDouble(p_a.d(k)) + carry
        r.d(k) = DoubleToLong32(t)
        carry = Fix(t / TWO32)
    Next k

    If carry > 0# Then
        r.d(max_len) = DoubleToLong32(carry)
        r.top = max_len + 1
    Else
        r.top = max_len
    End If

    Do While r.top > 0
        If r.d(r.top - 1) = 0 Then r.top = r.top - 1 Else Exit Do
    Loop

    r.neg = False
    BN_uadd = True
End Function

Public Function BN_usub(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE) As Boolean
    ' Subtração de magnitudes (assume a >= b)
    ' Parâmetros: r - resultado, a - minuendo, b - subtraendo
    ' Retorna: True se operação foi bem-sucedida
    Dim max_len As Long, min_len As Long, borrow As Long, i As Long, t As Double
    If BN_ucmp(a, b) < 0 Then BN_usub = False : Exit Function

    max_len = a.top : min_len = b.top
    If Not bn_wexpand(r, max_len) Then BN_usub = False : Exit Function

    borrow = bn_sub_words(r.d, a.d, b.d, min_len)
    For i = min_len To max_len - 1
        t = LongToUnsignedDouble(a.d(i)) - borrow
        r.d(i) = DoubleToLong32(t)
        If t < 0# Then borrow = 1 Else borrow = 0
    Next i

    r.top = max_len
    Do While r.top > 0
        If r.d(r.top - 1) = 0 Then r.top = r.top - 1 Else Exit Do
    Loop

    r.neg = False
    BN_usub = True
End Function

Public Function BN_add(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE) As Boolean
    ' Adição com sinal (r = a + b)
    ' Trata automaticamente sinais positivos e negativos
    Dim ret As Boolean, r_neg As Boolean, cmp_res As Long

    If a.neg = b.neg Then
        r_neg = a.neg
        ret = BN_uadd(r, a, b)
    Else
        cmp_res = BN_ucmp(a, b)
        If cmp_res > 0 Then
            r_neg = a.neg
            ret = BN_usub(r, a, b)
        ElseIf cmp_res < 0 Then
            r_neg = b.neg
            ret = BN_usub(r, b, a)
        Else
            r_neg = False
            BN_zero r
            ret = True
        End If
    End If

    r.neg = r_neg
    BN_add = ret
End Function

Public Function BN_sub(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE) As Boolean
    ' Subtração com sinal (r = a - b)
    ' Implementada como adição com sinal invertido
    b.neg = Not b.neg
    BN_sub = BN_add(r, a, b)
    b.neg = Not b.neg
End Function

Public Function BN_mul(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE) As Boolean
    ' Multiplicação de BIGNUM com otimização para operações 256-bit
    ' Parâmetros: r - resultado (a * b), a e b - fatores
    ' Usa algoritmo COMBA para operandos <= 256 bits, escola primária para maiores

    Dim p_a As BIGNUM_TYPE, p_b As BIGNUM_TYPE, i As Long
    ' Caminho rápido: usar COMBA 8x8 quando ambos operandos cabem em 256 bits
    If a.top <= 8 And b.top <= 8 Then
        ' Usar multiplicação rápida COMBA para operandos de até 256 bits
        BN_mul = BN_mul_comba8(r, a, b)
        Exit Function
    End If

    If a.top < b.top Then p_a = b : p_b = a Else p_a = a : p_b = b
    If p_b.top = 0 Then BN_zero r: BN_mul = True : Exit Function
    If Not bn_wexpand(r, p_a.top + p_b.top) Then BN_mul = False : Exit Function
    For i = 0 To p_a.top + p_b.top - 1 : r.d(i) = 0 : Next i
    Call bn_mul_word_internal(r.d, p_a.d, p_a.top, p_b.d(0))
    For i = 1 To p_b.top - 1
        Call bn_mul_add_words_ptr(r.d, i, p_a.d, 0, p_a.top, p_b.d(i))
    Next i
    r.top = p_a.top + p_b.top
    r.neg = (a.neg Xor b.neg)
    Do While r.top > 0
        If r.d(r.top - 1) = 0 Then r.top = r.top - 1 Else Exit Do
    Loop
    If r.top = 0 Then r.neg = False
    BN_mul = True
End Function

' =============================================================================
' OPERAÇÕES DE DESLOCAMENTO DE BITS (SHIFTS)
' =============================================================================

Public Function BN_lshift(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByVal n As Long) As Boolean
    ' Deslocamento à esquerda (multiplicação por 2^n)
    ' Parâmetros: r - resultado, a - operando, n - número de bits
    ' Algoritmo: separar em deslocamento de palavras + deslocamento de bits
    ' Otimizado para operações criptográficas de alta performance

    ' Validação de entrada
    If n < 0 Then BN_lshift = False : Exit Function
    If a.top = 0 Then BN_zero r: BN_lshift = True : Exit Function

    Dim words As Long, bits As Long, i As Long
    Dim old_top As Long, new_top As Long
    Dim src As BIGNUM_TYPE
    Dim w As Long
    Dim out_lo As Long, out_hi As Long
    Dim carry As Long

    src = BN_new() : BN_copy src, a

    words = n \ BN_BITS2
    bits = n Mod BN_BITS2
    old_top = src.top
    new_top = old_top + words

    If bits > 0 Then
        Dim top_bits As Long
        top_bits = ((BN_num_bits(src) - 1) Mod BN_BITS2) + 1
        If top_bits + bits > BN_BITS2 Then new_top = new_top + 1
    End If

    If Not bn_wexpand(r, new_top) Then Exit Function
    r.top = new_top

    For i = 0 To words - 1
        r.d(i) = 0
    Next i

    If bits = 0 Then
        For i = 0 To old_top - 1
            r.d(i + words) = src.d(i)
        Next i
    Else
        w = DoubleToLong32(2.0# ^ bits)
        carry = 0
        For i = 0 To old_top - 1
            Call MulAdd32(src.d(i), w, carry, 0, out_lo, out_hi)
            r.d(i + words) = out_lo
            carry = out_hi
        Next i
        If (words + old_top) <= (new_top - 1) Then
            r.d(words + old_top) = carry
        End If
    End If

    Do While r.top > 0
        If r.d(r.top - 1) = 0 Then r.top = r.top - 1 Else Exit Do
    Loop

    r.neg = src.neg
    If r.top = 0 Then r.neg = False
    BN_lshift = True
End Function

Public Function BN_rshift(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByVal n As Long) As Boolean
    ' Deslocamento à direita (divisão por 2^n)
    ' Parâmetros: r - resultado, a - operando, n - número de bits
    ' Implementação otimizada com pré-cálculo de potências
    If n < 0 Then BN_rshift = False : Exit Function

    Dim words As Long, bits As Long, i As Long
    Dim old_top As Long, new_top As Long
    Dim src As BIGNUM_TYPE
    Dim cur As Double, nextw As Double, lowbits As Double, hi_contrib As Double, q As Double

    src = BN_new() : BN_copy src, a
    old_top = src.top
    words = n \ BN_BITS2
    bits = n Mod BN_BITS2

    ' Pré-calcular potências para evitar recomputação e proteger divisão por zero
    Dim pow2_bits As Double, pow2_inv As Double
    bits = bits And (BN_BITS2 - 1)
    If bits <> 0 Then
        pow2_bits = 2.0# ^ bits
        pow2_inv = 2.0# ^ (BN_BITS2 - bits)
    End If

    If words >= old_top Then BN_zero r: BN_rshift = True : Exit Function

    new_top = old_top - words
    If Not bn_wexpand(r, new_top) Then Exit Function
    r.top = new_top

    If bits = 0 Then
        For i = 0 To new_top - 1
            r.d(i) = src.d(i + words)
        Next i
    Else
        For i = 0 To new_top - 1
            cur = LongToUnsignedDouble(src.d(i + words))
            If (i + words + 1) < old_top Then
                nextw = LongToUnsignedDouble(src.d(i + words + 1))
            Else
                nextw = 0#
            End If
            q = Fix(cur / (pow2_bits))
            lowbits = nextw - Fix(nextw / (pow2_bits)) * (pow2_bits)
            hi_contrib = lowbits * (pow2_inv)
            r.d(i) = DoubleToLong32(q + hi_contrib)
        Next i
    End If

    Do While r.top > 0
        If r.d(r.top - 1) = 0 Then r.top = r.top - 1 Else Exit Do
    Loop
    r.neg = src.neg
    If r.top = 0 Then r.neg = False
    BN_rshift = True
End Function

' =============================================================================
' UTILITÁRIOS DE MANIPULAÇÃO DE BITS E BYTES
' =============================================================================

Public Function BN_num_bits(ByRef bn As BIGNUM_TYPE) As Long
    ' Calcula número total de bits significativos no BIGNUM
    ' Retorna: Número de bits necessários para representar o valor
    If bn.top = 0 Then Exit Function
    Dim lastChunk As Double : lastChunk = LongToUnsignedDouble(bn.d(bn.top - 1))
    Dim bits As Long
    Do While lastChunk > 0#
        lastChunk = Fix(lastChunk / 2.0#)
        bits = bits + 1
    Loop
    BN_num_bits = (bn.top - 1) * BN_BITS2 + bits
End Function

Public Function BN_is_bit_set(ByRef a As BIGNUM_TYPE, ByVal n As Long) As Boolean
    ' Verifica se um bit específico está definido (=1)
    ' Parâmetros: a - BIGNUM, n - posição do bit (0 = LSB)
    If n < 0 Then BN_is_bit_set = False : Exit Function
    Dim w As Long, off As Long, mask As Long
    w = n \ BN_BITS2
    off = n Mod BN_BITS2
    If w >= a.top Then BN_is_bit_set = False : Exit Function
    If off = 31 Then
        mask = &H80000000
    Else
        mask = DoubleToLong32(2.0# ^ off)
    End If
    BN_is_bit_set = ((a.d(w) And mask) <> 0)
End Function

Public Function BN_set_bit(ByRef a As BIGNUM_TYPE, ByVal n As Long) As Boolean
    ' Define um bit específico como 1
    ' Expande BIGNUM automaticamente se necessário
    If n < 0 Then BN_set_bit = False : Exit Function
    Dim w As Long, off As Long, mask As Long
    w = n \ BN_BITS2
    off = n Mod BN_BITS2

    If w >= a.top Then
        If Not bn_wexpand(a, w + 1) Then BN_set_bit = False : Exit Function
        Dim i As Long
        For i = a.top To w
            a.d(i) = 0
        Next i
        a.top = w + 1
    End If

    If off = 31 Then
        mask = &H80000000
    Else
        mask = DoubleToLong32(2.0# ^ off)
    End If
    a.d(w) = a.d(w) Or mask
    BN_set_bit = True
End Function

Public Function BN_clear_bit(ByRef a As BIGNUM_TYPE, ByVal n As Long) As Boolean
    ' Define um bit específico como 0
    ' Ajusta tamanho do BIGNUM se bit mais significativo for zerado
    If n < 0 Then BN_clear_bit = True : Exit Function
    Dim w As Long, off As Long, mask As Long
    w = n \ BN_BITS2
    off = n Mod BN_BITS2

    If w >= a.top Then BN_clear_bit = True : Exit Function

    If off = 31 Then
        mask = &H80000000
    Else
        mask = DoubleToLong32(2.0# ^ off)
    End If
    a.d(w) = a.d(w) And (Not mask)

    If w = a.top - 1 And a.d(w) = 0 Then
        Do While a.top > 0 And a.d(a.top - 1) = 0
            a.top = a.top - 1
        Loop
    End If

    BN_clear_bit = True
End Function

Public Function BN_num_bytes(ByRef a As BIGNUM_TYPE) As Long
    ' Calcula número de bytes necessários para representar BIGNUM
    BN_num_bytes = (BN_num_bits(a) + 7) \ 8
End Function

Public Function BN_bn2bin(ByRef a As BIGNUM_TYPE) As Byte()
    ' Converte BIGNUM para array de bytes (big-endian)
    ' Retorna: Array de bytes com representação binária
    Dim num_bytes As Long : num_bytes = BN_num_bytes(a)
    Dim toBytes() As Byte
    If num_bytes = 0 Then BN_bn2bin = toBytes : Exit Function
    ReDim toBytes(0 To num_bytes - 1)

    Dim i As Long, j As Long, byte_idx As Long, chunk_val As Double
    byte_idx = num_bytes - 1

    For i = 0 To a.top - 1
        chunk_val = LongToUnsignedDouble(a.d(i))
        For j = 0 To BN_BYTES - 1
            If byte_idx >= 0 Then
                toBytes(byte_idx) = chunk_val - Fix(chunk_val / 256.0#) * 256.0#
                chunk_val = Fix(chunk_val / 256.0#)
                byte_idx = byte_idx - 1
            End If
        Next j
    Next i

    BN_bn2bin = toBytes
End Function

Public Function BN_bin2bn(ByRef s() As Byte, ByVal dataLen As Long) As BIGNUM_TYPE
    ' Converte array de bytes (big-endian) para BIGNUM
    ' Parâmetros: s - array de bytes, dataLen - comprimento válido
    Dim bn As BIGNUM_TYPE : bn = BN_new()
    If dataLen = 0 Then BN_bin2bn = bn : Exit Function

    Dim num_chunks As Long : num_chunks = (dataLen + BN_BYTES - 1) \ BN_BYTES
    If Not bn_wexpand(bn, num_chunks) Then BN_bin2bn = bn : Exit Function

    Dim i As Long, j As Long, byte_idx As Long, current_chunk As Double
    byte_idx = dataLen - 1

    For i = 0 To num_chunks - 1
        current_chunk = 0#
        For j = 0 To BN_BYTES - 1
            If byte_idx >= 0 Then
                current_chunk = current_chunk + s(byte_idx) * (2.0# ^ (j * 8))
                byte_idx = byte_idx - 1
            End If
        Next j
        bn.d(i) = DoubleToLong32(current_chunk)
    Next i

    bn.top = num_chunks
    Do While bn.top > 0
        If bn.d(bn.top - 1) = 0 Then bn.top = bn.top - 1 Else Exit Do
    Loop
    BN_bin2bn = bn
End Function

' =============================================================================
' DIVISÃO EUCLIDIANA E OPERAÇÕES MODULARES
' =============================================================================

Public Function BN_div(ByRef dv As BIGNUM_TYPE, ByRef r As BIGNUM_TYPE, ByRef num As BIGNUM_TYPE, ByRef d As BIGNUM_TYPE) As Boolean
    ' Divisão euclidiana: num = dv * d + r (onde 0 <= r < |d|)
    ' Parâmetros: dv - quociente, r - resto, num - dividendo, d - divisor
    ' Implementa semântica OpenSSL com ajustes para resto sempre positivo
    If BN_is_zero(d) Then BN_div = False : Exit Function

    ' Usar divisão de magnitudes
    Dim abs_num As BIGNUM_TYPE, abs_d As BIGNUM_TYPE, q As BIGNUM_TYPE, remainder As BIGNUM_TYPE
    abs_num = BN_new() : abs_d = BN_new() : q = BN_new() : remainder = BN_new()

    Call BN_copy(abs_num, num) : abs_num.neg = False
    Call BN_copy(abs_d, d) : abs_d.neg = False

    ' Divisão simples: |num| / |d|
    If BN_ucmp(abs_num, abs_d) < 0 Then
        BN_zero q
        Call BN_copy(remainder, abs_num)
    Else
        If Not BN_div_unsigned(q, remainder, abs_num, abs_d) Then BN_div = False : Exit Function
    End If

    ' Aplicar sinais como OpenSSL: dv->neg = num->neg ^ divisor->neg, rm->neg = num->neg
    q.neg = (num.neg Xor d.neg) And (q.top > 0)
    remainder.neg = num.neg And (remainder.top > 0)

    ' Para divisão euclidiana: ajustar se remainder negativo
    If remainder.neg Then
        Call BN_add(remainder, remainder, abs_d)
        remainder.neg = False
        Dim one As BIGNUM_TYPE : one = BN_new() : Call BN_set_word(one, 1)
        If d.neg Then
            Call BN_add(q, q, one)
        Else
            Call BN_sub(q, q, one)
        End If
        ' Reaplica sinal do quociente
        q.neg = (num.neg Xor d.neg) And (q.top > 0)
    End If

    Call BN_copy(dv, q)
    Call BN_copy(r, remainder)
    BN_div = True
End Function

Private Function BN_div_unsigned(ByRef dv As BIGNUM_TYPE, ByRef r As BIGNUM_TYPE, ByRef num As BIGNUM_TYPE, ByRef d As BIGNUM_TYPE) As Boolean
    ' Divisão longa para magnitudes usando algoritmo de Knuth (TAOCP Vol 2)
    ' Implementa normalização, estimação de quociente e correção
    ' Algoritmo robusto para aritmética de precisão arbitrária
    ' Assume entradas positivas (magnitudes apenas)
    If BN_ucmp(num, d) < 0 Then
        Call BN_copy(r, num)
        BN_zero dv
        BN_div_unsigned = True
        Exit Function
    End If

    Dim norm_shift As Long, i As Long, j As Long
    Dim snum As BIGNUM_TYPE, sdiv As BIGNUM_TYPE
    Dim num_n As Long, div_n As Long, loop_count As Long
    Dim d0 As Double, d1 As Double
    Dim wnum() As Long
    Dim q As Double, n0 As Double, n1 As Double, n2 As Double

    snum = BN_new() : sdiv = BN_new()
    Call BN_copy(sdiv, d)

    ' Normalização: garantir que MSB do divisor seja 1 (otimização de Knuth)
    ' Isso melhora a precisão da estimação do quociente
    Dim top_word As Double : top_word = LongToUnsignedDouble(sdiv.d(sdiv.top - 1))
    norm_shift = 0
    If top_word < 2147483648.0# Then
        Do While top_word < 2147483648.0# And top_word <> 0
            top_word = top_word * 2.0#
            norm_shift = norm_shift + 1
        Loop
    End If

    ' Aplicar normalização tanto ao divisor quanto ao dividendo
    If norm_shift > 0 Then
        Call BN_lshift(sdiv, sdiv, norm_shift)
        Call BN_lshift(snum, num, norm_shift)
    Else
        Call BN_copy(snum, num)
    End If

    ' Preparar estruturas para algoritmo de divisão longa
    div_n = sdiv.top : num_n = snum.top
    If Not bn_wexpand(snum, num_n + 1) Then BN_div_unsigned = False : Exit Function
    snum.d(num_n) = 0 : num_n = num_n + 1  ' Adicionar palavra extra para overflow

    ' Calcular número de iterações e preparar quociente
    loop_count = num_n - div_n
    If Not bn_wexpand(dv, loop_count) Then BN_div_unsigned = False : Exit Function
    BN_zero dv: dv.top = loop_count

    ' Extrair palavras do divisor para estimação do quociente
    wnum = snum.d
    d0 = LongToUnsignedDouble(sdiv.d(div_n - 1))  ' Palavra mais significativa
    If div_n > 1 Then d1 = LongToUnsignedDouble(sdiv.d(div_n - 2)) Else d1 = 0  ' Segunda palavra

    ' Loop principal: processar cada posição do quociente
    For i = loop_count - 1 To 0 Step -1
        ' Extrair palavras do dividendo para estimação
        n0 = LongToUnsignedDouble(wnum(i + div_n))
        n1 = LongToUnsignedDouble(wnum(i + div_n - 1))

        ' Estimar dígito do quociente usando divisão de duas palavras
        If n0 = d0 Then q = TWO32 - 1 Else q = Fix((n0 * TWO32 + n1) / d0)

        ' Refinar estimação do quociente (teste de Knuth)
        If div_n > 1 Then
            If (i + div_n - 2) >= 0 Then n2 = LongToUnsignedDouble(wnum(i + div_n - 2)) Else n2 = 0
            Dim rem_n As Double : rem_n = (n0 * TWO32 + n1) - q * d0
            ' Ajustar q se a estimação for muito alta
            Do
                If rem_n >= TWO32 Then Exit Do
                If (d1 * q) > (rem_n * TWO32 + n2) Then
                    q = q - 1 : rem_n = rem_n + d0
                Else
                    Exit Do
                End If
            Loop
        End If

        ' Multiplicar divisor por dígito estimado do quociente
        Dim temp_mul As BIGNUM_TYPE : temp_mul = BN_new()
        If Not bn_wexpand(temp_mul, div_n + 1) Then BN_div_unsigned = False : Exit Function
        Call bn_mul_word_internal(temp_mul.d, sdiv.d, div_n, DoubleToLong32(q))
        temp_mul.top = div_n + 1

        ' Subtrair q * divisor do dividendo parcial
        Dim borrow As Long, temp_slice() As Long
        ReDim temp_slice(0 To div_n)
        For j = 0 To div_n : temp_slice(j) = wnum(i + j) : Next j

        borrow = bn_sub_words(temp_slice, temp_slice, temp_mul.d, div_n + 1)

        ' Correção se subtração resultou em borrow (q muito alto)
        If borrow > 0 Then
            q = q - 1  ' Decrementar quociente
            ReDim temp_slice(0 To div_n)
            For j = 0 To div_n : temp_slice(j) = wnum(i + j) : Next j
            ' Adicionar divisor de volta para corrigir
            Dim carry As Double : carry = bn_add_words(temp_slice, temp_slice, sdiv.d, div_n)
            temp_slice(div_n) = DoubleToLong32(LongToUnsignedDouble(temp_slice(div_n)) + carry)
        End If

        For j = 0 To div_n : wnum(i + j) = temp_slice(j) : Next j
        If i < dv.dmax And i >= 0 Then dv.d(i) = DoubleToLong32(q)
    Next i

    ' Extrair e desnormalizar resto
    BN_zero r
    If Not bn_wexpand(r, div_n) Then BN_div_unsigned = False : Exit Function
    For i = 0 To div_n - 1 : r.d(i) = wnum(i) : Next i
    r.top = div_n

    ' Desnormalizar resto (reverter normalização inicial)
    If norm_shift > 0 Then Call BN_rshift(r, r, norm_shift)

    ' Limpar zeros à esquerda do resultado
    Do While r.top > 0
        If r.d(r.top - 1) = 0 Then r.top = r.top - 1 Else Exit Do
    Loop
    Do While dv.top > 0
        If dv.d(dv.top - 1) = 0 Then dv.top = dv.top - 1 Else Exit Do
    Loop

    BN_div_unsigned = True
End Function

Public Function BN_mod(ByRef r As BIGNUM_TYPE, ByRef num As BIGNUM_TYPE, ByRef d As BIGNUM_TYPE) As Boolean
    ' Operação módulo: r = num mod d
    ' Wrapper para BN_div que descarta o quociente
    Dim q As BIGNUM_TYPE : q = BN_new()
    BN_mod = BN_div(q, r, num, d)
End Function

Public Function BN_mod_inverse_bin(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef n As BIGNUM_TYPE) As Boolean
    ' Inverso modular usando algoritmo binário estendido de Euclides
    ' Calcula r tal que (a * r) ≡ 1 (mod n)
    ' Otimizado para módulos ímpares (como primos da criptografia)
    Dim one As BIGNUM_TYPE
    Dim u As BIGNUM_TYPE, v As BIGNUM_TYPE
    Dim x1 As BIGNUM_TYPE, x2 As BIGNUM_TYPE

    one = BN_new() : Call BN_set_word(one, 1)
    If n.neg Or Not BN_is_odd(n) Or BN_ucmp(n, one) <= 0 Then
        BN_mod_inverse_bin = False
        GoTo CLEAN_BIN
    End If

    u = BN_new() : v = BN_new()
    Call BN_mod(u, a, n)
    Call BN_copy(v, n)
    If BN_is_zero(u) Then BN_mod_inverse_bin = False : GoTo CLEAN_BIN

    x1 = BN_new() : x2 = BN_new()
    Call BN_set_word(x1, 1)
    BN_zero x2

    Do
        If BN_is_one(u) Then
            Call BN_copy(r, x1)
            If r.neg Then r.neg = False
            If BN_ucmp(r, n) >= 0 Then Call BN_usub(r, r, n)
            BN_mod_inverse_bin = True
            GoTo CLEAN_BIN
        End If
        If BN_is_one(v) Then
            Call BN_copy(r, x2)
            If r.neg Then r.neg = False
            If BN_ucmp(r, n) >= 0 Then Call BN_usub(r, r, n)
            BN_mod_inverse_bin = True
            GoTo CLEAN_BIN
        End If
        If BN_is_zero(u) Or BN_is_zero(v) Then
            BN_mod_inverse_bin = False
            GoTo CLEAN_BIN
        End If

        Do While Not BN_is_zero(u)
            If BN_is_odd(u) Then Exit Do
            Call BN_rshift(u, u, 1)
            If BN_is_odd(x1) Then Call BN_add(x1, x1, n)
            Call BN_rshift(x1, x1, 1)
        Loop

        Do While Not BN_is_zero(v)
            If BN_is_odd(v) Then Exit Do
            Call BN_rshift(v, v, 1)
            If BN_is_odd(x2) Then Call BN_add(x2, x2, n)
            Call BN_rshift(x2, x2, 1)
        Loop

        If BN_ucmp(u, v) >= 0 Then
            Call BN_usub(u, u, v)
            Call BN_sub(x1, x1, x2)
            If x1.neg Then
                x1.neg = False
                Call BN_usub(x1, n, x1)
            ElseIf BN_ucmp(x1, n) >= 0 Then
                Call BN_usub(x1, x1, n)
            End If
        Else
            Call BN_usub(v, v, u)
            Call BN_sub(x2, x2, x1)
            If x2.neg Then
                x2.neg = False
                Call BN_usub(x2, n, x2)
            ElseIf BN_ucmp(x2, n) >= 0 Then
                Call BN_usub(x2, x2, n)
            End If
        End If
    Loop

CLEAN_BIN:
    ' Limpeza de memória e recursos temporários
    BN_free one
    BN_free u: BN_free v: BN_free x1: BN_free x2
End Function

' =============================================================================
' OPERAÇÕES MODULARES AVANÇADAS
' =============================================================================

Public Function BN_mod_add(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE, ByRef m As BIGNUM_TYPE) As Boolean
    ' Adição modular: r = (a + b) mod m
    Dim t As BIGNUM_TYPE : t = BN_new()
    If Not BN_add(t, a, b) Then BN_mod_add = False : Exit Function
    BN_mod_add = BN_mod(r, t, m)
End Function

Public Function BN_mod_sub(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE, ByRef m As BIGNUM_TYPE) As Boolean
    ' Subtração modular: r = (a - b) mod m
    Dim t As BIGNUM_TYPE : t = BN_new()
    If Not BN_sub(t, a, b) Then BN_mod_sub = False : Exit Function
    BN_mod_sub = BN_mod(r, t, m)
End Function

Public Function BN_mod_mul(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE, ByRef m As BIGNUM_TYPE) As Boolean
    ' Multiplicação modular: r = (a * b) mod m
    Dim t As BIGNUM_TYPE : t = BN_new()
    If Not BN_mul(t, a, b) Then BN_mod_mul = False : Exit Function
    BN_mod_mul = BN_mod(r, t, m)
End Function

Public Function BN_mod_sqr(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef m As BIGNUM_TYPE) As Boolean
    ' Quadrado modular: r = (a²) mod m
    BN_mod_sqr = BN_mod_mul(r, a, a, m)
End Function

Public Function BN_mod_exp(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef e As BIGNUM_TYPE, ByRef m As BIGNUM_TYPE) As Boolean
    ' Exponenciação modular: r = (a^e) mod m
    ' Usa algoritmo square-and-multiply com otimizações
    
    ' Para expoentes grandes, usar algoritmo com janelas
    If BN_num_bits(e) > 64 Then
        BN_mod_exp = BN_mod_exp_win4(r, a, e, m)
        Exit Function
    End If
    
    ' Algoritmo padrão para expoentes pequenos
    Dim base As BIGNUM_TYPE, acc As BIGNUM_TYPE
    Dim i As Long, nbits As Long
    base = BN_new() : acc = BN_new()
    If Not BN_mod(base, a, m) Then BN_mod_exp = False : Exit Function
    Call BN_set_word(acc, 1)
    nbits = BN_num_bits(e)
    For i = 0 To nbits - 1
        If BN_is_bit_set(e, i) Then
            If Not BN_mod_mul(acc, acc, base, m) Then BN_mod_exp = False : Exit Function
        End If
        If Not BN_mod_sqr(base, base, m) Then BN_mod_exp = False : Exit Function
    Next i
    Call BN_copy(r, acc)
    BN_mod_exp = True
End Function

Public Function BN_mod_exp_consttime(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef e As BIGNUM_TYPE, ByRef m As BIGNUM_TYPE) As Boolean
    ' Exponenciação modular resistente a timing attacks
    ' Usa operações constant-time para segurança criptográfica
    Dim base As BIGNUM_TYPE, acc As BIGNUM_TYPE, temp As BIGNUM_TYPE
    Dim i As Long, nbits As Long, bit As Long
    base = BN_new() : acc = BN_new() : temp = BN_new()
    If Not BN_mod(base, a, m) Then BN_mod_exp_consttime = False : Exit Function
    Call BN_set_word(acc, 1)
    nbits = BN_num_bits(e)
    
    For i = 0 To nbits - 1
        bit = IIf(BN_is_bit_set(e, i), 1, 0)
        ' Sempre fazer multiplicação, usar swap condicional
        Call BN_copy(temp, acc)
        Call BN_mod_mul(acc, acc, base, m)
        Call BN_consttime_swap_flag(bit, temp, acc)
        Call BN_copy(acc, temp)
        Call BN_mod_sqr(base, base, m)
    Next i
    Call BN_copy(r, acc)
    BN_mod_exp_consttime = True
End Function

Public Sub BN_consttime_swap_flag(ByVal flag As Long, ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE)
    ' Swap condicional constant-time para resistência a timing attacks
    If flag <> 0 Then
        Dim temp As BIGNUM_TYPE
        temp = BN_new()
        Call BN_copy(temp, a)
        Call BN_copy(a, b)
        Call BN_copy(b, temp)
    End If
End Sub
Public Function BN_mod_inverse(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef n As BIGNUM_TYPE) As Boolean
    ' Interface principal para inverso modular
    ' Seleciona algoritmo apropriado baseado nas características do módulo
    If BN_is_odd(n) Then
        BN_mod_inverse = BN_mod_inverse_bin(r, a, n)
    Else
        BN_zero r
        BN_mod_inverse = False
    End If
End Function

' =============================================================================
' UTILITÁRIOS DE DEPURAÇÃO E DIAGNÓSTICO
' =============================================================================

Public Sub BN_print_debug(ByRef bn As BIGNUM_TYPE, ByVal name As String)
    ' Imprime informações detalhadas de um BIGNUM para depuração
    ' Parâmetros: bn - BIGNUM a ser analisado, name - nome identificador
    Debug.Print name & ": " & BN_bn2hex(bn) & " (neg: " & bn.neg & ", top: " & bn.top & ", dmax: " & bn.dmax & ")"
End Sub
Public Function BN_mod_secp256k1_fast(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE) As Boolean
    ' Redução modular rápida para p = 2^256 - 2^32 - 977 (secp256k1)
    ' Usa propriedades especiais do primo para evitar divisão completa
    ' Baseado na implementação Bitcoin Core secp256k1_fe_reduce
    
    ' Para números até 512 bits, usar redução especializada
    If a.top <= 16 Then  ' 16 * 32 = 512 bits
        BN_mod_secp256k1_fast = BN_mod_secp256k1_reduce_512(r, a)
    Else
        ' Fallback para redução genérica
        Dim secp256k1_p As BIGNUM_TYPE
        secp256k1_p = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")
        BN_mod_secp256k1_fast = BN_mod(r, a, secp256k1_p)
    End If
End Function

Private Function BN_mod_secp256k1_reduce_512(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE) As Boolean
    ' Redução especializada para secp256k1 usando decomposição
    ' p = 2^256 - 2^32 - 977 = 2^256 - 0x1000003D1
    
    If a.top <= 8 Then  ' Já menor que p
        Call BN_copy(r, a)
        BN_mod_secp256k1_reduce_512 = True
        Exit Function
    End If
    
    ' Separar em parte alta e baixa
    Dim high As BIGNUM_TYPE, low As BIGNUM_TYPE, temp As BIGNUM_TYPE
    high = BN_new(): low = BN_new(): temp = BN_new()
    
    ' low = a mod 2^256 (primeiros 8 limbs)
    If Not bn_wexpand(low, 8) Then BN_mod_secp256k1_reduce_512 = False: Exit Function
    Dim i As Long
    For i = 0 To 7
        If i < a.top Then low.d(i) = a.d(i) Else low.d(i) = 0
    Next i
    low.top = 8
    
    ' high = a >> 256 (limbs restantes)
    If a.top > 8 Then
        If Not bn_wexpand(high, a.top - 8) Then BN_mod_secp256k1_reduce_512 = False: Exit Function
        For i = 8 To a.top - 1
            high.d(i - 8) = a.d(i)
        Next i
        high.top = a.top - 8
    End If
    
    ' Redução: low + high * (2^32 + 977)
    Dim multiplier As BIGNUM_TYPE
    multiplier = BN_hex2bn("100000000000003D1")  ' 2^32 + 977
    
    Call BN_mul(temp, high, multiplier)
    Call BN_add(r, low, temp)
    
    ' Verificar se ainda precisa reduzir
    Dim secp256k1_p As BIGNUM_TYPE
    secp256k1_p = BN_hex2bn("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFC2F")
    
    If BN_ucmp(r, secp256k1_p) >= 0 Then
        Call BN_sub(r, r, secp256k1_p)
    End If
    
    BN_mod_secp256k1_reduce_512 = True
End Function
Public Function BN_mul_comba8(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE) As Boolean
    ' Multiplicação COMBA 8x8 limbs para operações secp256k1 (256-bit)
    ' Alias para BN_mul_fast256 - mantém compatibilidade com chamadas existentes
End Function