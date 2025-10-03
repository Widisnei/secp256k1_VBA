Attribute VB_Name = "BigInt_Karatsuba"
Option Explicit
Option Compare Binary
Option Base 0

' =============================================================================
' BIGINT KARATSUBA VBA - MULTIPLICAÇÃO RÁPIDA KARATSUBA
' =============================================================================
' Implementação do algoritmo de Karatsuba para multiplicação eficiente
' Complexidade: O(n^1.585) vs O(n^2) da multiplicação clássica
' Otimizado para números grandes (>512 bits) com dispatcher inteligente
' Usa divisão e conquista recursiva para reduzir operações
' =============================================================================
' =============================================================================
' MULTIPLICAÇÃO KARATSUBA RECURSIVA
' =============================================================================

Public Function BN_mul_karatsuba(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE) As Boolean
    ' Implementa algoritmo de Karatsuba para multiplicação eficiente
    ' Parâmetros: r = a * b (resultado), a e b = operandos
    ' Algoritmo: divide números em partes e usa 3 multiplicações em vez de 4
    ' Fórmula: (a1*B + a0) * (b1*B + b0) = a1*b1*B² + ((a1+a0)*(b1+b0) - a1*b1 - a0*b0)*B + a0*b0
    
    Dim na As Long, nb As Long, n As Long
    na = a.top: nb = b.top
    
    ' Usar multiplicação clássica para números pequenos (caso base da recursão)
    If na <= 8 Or nb <= 8 Then
        BN_mul_karatsuba = BN_mul(r, a, b)
        Exit Function
    End If
    
    ' Determinar ponto de divisão: metade do maior operando (arredondado para par)
    n = IIf(na > nb, na, nb)
    If n Mod 2 = 1 Then n = n + 1  ' Garantir que n seja par para divisão uniforme
    n = n \ 2
    
    ' Dividir operandos: a = a1*B + a0, b = b1*B + b0 onde B = 2^(n*32)
    ' a1, b1 = partes altas; a0, b0 = partes baixas
    Dim a0 As BIGNUM_TYPE, a1 As BIGNUM_TYPE, b0 As BIGNUM_TYPE, b1 As BIGNUM_TYPE
    a0 = BN_new(): a1 = BN_new(): b0 = BN_new(): b1 = BN_new()
    
    ' Executar divisão dos números em partes alta e baixa
    Call split_number(a0, a1, a, n)
    Call split_number(b0, b1, b, n)

    ' =============================================================================
    ' ALGORITMO KARATSUBA: 3 MULTIPLICAÇÕES EM VEZ DE 4
    ' =============================================================================

    ' Fórmula: z2 = a1*b1, z0 = a0*b0, z1 = (a1+a0)*(b1+b0) - z2 - z0
    ' Resultado: a*b = z2*B² + z1*B + z0
    Dim z0 As BIGNUM_TYPE, z1 As BIGNUM_TYPE, z2 As BIGNUM_TYPE
    Dim temp1 As BIGNUM_TYPE, temp2 As BIGNUM_TYPE, temp3 As BIGNUM_TYPE
    z0 = BN_new(): z1 = BN_new(): z2 = BN_new()
    temp1 = BN_new(): temp2 = BN_new(): temp3 = BN_new()
    
    ' Primeira multiplicação: z2 = a1 * b1 (partes altas)
    If Not BN_mul_karatsuba(z2, a1, b1) Then BN_mul_karatsuba = False: Exit Function
    
    ' Segunda multiplicação: z0 = a0 * b0 (partes baixas)
    If Not BN_mul_karatsuba(z0, a0, b0) Then BN_mul_karatsuba = False: Exit Function
    
    ' Preparar para terceira multiplicação: somar partes
    ' temp1 = a1 + a0, temp2 = b1 + b0
    If Not BN_add(temp1, a1, a0) Then BN_mul_karatsuba = False: Exit Function
    If Not BN_add(temp2, b1, b0) Then BN_mul_karatsuba = False: Exit Function
    
    ' Terceira multiplicação: temp3 = (a1 + a0) * (b1 + b0)
    If Not BN_mul_karatsuba(temp3, temp1, temp2) Then BN_mul_karatsuba = False: Exit Function
    
    ' Calcular termo cruzado: z1 = temp3 - z2 - z0
    ' Isso dá o coeficiente do termo B na expansão
    If Not BN_sub(z1, temp3, z2) Then BN_mul_karatsuba = False: Exit Function
    If Not BN_sub(z1, z1, z0) Then BN_mul_karatsuba = False: Exit Function

    ' =============================================================================
    ' RECOMBINAÇÃO FINAL: r = z2*B² + z1*B + z0
    ' =============================================================================

    ' Montar resultado final usando deslocamentos e adições
    Call BN_lshift(temp1, z2, n * 64)  ' z2 * B² (deslocar 2n limbs)
    Call BN_lshift(temp2, z1, n * 32)  ' z1 * B (deslocar n limbs)
    
    ' Somar os três termos: z2*B² + z1*B + z0
    If Not BN_add(r, temp1, temp2) Then BN_mul_karatsuba = False: Exit Function
    If Not BN_add(r, r, z0) Then BN_mul_karatsuba = False: Exit Function
    
    ' Definir sinal do resultado e tratar caso especial de zero
    r.neg = (a.neg Xor b.neg)
    If r.top = 0 Then r.neg = False
    
    BN_mul_karatsuba = True
End Function

' =============================================================================
' FUNÇÃO AUXILIAR DE DIVISÃO DE NÚMEROS
' =============================================================================

Private Sub split_number(ByRef low As BIGNUM_TYPE, ByRef high As BIGNUM_TYPE, ByRef num As BIGNUM_TYPE, ByVal n As Long)
    ' Divide número em partes baixa e alta no ponto n
    ' Parâmetros: low = parte baixa, high = parte alta, num = número original, n = ponto de divisão
    ' Fórmula: low = num mod 2^(n*32), high = num >> (n*32)
    
    ' Extrair parte baixa: primeiros n limbs
    Call BN_copy(low, num)
    If low.top > n Then low.top = n  ' Truncar para n limbs
    
    ' Normalizar parte baixa removendo zeros à esquerda
    Do While low.top > 0
        If low.d(low.top - 1) = 0 Then low.top = low.top - 1 Else Exit Do
    Loop
    
    ' Extrair parte alta: deslocar direita por n*32 bits
    Call BN_rshift(high, num, n * 32)
End Sub

' =============================================================================
' DISPATCHER INTELIGENTE DE MULTIPLICAÇÃO
' =============================================================================

Public Function BN_mul_optimized(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef b As BIGNUM_TYPE) As Boolean
    ' Seleciona algoritmo de multiplicação ótimo baseado no tamanho dos operandos
    ' Parâmetros: r = a * b (resultado), a e b = operandos
    ' Estratégia: COMBA para pequenos, Karatsuba para médios, clássico para grandes
    
    Dim total_bits As Long
    total_bits = BN_num_bits(a) + BN_num_bits(b)
    
    ' Seleção baseada em benchmarks empíricos
    If total_bits <= 512 Then
        ' Números pequenos: usar COMBA (mais rápido para até 256-bit)
        BN_mul_optimized = BN_mul_fast256(r, a, b)
    ElseIf total_bits <= 2048 Then
        ' Números médios: usar Karatsuba (vantagem assintótica)
        BN_mul_optimized = BN_mul_karatsuba(r, a, b)
    Else
        ' Números muito grandes: usar clássico (menor overhead recursivo)
        BN_mul_optimized = BN_mul(r, a, b)
    End If
End Function