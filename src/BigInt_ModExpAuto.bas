Attribute VB_Name = "BigInt_ModExpAuto"
Option Explicit
Option Compare Binary
Option Base 0

' =============================================================================
' BIGINT MODULAR EXPONENTIATION AUTO VBA - SELETOR INTELIGENTE
' =============================================================================
' Implementa seletor automático de algoritmo de exponenciação modular
' Analisa características do expoente (tamanho e densidade) para escolher
' o algoritmo mais eficiente: básico vs windowing de 4 bits
' Otimizado para casos comuns como RSA (e=65537) e criptografia ECC
' =============================================================================

Public BN_mod_exp_auto_last_algorithm As String

Public Sub BN_mod_exp_auto_reset_diagnostics()
    BN_mod_exp_auto_last_algorithm = ""
End Sub

Private Sub BN_mod_exp_auto_record_backend(ByVal backendName As String)
    BN_mod_exp_auto_last_algorithm = backendName
End Sub

' =============================================================================
' CONTAGEM RÁPIDA DE BITS (POPCOUNT)
' =============================================================================

Private Function PC8(ByVal b As Integer) As Integer
    ' Conta bits definidos em um byte usando tabela pré-computada
    ' Parâmetro: b - byte a ser analisado (0-255)
    ' Retorna: Número de bits '1' no byte
    ' Otimização: tabela estática para evitar recálculo
    
    Static t(0 To 255) As Integer, init As Boolean
    Dim i As Integer
    
    ' Inicializar tabela de popcount na primeira chamada
    If Not init Then
        For i = 0 To 255
            ' Contar cada bit individualmente para construir tabela
            t(i) = ((i And &H1) <> 0) + ((i And &H2) <> 0) + ((i And &H4) <> 0) + ((i And &H8) <> 0) + _
                   ((i And &H10) <> 0) + ((i And &H20) <> 0) + ((i And &H40) <> 0) + ((i And &H80) <> 0)
        Next i
        init = True
    End If
    
    ' Retornar valor pré-computado da tabela
    PC8 = t(b And &HFF&)
End Function

Public Function BN_popcount(ByRef a As BIGNUM_TYPE) As Long
    ' Conta o número total de bits '1' em um BIGNUM
    ' Parâmetro: a - BIGNUM a ser analisado
    ' Retorna: Número de bits definidos (peso de Hamming)
    ' Usado para determinar densidade do expoente em exponenciação modular
    
    Dim byts() As Byte, i As Long, n As Long
    
    ' Converter BIGNUM para representação binária
    byts = BN_bn2bin(a)
    
    ' Verificar se array foi alocado (BIGNUM não-zero)
    If (Not Not byts) <> 0 Then
        n = UBound(byts) + 1
        ' Somar popcount de cada byte usando tabela otimizada
        For i = 0 To n - 1
            BN_popcount = BN_popcount + PC8(byts(i))
        Next i
    End If
End Function

' =============================================================================
' SELETOR AUTOMÁTICO DE ALGORITMO DE EXPONENCIAÇÃO MODULAR
' =============================================================================

Public Function BN_mod_exp_auto(ByRef r As BIGNUM_TYPE, ByRef a As BIGNUM_TYPE, ByRef e As BIGNUM_TYPE, ByRef m As BIGNUM_TYPE) As Boolean
    ' Seleciona algoritmo ótimo de exponenciação modular baseado nas características do expoente
    ' Parâmetros: r = a^e mod m (resultado), a = base, e = expoente, m = módulo
    ' Estratégia: analisar tamanho e densidade do expoente para escolher algoritmo
    ' Algoritmos: básico (square-and-multiply) vs windowing de 4 bits
    
    Dim nbits As Long, ones As Long
    
    ' Analisar características do expoente
    nbits = BN_num_bits(e)  ' Tamanho em bits
    ones = BN_popcount(e)   ' Número de bits '1' (densidade)

    If require_constant_time() Then
        BN_mod_exp_auto_record_backend "CONSTTIME"
        BN_mod_exp_auto = BigInt_ConstTime.BN_mod_exp_consttime(r, a, e, m)
        Exit Function
    End If

    ' Heurística de seleção baseada em benchmarks empíricos
    ' Para expoentes pequenos ou muito esparsos (ex: RSA e=65537), algoritmo básico é mais rápido
    If nbits <= 160 Or ones <= 8 Or ones < (nbits \ 4) Then
        ' Usar algoritmo básico square-and-multiply
        BN_mod_exp_auto_record_backend "BASIC"
        BN_mod_exp_auto = BN_mod_exp(r, a, e, m)
    Else
        ' Usar windowing de 4 bits para expoentes densos e grandes
        BN_mod_exp_auto_record_backend "WIN4"
        BN_mod_exp_auto = BN_mod_exp_win4(r, a, e, m)
    End If
End Function
