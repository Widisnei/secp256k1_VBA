Attribute VB_Name = "Test_Table_Data"
Option Explicit

Public Sub test_table_data_format()
    Debug.Print "=== TESTE FORMATO DADOS TABELA ==="

    Call secp256k1_init
    Dim ctx As SECP256K1_CTX
    ctx = secp256k1_context_create()

    ' Testar usando Manager (forma correta como em Test_Tabelas_Integration)
    Dim scalar_test As BIGNUM_TYPE, point_fast As EC_POINT, point_regular As EC_POINT
    scalar_test = BN_hex2bn("2")  ' Testar 2*G
    point_fast = ec_point_new()
    point_regular = ec_point_new()

    ' Multiplicar usando tabelas (Manager)
    Dim fast_ok As Boolean
    fast_ok = EC_Precomputed_Manager.ec_generator_mul_fast(point_fast, scalar_test, ctx)

    ' Multiplicar usando método regular
    Call ec_point_mul(point_regular, scalar_test, ctx.g, ctx)

    Debug.Print "Multiplicação com tabelas: " & fast_ok
    Debug.Print "Ponto tabela na curva: " & ec_point_is_on_curve(point_fast, ctx)
    Debug.Print "Ponto regular na curva: " & ec_point_is_on_curve(point_regular, ctx)

    ' Comparar resultados
    If ec_point_cmp(point_fast, point_regular, ctx) = 0 Then
        Debug.Print "SUCESSO: Tabelas funcionando corretamente"
    Else
        Debug.Print "AVISO: Diferença entre tabelas e multiplicação regular"
    End If

    Debug.Print "=== TESTE CONCLUÍDO ==="
End Sub

Public Sub test_bitcoin_core_format()
    Debug.Print "=== TESTE FORMATO BITCOIN CORE ==="

    Dim entry As String
    entry = get_gen_point(0, 1)
    Debug.Print "Entrada raw: " & entry

    Debug.Print "=== TESTE CONCLUÍDO ==="
End Sub