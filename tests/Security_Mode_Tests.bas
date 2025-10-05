Attribute VB_Name = "Security_Mode_Tests"
Option Explicit
Option Compare Binary
Option Base 0

'==============================================================================
' TESTES DE SEGURANÇA: MODO CONSTANT-TIME
'==============================================================================

Public Sub Test_Security_Mode_Defaults()
    ' Garante que o modo constant-time permanece habilitado por padrão mesmo após resets de contexto
    Call secp256k1_init
    If Not require_constant_time() Then
        Err.Raise vbObjectError + &H2001&, "Test_Security_Mode_Defaults", _
                  "O modo constant-time deveria estar habilitado imediatamente após secp256k1_init."
    End If

    ' Desativar explicitamente para simular cenários de benchmark
    Call disable_security_mode
    If require_constant_time() Then
        Err.Raise vbObjectError + &H2002&, "Test_Security_Mode_Defaults", _
                  "O modo constant-time deveria estar desabilitado após chamar disable_security_mode."
    End If

    ' Resetar o contexto para emular uma nova sessão
    Call secp256k1_reset_context_for_tests
    If Not require_constant_time() Then
        Err.Raise vbObjectError + &H2003&, "Test_Security_Mode_Defaults", _
                  "O modo constant-time deveria ser restaurado para True após resetar o contexto."
    End If

    ' Reexecutar a inicialização completa e validar novamente
    Call secp256k1_init
    If Not require_constant_time() Then
        Err.Raise vbObjectError + &H2004&, "Test_Security_Mode_Defaults", _
                  "O modo constant-time deveria permanecer habilitado após nova inicialização."
    End If
End Sub
