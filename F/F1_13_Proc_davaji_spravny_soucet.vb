Module F1_13_Proc_davaji_spravny_soucet
    '13) Zjistěte, proč ve Visual Studiu dává tento program i u subrutiny správný součet:
    Sub Mainx_1()

        Dim soucet As Byte
        Dim ret As String

        ret = ""

        soucet = SectiF(3, 4)
        ret = ret + Chr(10) + "Soucet podle funkce je " + Str(soucet)
        soucet = 8

        SectiS(3, 4, soucet)
        ret = ret + Chr(10) + "Soucet  podle subrutiny je " + Str(soucet)

        MsgBox(ret)
    End Sub

    Function SectiF(a As Byte, b As Byte) As Byte
        SectiF = a + b
    End Function

    Sub SectiS(a As Byte, b As Byte, ByVal vys As Byte)
        vys = a + b + 5
    End Sub

End Module
