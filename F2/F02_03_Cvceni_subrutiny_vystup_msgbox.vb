Module F02_03_Cvceni_subrutiny_vystup_msgbox
    '3) Cvičení na subrutinu: Napište subrutinu, která vytiskne řádek obsahující N znaků Z.
    'Údaje N a Z jsou vstupními parametry subrutiny.
    'Pozn.: toto nelze provést jako funkci! Pozn.: použití MsgBox v subrutině či funkci je výjimečné,
    'nepoužívejte tento postup v jiných programech.
    Sub Mainx()
        TiskniZ(5, "Z")
    End Sub

    Sub TiskniZ(n As Long, ByRef z As String)
        Dim i As Long
        Dim vystup As String

        vystup = ""

        For i = 1 To n
            vystup += "Z"
        Next
        MsgBox(vystup)
    End Sub
End Module
