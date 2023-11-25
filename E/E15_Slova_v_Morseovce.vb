Module E15_Slova_v_Morseovce
    '15) Uživatel zadá několik slov (bez diakritiky). Program zapíše tento text v Morseově abecedě.
    'Oddělovačem znaků bude lomítko. zadat všechna slova do jednoho inputboxu, oddělená normálně mezerami.
    'A mezera se překládá taky jako lomítko. Tedy na hranicích slov budou vlastně dvě lomítka.- (Nutná znalost Morseovy abecedy). 
    Sub Mainx()
        MsgBox("Jsem E15")
        Dim pole_morseovky(25) As String
        Dim i As Integer
        Dim txt_vystup As String, txt_zadani As String, zadane_pismeno As String
        Dim prelozene_pismeno As String

        pole_morseovky(0) = ".-" ' A
        pole_morseovky(1) = "-..." ' B
        pole_morseovky(2) = "-.-." ' C
        pole_morseovky(3) = "-.." 'D
        pole_morseovky(4) = "." 'E
        pole_morseovky(5) = "..-." 'F
        pole_morseovky(6) = "--." 'G
        pole_morseovky(7) = "...." 'H
        pole_morseovky(8) = ".." 'I
        pole_morseovky(9) = ".---" 'J
        pole_morseovky(10) = "-.-" 'K
        pole_morseovky(11) = ".-.." 'L
        pole_morseovky(12) = "--" 'M
        pole_morseovky(13) = "-." 'N
        pole_morseovky(14) = "---" 'O
        pole_morseovky(15) = ".--." 'P
        pole_morseovky(16) = "--.-" 'Q
        pole_morseovky(17) = ".-." 'R
        pole_morseovky(18) = "..." 'S
        pole_morseovky(19) = "-" 'T
        pole_morseovky(20) = "..-" 'U
        pole_morseovky(21) = "...-" 'V
        pole_morseovky(22) = ".--" 'W
        pole_morseovky(23) = "-..-" 'X
        pole_morseovky(24) = "-.--" 'Y
        pole_morseovky(25) = "--.." 'Z

        prelozene_pismeno = ""
        txt_vystup = ""

        txt_zadani = InputBox("Zadej text pro převod do Moreovky: ")

        txt_zadani = UCase(txt_zadani) 'Převede všechno na velká písmena, morseovce je to jedno
        'Rozdělení slova na znaky oddělené "/" a nahrazení mezery za "/"
        For i = 1 To Len(txt_zadani)
            zadane_pismeno = Mid(txt_zadani, i, 1)
            If zadane_pismeno = " " Then
                prelozene_pismeno = "/"
            Else
                'Platí pro velká písmena, pro malá, pokud bych je nechal by bylo -97
                prelozene_pismeno = pole_morseovky(Asc(zadane_pismeno) - 65) + "/"
            End If
            txt_vystup += prelozene_pismeno

        Next

        MsgBox(txt_vystup)
    End Sub
End Module
