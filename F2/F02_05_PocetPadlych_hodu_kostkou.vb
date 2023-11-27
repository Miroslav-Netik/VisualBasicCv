Imports System.Net.Sockets

Module F02_05_PocetPadlych_hodu_kostkou
    '5) Sestavte funkci PocetPadlych, která kromě předaného pole má i parametr, kolikrát se kostkou hodilo.
    'Výsledkem funkce je jednorozměrné celočíselné pole udávající, kolikrát padla ta která hodnota na kostce.
    'Funkci ověřte na zadání: sestavte program, který se zeptá na počet hodů, vygeneruje pole a pak zobrazí,
    'kolikrát každé číslo padlo. Povšimněte si, že čím více hodů, tím více se počty vyrovnají
    '(podobné e9, ale zde je to s použitím funkce)
    Sub Mainx()
        Dim pole_hodu() As Integer = {}
        Dim pole_vysledku() As Integer = {}
        Dim i As Integer, j As Integer, pocet_hodu As Integer
        Dim txt_vystup As String, txt_hody As String

        txt_hody = ""
        txt_vystup = ""

        pocet_hodu = InputBox("Zadej počet hodů: ")

        ReDim pole_hodu(pocet_hodu)

        Randomize()

        For i = 1 To pocet_hodu
            pole_hodu(i) = Int(Rnd() * 6 + 1)
            txt_hody += Str(pole_hodu(i)) + ","
        Next

        pole_vysledku = PocetPadlych(pole_hodu, pocet_hodu)


        txt_vystup += "1 padla" + Str(pole_vysledku(0)) + " krát" + Chr(10) _
                    + "2 padla" + Str(pole_vysledku(1)) + " krát" + Chr(10) _
                    + "3 padla" + Str(pole_vysledku(2)) + " krát" + Chr(10) _
                    + "4 padla" + Str(pole_vysledku(3)) + " krát" + Chr(10) _
                    + "5 padla" + Str(pole_vysledku(4)) + " krát" + Chr(10) _
                    + "6 padla" + Str(pole_vysledku(5)) + " krát"

        MsgBox(txt_hody + Chr(10) + txt_vystup)
    End Sub
    Function PocetPadlych(vstupni_pole() As Integer, pocet As Integer) As Integer()
        'Proč tam je ten počet?
        Dim pole_pocitadel(5) As Integer
        Dim i As Integer

        For i = 0 To UBound(vstupni_pole)
            Select Case vstupni_pole(i)
                Case 1
                    pole_pocitadel(0) = pole_pocitadel(0) + 1
                Case 2
                    pole_pocitadel(1) = pole_pocitadel(1) + 1
                Case 3
                    pole_pocitadel(2) = pole_pocitadel(2) + 1
                Case 4
                    pole_pocitadel(3) = pole_pocitadel(3) + 1
                Case 5
                    pole_pocitadel(4) = pole_pocitadel(4) + 1
                Case 6
                    pole_pocitadel(5) = pole_pocitadel(5) + 1
            End Select
        Next

        Return pole_pocitadel
    End Function
End Module
