Module F02_01_Uspory_osob_a_uroky
    '1) Zadejte ve smyčce do pole úspory několika osob. Poté se program zeptá, jaká je roční úroková míra.
    'Pole předejte funkci, která vrátí pole po započítání úroků. Původní pole nesmí být
    'voláním funkce ovlivněno (to platí i pro všechny další příklady).
    Sub Mainx()
        MsgBox("Jsem F2-01")

        Dim pole_uspor() As Single, pole_urokovane() As Single
        Dim zadana_uspora As Single
        Dim i As Byte
        Dim i_2 As Byte
        Dim txt_vystup As String, txt_vystup_z_pole As String

        txt_vystup = ""
        txt_vystup_z_pole = ""
        i = 0
        zadana_uspora = InputBox("Zadej uspořenou částku: ")

        Do
            ReDim Preserve pole_uspor(i)
            pole_uspor(i) = zadana_uspora
            i += 1
            zadana_uspora = InputBox("Zadej uspořenou částku: ")
        Loop While zadana_uspora <> 0

        For i = 0 To UBound(pole_uspor)
            txt_vystup += Str(pole_uspor(i)) + ", "
        Next

        pole_urokovane = SpoctiUrok(pole_uspor, 10)

        For i_2 = 0 To UBound(pole_urokovane)
            txt_vystup_z_pole += Str(pole_urokovane(i_2)) + ", "
        Next


        MsgBox(txt_vystup + Chr(10) + txt_vystup_z_pole)
    End Sub

    Function SpoctiUrok(pole() As Single, urok As Single) As Single()
        Dim prepocet_uroku As Single
        Dim odkladaci As Single
        Dim vnitrni_pole() As Single
        Dim i As Byte

        prepocet_uroku = urok / 100

        For i = 0 To UBound(pole)
            ReDim Preserve vnitrni_pole(i)
            odkladaci = pole(i) + pole(i) * prepocet_uroku
            vnitrni_pole(i) = odkladaci
        Next
        Return vnitrni_pole
    End Function
End Module
