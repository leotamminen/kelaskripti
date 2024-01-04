Sub Junamakro()
'
' Lasketaan junassa tulevien, eli kelojen, joilla on "JUNA" varastopaikka, määrä
'
'
    With Session
        .TransmitTerminalKey rcIBMClearKey
        .WaitForEvent rcKbdEnabled, "30", "0", 1, 1
        .WaitForEvent rcEnterPos, "30", "0", 1, 1
        .TransmitANSI "/for modb127"
        .TransmitTerminalKey rcIBMEnterKey
        .WaitForEvent rcKbdEnabled, "30", "0", 1, 1
        .WaitForEvent rcEnterPos, "30", "0", 2, 31
        .WaitForDisplayString "TYÖ:", "30", 2, 26
        .TransmitTerminalKey rcIBMTabKey
        .TransmitTerminalKey rcIBMTabKey
        .TransmitTerminalKey rcIBMTabKey
        .TransmitTerminalKey rcIBMTabKey
        .TransmitANSI "015"
        .TransmitTerminalKey rcIBMEnterKey
'
' Tämä alla oleva kohta kopioi kaikki sivulle mahtuvat varastopaikat leikepöydälle.
'
        Dim str As String, Viesti As String, NimiJotaHaetaan As String, Juna_maara As Integer, Kokonaislaskuri As Integer
        Juna_maara = 0
        Kokonaislaskuri = 0
        NimiJotaHaetaan = "JUNA"
                
kopsaus:

        .WaitForEvent rcKbdEnabled, "30", "0", 1, 1
        .WaitForEvent rcEnterPos, "30", "0", 3, 60
        .WaitForDisplayString "LAITE/T:", "30", 3, 51
        .TransmitTerminalKey rcIBMNewLineKey
        .TransmitTerminalKey rcIBMNewLineKey
        .TransmitTerminalKey rcIBMNewLineKey
        .TransmitTerminalKey rcIBMTabKey
        .TransmitTerminalKey rcIBMTabKey
        .TransmitTerminalKey rcIBMTabKey
        .TransmitTerminalKey rcIBMRightKey
        .TransmitTerminalKey rcIBMRightKey
   
        .SetSelectionStartPos 7, 66
        .WaitForEvent rcEnterPos, "30", "0", 7, 66
        .ExtendSelectionRect 27, 69
        .CopySelection
        str = GetClipboardText()
       
        If Len(str) < 1 Then GoTo loppu
           
            Juna_maara = (Len(str) - Len(Replace(str, NimiJotaHaetaan, ""))) / Len(NimiJotaHaetaan)
            Kokonaislaskuri = Juna_maara + Kokonaislaskuri
           
    If Len(str) <> 0 Then
            If GetDisplayText(30, 19, 10) <> "JATKUU PA1" Then GoTo loppu
    End If
   
            .WaitForDisplayString "KIRJOITIN:", "30", 3, 66
            .TransmitTerminalKey rcIBMPA1Key
           
            If Len(str) <> 0 Then
                If GetDisplayText(30, 19, 10) = "JATKUU PA1" Then GoTo kopsaus
            End If
       
loppu:

    MsgBox ("Sijoittamattomia keloja yht " & CStr(Kokonaislaskuri) & " kpl.")
   

    End With
End Sub