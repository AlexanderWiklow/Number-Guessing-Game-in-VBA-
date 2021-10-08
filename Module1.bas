Attribute VB_Name = "Module1"
Option Explicit

Sub PracticeRandom()
    Dim Tal1, Val, Svar, AntalR, AntalF, MaxV, i, j
    
    Application.ScreenUpdating = False
    
    'Se till att unika slumptal skapas, funktion/systemanrop? --
    Randomize
    
    'Initiera r�knarvariabler till noll --
    AntalR = 0
    AntalF = 0
    
    i = 0
    
    Range("A2").Select
    Do Until Val = vbNo
    
        'Generera tv� slumptal mellan 1 och 200 --
        'Int konverterar till heltal och tar bort decimalerna utan att avrunda v�rdet --
        Tal1 = Int(Rnd * 200 + 1)

        Svar = InputBox("Gissa ett tal mellan 0 - 200")

        'Kontrollerar om anv�ndaren svarat r�tt och ge feedback --
        If Tal1 = Svar Then
            MsgBox "Bravo r�tt svar", vbInformation
            AntalR = AntalR + 1
        Else
            MsgBox "Tyv�rr fel svar!", vbExclamation
        End If

        'R�kna upp antalet f�rs�k
        AntalF = AntalF + 1
        
        ActiveCell.Value = "Tal " & i + 1
        ActiveCell.Offset(0, 1) = Tal1
        
        ActiveCell.Offset(1, 0).Select

        'Fr�gar anv�ndaren om hen vill ha ett nytt svar --
        Val = MsgBox("Vill du ha ett nytt tal?", vbYesNo)
        
        i = i + 1
        
    'F�r loopen att repetera --
    Loop
    
    'Ger feedback till anv�ndaren
    MsgBox "Du hade " & AntalR & " r�tt av " & AntalF & " f�rs�k"
    
        Application.ScreenUpdating = True

End Sub


