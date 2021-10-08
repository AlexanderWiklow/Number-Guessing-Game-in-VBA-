Attribute VB_Name = "Module1"
Option Explicit

Sub PracticeRandom()
    Dim Tal1, Val, Svar, AntalR, AntalF, MaxV, i, j
    
    Application.ScreenUpdating = False
    
    'Se till att unika slumptal skapas, funktion/systemanrop? --
    Randomize
    
    'Initiera räknarvariabler till noll --
    AntalR = 0
    AntalF = 0
    
    i = 0
    
    Range("A2").Select
    Do Until Val = vbNo
    
        'Generera två slumptal mellan 1 och 200 --
        'Int konverterar till heltal och tar bort decimalerna utan att avrunda värdet --
        Tal1 = Int(Rnd * 200 + 1)

        Svar = InputBox("Gissa ett tal mellan 0 - 200")

        'Kontrollerar om användaren svarat rätt och ge feedback --
        If Tal1 = Svar Then
            MsgBox "Bravo rätt svar", vbInformation
            AntalR = AntalR + 1
        Else
            MsgBox "Tyvärr fel svar!", vbExclamation
        End If

        'Räkna upp antalet försök
        AntalF = AntalF + 1
        
        ActiveCell.Value = "Tal " & i + 1
        ActiveCell.Offset(0, 1) = Tal1
        
        ActiveCell.Offset(1, 0).Select

        'Frågar användaren om hen vill ha ett nytt svar --
        Val = MsgBox("Vill du ha ett nytt tal?", vbYesNo)
        
        i = i + 1
        
    'Får loopen att repetera --
    Loop
    
    'Ger feedback till användaren
    MsgBox "Du hade " & AntalR & " rätt av " & AntalF & " försök"
    
        Application.ScreenUpdating = True

End Sub


