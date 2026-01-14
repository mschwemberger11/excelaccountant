Attribute VB_Name = "KopierBuchung"
'*****************************************************************************
'* KopierBuchen *   Tastenkombination: Strg+c                                *
'*****************************************************************************

Public KopierteZeilenKorrekt As Boolean
Sub BuchZeileKopieren()
Attribute BuchZeileKopieren.VB_ProcData.VB_Invoke_Func = "c\n14"
'Im Blatt ArProt werden die Spalten D mit G der Herzeile (bzw. der Herzeilen) in
'die Spalten D mit G der Hinzeile (bzw. der Hinzeilen) kopiert. Die (erste) HinZeile
'ist durch die beim Aufruf aktive Zelle gegeben. Die erste HerZeile und die Anzahl der
'Zeilen werden über eine Input-Box erfragt.
'In die H-Spalte der Hinzeile (der Hinzeilen) wird das Kopierbuchungszeichen "*****"
'eingefügt. Dann wird eine neue ArProt-Zeile für die nächste Buchung erzeugt.
'Ab der zweiten Hinzeile werden die Spalten B (Transaktionsdatum) und C (Belegnummer)
'von der ersten Hinzeile kopiert; im Falle der Spalte C aber nur, wenn die Belegnummer
'nichtnumerisch ist. Sonst wird eine konstante Differenz zwischen Belegnummern der
'Herzeile und der Hinzeile eingehalten.
Dim HerZeile As Long, HinZeile As Long, APZeile As Long
Dim HinDatum, HerBeleg, HinBeleg, BelegDifferenz As Long
Dim BelegNrNumerisch As Boolean
Dim AZelle As Range, HerBereich As Range
Dim AnzahlZeilen As Integer

'1 Kop -------------------------- Blatt ArProt erzwingen --------------------------
With ActiveWindow
  KopierteZeilenKorrekt = False
  If ActiveSheet.Name <> "ArProt" Then
    A = MsgBox("kann nur vom Blatt ''ArProt'' aus verwendet werden." & Chr(10) & _
               "Dorthin wechseln?", vbOKCancel, "Tastenkombination ''Strg+c'' ")
    If A = vbOK Then                 'kein Aktivieren von ArProt, wenn
      Worksheets("ArProt").Activate  'mit Abbrechen quittiert wird
        Cells(Cells(1, 1).Value + 2, 2).Activate
    End If
    Exit Sub
  End If
  With Worksheets("ArProt")
    .Activate
    Set AZelle = ActiveCell
    APZeile = ActiveCell.Row
'2 Kop -------------- Zeilen Kopieren nur aus Spalte 4 "Sollkonto" ----------------
    If ActiveCell.Column <> 4 Then
      A = MsgBox("Für ZEILENKOPIEREN muss" & Chr(10) & Chr(10) & _
                 "eine Zelle in Spalte ''D'' (Sollkonto)" & _
                 Chr(10) & Chr(10) & "aktiviert sein!", 0, _
                 "Tastenkombination ''Strg+c'' ")
      Exit Sub
    End If
'3 Kop ----------------- Außerhalb des ArProt-Arbeitsbereichs ---------------------
    If ActiveCell.Column > 10 Or ActiveCell.Row < 3 Then
      MsgBox ("Tastenkombination ''Strg+c'' hier wirkungslos")
      Exit Sub
    End If
'4 Kop ---------------------- Ermitteln HerZeile und HinZeile ---------------------
    HinZeile = APZeile
    Set HerBereich = Application.InputBox(prompt:= _
       "Bitte eine oder mehrere Spalten der hierher zu kopierenden" & _
       "Zeilen selektieren und Knopf ''OK'' drücken." & Chr(10) & Chr(10) & _
       "(Wenn Knopf ''Abbrechen'' gedrückt wird, in der folgenden" & Chr(10) & _
       "Laufzeitfehlerbox den Knopf ''Beenden'' drücken!)", _
       Title:="Kopieren von Buchungen ähnlichen Inhalts", _
       Default:="$A$0:$A$0", Type:=8, Left:=450, Top:=300)
    HerBereich.Select
    AnzahlZeilen = Selection.Rows.Count
    HerZeile = Selection.Row
'5 Kop -------------------- Anfangswerte Datum und Belegnummern --------------------
    HinDatum = Cells(HinZeile, 2).Value
    HerBeleg = Cells(HerZeile, 3).Value
    HinBeleg = Cells(HinZeile, 3).Value
    If IsNumeric(HerBeleg) = True And IsNumeric(HinBeleg) = True Then
      BelegNrNumerisch = True
      BelegDifferenz = HinBeleg - HerBeleg
    End If
'6 Kop --------------- Kopieren D:G der HerZeile in die HinZeile -------------------
KopierVorgang:
    HerBeleg = Cells(HerZeile, 3).Value
    Cells(HerZeile, 4).Activate
    ActiveCell.Range("A1:D1").Select
    Selection.Copy
    Cells(HinZeile, 4).Activate
    ActiveCell.Range("A1:D1").Select
    ActiveSheet.Paste
'7 Kop --------------- CopyMode beenden, Kopiersterne in Spalte H ------------------
    Cells(HinZeile, 8).Activate
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "*****"
'8 Kop ------------------ Datum und Beleg der HinZeile kopieren --------------------
    Cells(HinZeile, 2).Value = HinDatum
    If IsNumeric(HerBeleg) = True Then
      Cells(HinZeile, 3).Value = HerBeleg + BelegDifferenz
    Else
      Cells(HinZeile, 3).Value = HinBeleg
    End If
'9 Kop -------------------------- Nächste ArProt-Zeile erzeugen --------------------
    Cells(HinZeile + 1, 1).EntireRow.Select
    Selection.Insert shift:=xlDown
    Cells(1, 1).Value = Cells(1, 1).Value + 1
    Cells(HinZeile + 1, 1).Value = Cells(HinZeile, 1).Value + 1
    Cells(HinZeile + 1, 8).Value = "***"
    Cells(HinZeile + 1, 2).Activate
'10 Kop -------------------------- weitere Zeilen kopieren? ------------------------
    AnzahlZeilen = AnzahlZeilen - 1
    If AnzahlZeilen > 0 Then
      HerZeile = HerZeile + 1
      HinZeile = HinZeile + 1
      GoTo KopierVorgang
    End If
  End With
End With 'ActiveWindow
End Sub  'BuchZeileKopieren



 
