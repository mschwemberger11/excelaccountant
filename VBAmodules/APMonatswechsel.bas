Attribute VB_Name = "APMonatswechsel"
'**************************************************************************************
'* ArProtMonatswechsel *                                                              *
'**************************************************************************************
'
' Tastenkombination: Strg+m
'
Sub ArProtMonatswechsel()
Attribute ArProtMonatswechsel.VB_ProcData.VB_Invoke_Func = "m\n14"
'Fügt im ArProt eine Spaltenüberschriftszeile an der Stelle der Aktiven Zelle ein,
'indem die darunterliegenden Zeilen nach unten verschoben werden, wodurch sie ihre
'Formate behalten. Die Aktive Zelle muss in der Datumsspalte sein.
  Dim APZelle As Range, APZeile As Variant, ZNr As String
  Dim APZ As Long, AnfangsZeile As Long
With ActiveWindow
  Application.CutCopyMode = False
'MW1 ------------------------ Blatt ArProt erzwingen -------------------------
  If ActiveSheet.Name <> "ArProt" Then
    A = MsgBox("kann nur vom Blatt ''ArProt'' aus verwendet werden." & Chr(10) & _
             "Dorthin wechseln?", vbOKCancel, _
              "Tastenkombination ''Strg+m'' ArProt-Monatswechsel")
    If A = vbOK Then                 'kein Aktivieren von ArProt, wenn
      Worksheets("ArProt").Activate  'mit Abbrechen quittiert wird
      Cells(Cells(1, 1).Value + 2, 2).Activate
    End If
    Exit Sub
  End If
 'MW2 ------------------------Erlaubte Spalte erzwingen -------------------------
  With Worksheets("ArProt")
    .Activate
    Set APZelle = ActiveCell  'Wird z.B. in EinträgeLöschen gebraucht
    APZeile = ActiveCell.Row
    AnfangsZeile = APZeile
    If ActiveCell.Column <> APCDatum Or ActiveCell.Row > Cells(1, 3) Then
      A = MsgBox("Tastenkombination ''Strg+m'' hier wirkungslos" & Chr(10) & _
               "Zelle in Spalte 2 aktivieren?", vbOKCancel + vbQuestion, _
               "Monatswechsel im Arbeitsprotokoll")
      If A = vbCancel Then
        Exit Sub
      End If
      Cells(ActiveCell.Row, APCDatum).Activate
    End If
    ZNr = CStr(APZeile)
    Range("A" & ZNr & ":L" & ZNr).Select
    Selection.EntireRow.Insert
    A = MsgBox("An dieser Stelle die Kopfzeile Einfügen? ", vbYesNo + vbQuestion, _
                     "Monatswechsel im Arbeitsprotokoll")
    If A = vbYes Then
      Range("A2:L2").Select
      Selection.Copy
      Range("A" & ZNr & ":L" & ZNr).Select
      ActiveSheet.Paste
      Application.CutCopyMode = False
      For APZ = 3 To Cells(1, 3).Value + 50
        If Cells(APZ, APCgebucht) = "***" Then
          Cells(1, 3) = APZ
          Exit For
        End If
      Next APZ
      Cells(Cells(1, 3), APCDatum).Activate
    Else
      Selection.EntireRow.Delete
    End If
    Cells(AnfangsZeile, 2).Activate
    Exit Sub
  End With 'Worksheets("ArProt")
End With 'ActiveWindow
End Sub


