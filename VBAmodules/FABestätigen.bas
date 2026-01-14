Attribute VB_Name = "FABestätigen"
'*************************************************************************************
'* Makro FABestätigen *                                                              *
'*************************************************************************************
' Aufruf: Tastenkombination: Strg+f  von einer Kontenplanzeile aus, falls der Kontenplan
'einen Personenkontobereich mit Mitggliedskonten (Kontoart 10, Sammelkto 1000) und/oder
'Spenderkonten Kontoart11, SammelKto 2001) und/oder Aufwandsentschädigungskonten
'(KtoArt 12, SammelKonto 200) enthält.

'Erstellung einer Beitrags- oder Spendenbescheinigung für das im Kontenplan
'angewählte Mitglied (Kontoart 10 und Sammelkonto 1000 oder den angewählten
'Spender (Kontoart 11 und Sammelkont 2001) und bietet an, das Blatt auszudrucken.
'Hat das angewählte Konto das Sammelkonto 260, so ist es eine Auflistung der
'Aufwandsentschädigungen

'Enthält neben
'Sub FinAmtBestaetigen():
'Function BetragInWorten(Betrag As Variant) As String
'Function EinerZehner(ByVal Ziffern As String,Stellung As Integer)As String
'Letztere wird von BetragInWorten verwendet.

'in Formularprüfen -- Buchungsjahr in die Finanzamtbestätigungsvorlage eintragen ---
'    With Sheets("FABestVorl")
'      .Activate
'      Cells(10, 4).Value = BuchJahr  '********
'    End With
 Option Explicit
 Const TiT = "Beitragsbestätigung"      'MsgBox-Parameter, die
 Const HilfeDatei = "ExAccHilfe.hlp"     'für den ganzen Modul gelten
 
Sub FinAmtBestaetigen()
Attribute FinAmtBestaetigen.VB_ProcData.VB_Invoke_Func = "f\n14"

  Dim Bedingung, Schalter, HilfeKontxt, Antwort, W, DruckBereich As String
  Dim StartZelle As Range, AktZeile As Integer, AktSpalte As Integer
  Dim PersKtoNr As Integer, Beschreibung As String, IstAEKonto As Boolean
  Dim Bestätigung As Boolean, ZBetrLänge As Long
  Dim ZiffBetrag As Variant, EuroBetrag As String, WortBetrag As String, FName As String
  Dim Centbetrag As Integer
  Dim MStraße As String, MOrt As String, TagLetztZuw As String, ArtDerZuwendung As String
  Dim BitteWenden, Voraussetzung As String
  Dim BestName As String, BlattName As String, KontoArt As Long, AktJahr As Long
  Dim NamenLänge As Long, I As Long, B, D
  Dim EndText As String, Meldung2 As String
  Dim FABestVorlVorh As Boolean, SheetNr As Integer
  Dim NullBescheinigung As Boolean, NurSpenden As Boolean, Drucken As Boolean
  Dim AlteExAccVersion As String
  Dim ExAccVersion As String
  
  
With ActiveWindow
'FAB 1 ------------------------ Anfangssituation -----------------------------
  ExAccVersion = ThisWorkbook.Name
  MappenName = ActiveWorkbook.Name  'wird ggf. überschrieben
  Set StartZelle = ActiveCell
  AktZeile = ActiveCell.Row
  AktSpalte = ActiveCell.Column
  AlteExAccVersion = Sheets("Kontenplan").Cells(1, 12)
  Bestätigung = False
  NullBescheinigung = False
  MELDUNG = ""
'FAB2 ------------------- Aufrufbedingung Kontenplan ----------------------------
  With ActiveSheet
    If ActiveSheet.Name <> "Kontenplan" Then
      Bedingung = "nur von Blatt ''Kontenplan'' aus möglich" & Chr(10) & Chr(10) & _
                "Dorthin Wechseln?"
'      Schalter = vbYesNo '+ vbCritical + vbDefaultButton2 'Schaltflächen definieren.
      Antwort = MsgBox(Bedingung, vbYesNo, TiT) ', Hilfe, Ktxt)
      If Antwort = vbYes Then
        Worksheets("KontenPlan").Activate
      Else
        MELDUNG = MELDUNG & Chr(10) & _
        "Spendenbescheinigungs-Erstellung beendet."
        Exit Sub
      End If
    End If
  End With ' ActiveSheet
'FAB3 ---------------- Kontenplanstruktur, PersonenKtoBereich -------------------
    Call KontenplanStruktur
    If ABBRUCH = True Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Kontenplanstruktur defekt"
      GoTo BestBeenden
    End If
    If PersonKtoBereichVorhanden = False Then    'aus Kontenplanstruktur)
      MELDUNG = MELDUNG & Chr(10) & _
           "Die Applikation (Buchungsprojekt)" & Chr(10) & _
           MappenName & Chr(10) & _
           "enthält keine Personenkonten."
      ABBRUCH = True
      GoTo BestBeenden
    End If
'FAB4 ---------------- Beitrags / Spendenkonto? -----------------------------
PersKtoNr = Cells(AktZeile, 2)
    Call KtoKennDat(PersKtoNr)
    If AKtoSamlKto = 200 Then
      IstAEKonto = True
      GoTo DruckDialog
    End If
'FAB4 ---------------- Neue FA-Bestätigungsvorlage anfordern ----------------
    Sheets("Kontenplan").Activate
    If AktZeile = KPKZMitglieder Or Cells(KPKZMitglieder, 8) = "" Then
      FABestVorlVorh = False
      Antwort = MsgBox("Für die Spendenbescheinigungs-Erstellung muss " & Chr(10) & _
                       "eine gültige Vorlage neu erstellt werden.", vbOKCancel, TiT)
'FAB5 ----------------- Andere Kopfzeilen nicht erlaubt --------------------
    If Cells(AktZeile, 2) = "" And AktZeile <> KPKZMitglieder Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Unzulässige Zeile gewählt. Personenkonto wählen!"
      ABBRUCH = True
      GoTo BestBeenden
    End If
'FAB6 ---------------- alte FABestVorl ggf. löschen -------------------
      If Antwort = vbOK Then
        For Each W In Worksheets
          With W
            .Activate
            If Left(W.Name, 10) = "FABestVorl" Then
              Application.DisplayAlerts = False
              ActiveSheet.Delete
              Application.DisplayAlerts = True
            End If
          End With             'aktuellen Version
        Next W
'FAB7 -------- Vorlage von ExAcc in die Anwendung vor PosAnkF ------------------
        Windows(ExAccVersion).Activate
        Sheets("FABestVorl").Select
        Sheets("FABestVorl").Copy Before:=Workbooks(MappenName).Sheets("PosAnkF")
        With Workbooks(MappenName)
          Sheets("FABestVorl").Select
          ActiveSheet.Name = "FABestVorl"
          BuchJahr = Sheets("Kontenplan").Cells(1, 5)
          Sheets("FABestVorl").Cells(9, 3) = BuchJahr
          Sheets("Kontenplan").Cells(KPKZMitglieder, 8) = "FABestVorlVorhanden"
          FABestVorlVorh = True
          Sheets("FABestVorl").Activate
          MELDUNG = MELDUNG & Chr(10) & _
          "Diese Spendenbescheinigungs-Vorlage wurde aus " & ExAccVersion & " erneuert." _
          & Chr(10) & "Sie wird für die folgenden Bescheinigungen verwendet." _
          & Chr(10) & "Jetzt Prüfen und ggf. Jahr korrigieren (Jahr)"
          FABestVorlVorh = True
          Sheets("Kontenplan").Cells(KPKZMitglieder, 8) = "FABestVorlVorhanden"
          GoTo BestBeenden
        End With
      End If 'Antwort = vbOK
      If Antwort = vbCancel Then
        MELDUNG = MELDUNG & Chr(10) & _
        "Spendenbescheinigungs-Erstellung beendet."
        GoTo BestBeenden
      End If 'Antwort = vbCancel
    End If 'Vorlage neu
'FAB8 --------------------- Vorlage vorhanden ------------------------------
PersonenKonto:  'Mitglieds/Spenderkonto?
    Workbooks(MappenName).Sheets("Kontenplan").Activate
    PersKtoNr = Cells(AktZeile, 2)
    Call KtoKennDat(PersKtoNr)
'FAB9 -------------------- angwähltes Konto ein AE-Konto? ------------------
 '   If AKtoArt = 12 Then
 '     MELDUNG = MELDUNG & Chr(10) & _
 '       "Konto " & PersKtoNr & " ist kein Beitragskonto"
 '     GoTo BestBeenden
'FAB10 ---------angewähltes Konto ein Mitglieds- oder Spender-Konto? --------
    If AKtoArt = MitgliedKto Or _
       AKtoArt = SpenderKto Then
      GoTo EPrüf
    Else
      MELDUNG = MELDUNG & Chr(10) & _
                "Konto " & AKtoBeschr & Chr(10) & _
                "ist kein Beitrags- oder Spender-Konto!"
      ABBRUCH = True
      GoTo BestBeenden
    End If
'FAB11 --------- ist das Mitglieds/Spenderkonto überhaupt eingerichtet? ------------
EPrüf:
    If AKtoStatus <= KtoLeerMitÜbertrag Then
      B = MsgBox("Für das Konto " & Chr(10) & Chr(10) & _
                  AKtoBeschr & Chr(10) & Chr(10) & _
                  "sind keine Beiträge vorhanden." & Chr(10) & _
                  "Soll trotzdem eine Bestätigung erstellt werden?", _
                  vbYesNo, "Beitragsbestätigung")
      If B = vbYes Then
        WortBetrag = "in diesem Jahr keine Zuwendung"
        NullBescheinigung = True
      Else
        ABBRUCH = True
        GoTo BestBeenden
      End If
    End If
'FAB12 ---------------------- vergewissern: Adresse vollständig --------------------
    If AKtoStrasse = "" Or AKtoOrt = "" Then
      Meldung2 = "die Adresse des Kontos ''" & AKtoBeschr & "'' ist unvollständig"
      B = MsgBox(prompt:=Meldung2 & Chr(10) & _
                 "Soll trotzdem eine Bestätigung erstellt werden?", _
                 Buttons:=vbYesNo, Title:=TiT)
      If B = vbNo Then
        ABBRUCH = True
        MELDUNG = MELDUNG & Chr(10) & Meldung2 & Chr(10) & _
        "Adressangaben in Spalten I und J der Kontenplanzeile " & _
        AKtoKPZeil & " sanieren!"
        GoTo BestBeenden
      End If
    End If
'FAB13 --------------------------------Name, Straße, Ort ----------------------------------
    KontoArt = AKtoArt
    BlattName = AKtoBlatt    'für Bestätigungsblatt
    Beschreibung = AKtoBeschr
    NamenLänge = Len(Beschreibung)
    For I = 1 To NamenLänge Step 1    'Vornamen vor Familiennamen
      FName = Left(Beschreibung, I)
      If Mid(Beschreibung, I + 1, 2) = ", " Then
        FName = Right(Beschreibung, NamenLänge - I - 2) & " " & FName
        Exit For
      End If
      If I = NamenLänge Then
        FName = Beschreibung
      End If
    Next I
    MStraße = Cells(AktZeile, KPCStraße).Value
    MOrt = Cells(AktZeile, KPCOrt).Value
    If MStraße = "" Or IsNumeric(Left(MOrt, 5)) = False Then
      Range(Cells(AktZeile, KPCStraße), Cells(AktZeile, KPCOrt)).Select
      MELDUNG = "Ist die Adresse verwendbar ?"    ' Meldung definieren.
      Schalter = vbYesNo
      Antwort = MsgBox(MELDUNG, Schalter, TiT) ', Hilfe, Ktxt)
      If Antwort = vbNo Then
        B = MsgBox("" & Chr(10) & _
                   "Keine Bestätigung für ''" & FName & "'' erstellt", , TiT)
        Exit Sub
      End If
    End If
    BestName = AKtoBlatt
  If NullBescheinigung = True Then GoTo BetrIWorten
'FAB14 ----------- unterscheidet Vorlage zwischen Mitglied und Spender? ----------------
  With Worksheets("FABestVorl")
    .Activate
    If Cells(7, 5).Value = 0 Then
      NurSpenden = True   'Vorlage enthält keinen "wie Kirchensteuer"-Text
    Else
      NurSpenden = False
    End If
  End With
'FAB15 ----------------- Im Mitgliedskontoblatt Betrag ermitteln ----------------------
  With Worksheets(BestName) 'Mitglieds/Spender-Konto
    .Activate
    Cells(Cells(1, 1), 2).Activate  '***-Zelle
    If ActiveCell <> "***" Or _
       Left(ActiveCell.Offset(2, 3).Value, 10) <> "Kontostand" Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Kontoblatt ''" & AKtoBeschr & "'' hat fehlerhafte Struktur."
      ABBRUCH = True
      GoTo BestBeenden
    End If
    TagLetztZuw = ActiveCell.Offset(-1, 0).Value
    ZiffBetrag = CStr(ActiveCell.Offset(2, 6).Value) 'Einzutragen als Zahl
    ZBetrLänge = Len(ZiffBetrag)
  End With 'Worksheets Mitglieds/Spender-Konto
'FAB16 ----------- Darstellungsmöglichkeiten von ZiffBetrag berücksichtigen ------------
    Centbetrag = "00"
    If Left(Right(ZiffBetrag, 3), 1) = "," Then
      EuroBetrag = Left(ZiffBetrag, ZBetrLänge - 3)
      Centbetrag = Right(ZiffBetrag, 2)
      GoTo ZiffBetragAlsText
    End If
    If Left(Right(ZiffBetrag, 2), 1) = "," Then
      EuroBetrag = Left(ZiffBetrag, ZBetrLänge - 2)
      Centbetrag = Right(ZiffBetrag, 1) & "0"
      GoTo ZiffBetragAlsText
    End If
    If Left(Right(ZiffBetrag, 1), 1) = "," Then
      EuroBetrag = Left(ZiffBetrag, ZBetrLänge - 1)
      Centbetrag = "00"
      GoTo ZiffBetragAlsText
    End If
    EuroBetrag = ZiffBetrag
    Centbetrag = "00"
ZiffBetragAlsText:
    If Centbetrag = 0 Then
      ZiffBetrag = EuroBetrag & ",00"
    Else
      ZiffBetrag = EuroBetrag & "," & Centbetrag
    End If
'FAB17 ------------------------ Betrag in Worten ---------------------------------
BetrIWorten:
    WortBetrag = BetragInWorten(EuroBetrag)    'Function in diesem Modul
    If ABBRUCH = True Then GoTo BestBeenden
    AktJahr = Sheets("Kontenplan").Cells(1, 5)
    If KontoArt = MitgliedKto And NurSpenden = False Then 'Unterscheidung Beitrag/Spende
      ArtDerZuwendung = "Geldzuwendung (Mitgliedsbeiträge " & AktJahr & ")"
      Voraussetzung = "Vor"
    Else
      ArtDerZuwendung = "Geldzuwendung (Spenden " & AktJahr & ")"
      Voraussetzung = ""
    End If
'FAB18 ----------------------- Bestätigungs-Blatt erzeugen -----------------------------
  BestName = "Best" & AKtoBlatt
  For Each W In Worksheets
    If W.Name = BestName Then
      B = MsgBox("Soll ''" & BestName & "'' überschrieben werden?", vbYesNo, _
             "Spendenbestätigung ''" & BestName & "'' schon vorhanden:")
      If B = vbYes Then
        Application.DisplayAlerts = False
        W.Delete
        Application.DisplayAlerts = True
        Exit For
      Else
        Call MsgBox("Vorhandenes Blatt  ''" & BestName & "''" & Chr(10) & _
                "umbenennen und " & Chr(10) & _
                "Bestätigungserstellung von Kontenplan aus neu starten", 0, _
                BestName & "  Nicht überschreiben")
        Exit Sub
      End If
    End If
  Next W
'FAB19 ------------------------ Bestätigungs-Blatt erzeugen -----------------------------
  Worksheets("FABestVorl").Copy Before:=Sheets("FABestVorl")
  Worksheets("FABestVorl (2)").Name = BestName
'FAB18 ---------------------- Werte in Bestätigung eintragen ----------------------------
  With Worksheets(BestName)
    .Activate
    Cells(Cells(1, 4), Cells(1, 5)).Value = ArtDerZuwendung
    Cells(Cells(2, 4), Cells(2, 5)).Value = FName
    Cells(Cells(2, 4), Cells(2, 5)).Offset(1, 0).Value = MStraße
    Cells(Cells(2, 4), Cells(2, 5)).Offset(2, 0).Value = MOrt
    ZiffBetrag = "*****" & ZiffBetrag & ""
    Cells(Cells(3, 4), Cells(3, 5)).Value = ZiffBetrag
    Cells(Cells(4, 4), Cells(4, 5)).Value = WortBetrag & "  "
    Cells(Cells(5, 4), Cells(5, 5)).Value = TagLetztZuw & " " & AktJahr
    Cells(Cells(6, 4), Cells(6, 5)).Value = Date
    If Voraussetzung <> "Vor" And Cells(7, 4) <> 0 Then        'Löschen
      Cells(Cells(7, 4), Cells(7, 5)).Value = ""               'des
      Cells(Cells(7, 4), Cells(7, 5)).Offset(1, 0).Value = ""  'Voraussetzungs-
      Cells(Cells(7, 4), Cells(7, 5)).Offset(2, 0).Value = ""  'textes
      Cells(Cells(7, 4), Cells(7, 5)).Offset(3, 0).Value = ""
      Cells(Cells(8, 4), Cells(8, 5)).Value = ""
    End If
'FAB20----------------------------- Seite einrichten ---------------------------------
    DruckBereich = "A1:C" & Cells(9, 4).Value
    Cells(1, 1).Range(DruckBereich).Select
    ActiveSheet.PageSetup.PrintArea = Selection.Address
    Cells(1, 1).Range("A1").Select  'zur Positionierung
    BlattName = ActiveSheet.Name
'FAB21 ------------------------------ Druckdialog ----------------------------------
DruckDialog:
'    If AEKonto = True Then
'      B = MsgBox("" & Chr(10) & "AE-Liste für ''" & FName & "'' drucken?", _
'               vbYesNo, "AE-Liste vorhanden")
'    Else
      B = MsgBox("" & Chr(10) & "Zuwendungsbestätigung für ''" & FName & "'' drucken?", _
               vbYesNo, "Bestätigung erstellt")
'   End If
    If B = vbYes Then
      Application.PrintCommunication = True
      ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
            IgnorePrintAreas:=False
'       If AEKonto = True Then
      EndText = "Bestätigung für  ''" & FName & "''  ausgedruckt"
      Drucken = True
    End If
    If B = vbNo Then '
'      If AEKonto = True Then
'        EndText = "AE-Liste für  ''" & FName & "'' erstellt" & Chr(10) & _
'                  "Blattname ''" & BlattName & "''"
'      Else
        EndText = "Bestätigung für  ''" & FName & "''  erstellt" & Chr(10) & _
                  "Blattname ''" & BlattName & "''"
'      End If
      If AKtoStatus < KtoGanzLeer Then
        EndText = "Bestätigung für  ''" & FName & "''  erstellt" & Chr(10) & _
                "Kein Kontoblatt vorhanden."
      End If
      Drucken = False
    End If 'B =vbNo
    With Sheets("Kontenplan")
      .Activate
      Cells(StartZelle.Row, 8).Activate
      ActiveCell.Value = Date
      Selection.NumberFormat = "d/m/yy;@"
      If Drucken = False Then
        ActiveCell.Value = "( " & ActiveCell.Value & " )"
      End If
      Call MsgBox(EndText, 0, "Beitrags/Spenden-Bestätigung")
    End With 'Sheets("Kontenplan")
  End With 'Worksheets(BestName)
'FAB22 ------------------ Endemeldung ----------------------------
BestBeenden:
  With Workbooks(MappenName)
    Sheets("Kontenplan").Activate
  End With
  If ABBRUCH = False Then
    EndText = MELDUNG
  End If
  If ABBRUCH = True Then
    EndText = MELDUNG & Chr(10) & _
      "Bestätigung für  ''" & AKtoBeschr & "''  abgebrochen"
  End If
  If ABBRUCH = False Then
    EndText = MELDUNG & Chr(10) & _
      "Bestätigung für  ''" & AKtoBeschr & "''  erstellt"
  End If
    Call MsgBox(EndText, vbOKOnly, _
    "Beitrags/Spenden-Bestätigung")
  End With 'ActiveWindow
End Sub 'FinAmtBestätigen
'
'
Function BetragInWorten(Betrag As Variant) As String
'Wandelt den in Ziffern (auch als String) angegebenen ganzzahligen Betrag
'(maximal 999999) in das Zahlwort um.  Benutzt:
'Sub TextScan(Text As String, StartStelle As Long, EndString As String)
'out: TextStück (String), LängeTextstück, Lesezeiger (Long),
'     TextEndeErreicht (Boolean)

  Dim WBetrag As String, BetragLänge As Long, Stelle As Long, Länge As Long
  Dim BetZ As String, BetT As String, A
 
  BetragLänge = Len(Betrag)
  If BetragLänge > 6 Then
    MELDUNG = MELDUNG & Chr(10) & _
    "In Worten umzuwandelnder Betrag mehr als sechsstellig." & Chr(10) & _
         "Umwandlung abgebrochen."
    ABBRUCH = True
    WBetrag = "zu groß"
    GoTo Fertig
  End If
  WBetrag = ""
  If BetragLänge = 0 Then
    WBetrag = "Null"
    GoTo Fertig
  End If
  If BetragLänge = 1 Or BetragLänge = 2 Then
    GoTo Stelle2u1
  End If
  If BetragLänge = 3 Then
    GoTo Stelle3
  End If
  If BetragLänge = 4 Or BetragLänge = 5 Then
    GoTo Stelle5u4
  End If
  If BetragLänge = 6 Then
    GoTo Stelle6
  End If
Stelle6:
    BetZ = Left(Betrag, 1)                  '6. Stelle
    BetT = EinerZehner(BetZ, 6)
    WBetrag = BetT & "hundert"
Stelle5u4:
    BetZ = Right(Betrag, 5)
    If Len(BetZ) = 5 Then
      BetZ = Left(BetZ, 2)                '5. und 4. Stelle
      BetT = EinerZehner(BetZ, 5)
    End If
    If Len(BetZ) = 4 Then
      BetZ = Left(BetZ, 1)                '5. und 4. Stelle
      BetT = EinerZehner(BetZ, 4)
    End If
    WBetrag = WBetrag & BetT & "tausend"
Stelle3:
    BetZ = Right(Betrag, 3)
    BetZ = Left(BetZ, 1)                       '3. Stelle
    BetT = EinerZehner(BetZ, 3)
    If BetT <> "" Then
      WBetrag = WBetrag & BetT & "hundert"
    End If
Stelle2u1:
    BetZ = Right(Betrag, 2)                    '2. und 1. Stelle
    If Len(BetZ) = 2 Then
      BetT = EinerZehner(BetZ, 2)
    End If
    If Len(BetZ) = 1 Then
      BetZ = Left(BetZ, 1)                 '2. und 1. Stelle
      BetT = EinerZehner(BetZ, 1)
    End If
    WBetrag = WBetrag & BetT
Fertig:
BetragInWorten = WBetrag
End Function 'BetragInWorten
'
'
'
Function EinerZehner(ByVal Ziffern As String, Stellung As Integer) As String
'Liefert aus dem maximal 2 Zeichen langen String "Ziffern" je nach "Stellung" im
'Gesamtstring einen Wortausdruck einer der Zahlen 1...99. Die Information
'"Stellung" wird dazu verwendet, zwischen "eins" und "ein" zu unterscheiden.
'"Stellung" wird in der vom aufrufenden Programm in Text umzuwandelnden Gesamtzahl
'von rechts gezählt, also Einer, Zehner, Hunderter,... Werden zwei Ziffern übergeben,
'bezeichnet "Stellung" die Stelle der rechten Ziffer (von rechts nach links gezählt).
'
Dim LängeZ As Integer, ZiffZ As String, ZiffT As String

  LängeZ = Len(Ziffern)   'nur 1 oder 2 kommen vor
  ZiffT = ""
'-------------------- Fall "00" ----------------------------
  If Ziffern = "00" Then
    GoTo Fertig
  End If
'--------------- Fall "letzte Eins" ------------------------
  If Ziffern = "01" And Stellung <= 2 Then
    ZiffT = "eins"
    GoTo Fertig
  End If
'----------------- Fall "10 bis 19" ------------------------
  If LängeZ = 2 And Left(Ziffern, 1) = "1" Then
    If Right(Ziffern, 1) = "0" Then ZiffT = "zehn"
    If Right(Ziffern, 1) = "1" Then ZiffT = "elf"
    If Right(Ziffern, 1) = "2" Then ZiffT = "zwölf"
    If Right(Ziffern, 1) = "3" Then ZiffT = "dreizehn"
    If Right(Ziffern, 1) = "4" Then ZiffT = "vierzehn"
    If Right(Ziffern, 1) = "5" Then ZiffT = "fünfzehn"
    If Right(Ziffern, 1) = "6" Then ZiffT = "sechzehn"
    If Right(Ziffern, 1) = "7" Then ZiffT = "siebzehn"
    If Right(Ziffern, 1) = "8" Then ZiffT = "achzehn"
    If Right(Ziffern, 1) = "9" Then ZiffT = "neunzehn"
    GoTo Fertig
  End If
'----------------- Fall "x1 bis x9, x<>1" -------------------
  If LängeZ = 2 And Left(Ziffern, 1) <> "1" And Right(Ziffern, 1) <> "0" Or LängeZ = 1 Then
    If Right(Ziffern, 1) = "1" Then ZiffT = "ein"
    If Right(Ziffern, 1) = "2" Then ZiffT = "zwei"
    If Right(Ziffern, 1) = "3" Then ZiffT = "drei"
    If Right(Ziffern, 1) = "4" Then ZiffT = "vier"
    If Right(Ziffern, 1) = "5" Then ZiffT = "fünf"
    If Right(Ziffern, 1) = "6" Then ZiffT = "sechs"
    If Right(Ziffern, 1) = "7" Then ZiffT = "sieben"
    If Right(Ziffern, 1) = "8" Then ZiffT = "acht"
    If Right(Ziffern, 1) = "9" Then ZiffT = "neun"
  End If
'-------------------- Fall "nur Einer" ----------------------
  If LängeZ = 1 Or (LängeZ = 2 And Left(Ziffern, 1) = 0) Then
    GoTo Fertig
  End If
'-------------------- Fall "auch Zehner <> 1" ----------------------
  If ZiffT <> "" And LängeZ = 2 And Left(Ziffern, 1) <> "0" Then
    ZiffZ = ZiffT & "und"
  Else
    ZiffZ = ZiffT
  End If
'-------------------- Fall "Zehner <> 1" --------------------------
  If Left(Ziffern, 1) = "2" Then ZiffT = "zwanzig"
  If Left(Ziffern, 1) = "3" Then ZiffT = "dreißig"
  If Left(Ziffern, 1) = "4" Then ZiffT = "vierzig"
  If Left(Ziffern, 1) = "5" Then ZiffT = "fünfzig"
  If Left(Ziffern, 1) = "6" Then ZiffT = "sechzig"
  If Left(Ziffern, 1) = "7" Then ZiffT = "siebzig"
  If Left(Ziffern, 1) = "8" Then ZiffT = "achzig"
  If Left(Ziffern, 1) = "9" Then ZiffT = "neunzig"
  ZiffT = ZiffZ & ZiffT
Fertig:
EinerZehner = ZiffT
End Function 'EinerZehner





