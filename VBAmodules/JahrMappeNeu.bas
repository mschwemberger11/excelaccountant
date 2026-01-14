Attribute VB_Name = "JahrMappeNeu"
'*********************************************************************************
'* Makro JahrMappeNeu *    Aufruf Strg+j von Kontenplan aus                      *
'*********************************************************************************
Option Explicit
Option Base 1
'----------------------- Von Jahreswechsel gebraucht --------------------------------------
Public AltBuchjahr As Integer, BuchJahr As Integer, TransaktJahr As String
Public LinkerHeader As String, RechterHeader As String
Public AltJüngsTransakDatum As String, AlthöchsTAN As Integer
Public KntoVorlVorhanden As Boolean, FarbeBeibehalten As Boolean
Public ExAccVersion As String, MappenName As String '/angegeben?
Dim StartBlatt As String, StartZeile As Integer, StartSpalte As Integer
Dim KPZeile As Integer, ArProtZeile As Integer, LetzteArProtZeile As Integer
Dim BlattName As String, KtoNr As Long, Uebertrag As Long
Dim AltesBuchjahr As Integer, AltjahrLetzTan As Integer, AltjahrJüngstBuDatum As String
Dim LetzteTAN As String, AktPfad As String, TADatZahl As Integer
Dim MaxTADatZahl As Integer, AktTADatum, A, B, W
Dim TiT As String
Dim BestKtoAnfZeile As Integer, BestKtoEndZeile As Integer
Dim LetztesBuchungsDatum
'Public AktPfad As String, AktuellerPfad As String '=CurDir 'Warum ohne die beiden letzten Verästelungen
'Public JNStartmappe As String '= ActiveWorkbook.Name
'Public JNMappenName As String '= ActiveWorkbook.Name  'wird ggf. überschrieben
'Public JNStartBlatt As String '= ActiveSheet.Name
'Public JNStartRow As Integer '= ActiveCell.Row
'Public JNStartCol As Integer '= ActiveCell.Column


Sub Jahreswechsel()
Attribute Jahreswechsel.VB_ProcData.VB_Invoke_Func = "j\n14"
'Stellt aus einer VorJahres-Anwendermappe eine auf 1.Januar initialisierte
'aktuelle Projektmappe her.
'Vorgang:
'Im Dialog sicherstellen, dass die zu initialisierende Mappe eine Kopie der Vorjahresmappe
'mit dem für das neue Jahr gewünschten Namen ist (1JN bis 5JN).
'Im Dialog sicherstellen, dass der Buchungsstand des Vorjahres den gewünschten Buchungs-
'schnitt darstellt.
'Für alle Bestandskonten im Kontenplan (Kontoart 1, 6, 7) den jeweiligen, von Sub KtoKennDat
'aus dem noch vorhandenen Kontoblatt des Altjahres ermittelten
'Endsaldo in Spalte I (9) schreiben, ebenso den Vertragstext in Spalte J (10) (6JN bis 8JN).
'Buchjahr in E1 und Schalttag in F1 des Kontenplans schreiben (9JN bis 10JN).
'DruckBlattkopfzeilen im Kontenplan in I1 und K1 schreiben (11JN)
'Projekt-Farbkennzeichen im Kontenplan-Zelle G1 schreiben, Kontenplan färben (12JN bis 14JN)
'ArProt durch Vorlage aus ExAcc ersetzen, färben, mit Buchjahr versehen (16JN)
'In einer For-Each-Schleife alle vorhandenen Blätter einzeln behandeln: Kontenplan, ArProt,
'PosAnkK, -S,-B,-F belassen, KntoVorl mit BuchJahr VorJahrsende und Farbe versehen,
'Bestandskonten aus ExAcc-KntoVorl ersetzen mit allen Datenversehen, alle anderen Dateien
'entfernen (17JN bis 18JN).

  Dim FolgejahrMappe As Boolean, NeuesBuchungsProjekt As Boolean
  With ActiveWindow
BeginnDerNeuenMappe:
    Application.CutCopyMode = False
'1JN --------------- Namen der Mappe und des aktiven Blatts sichern ---------------------
    ExAccVersion = ThisWorkbook.Name
    MappenName = ActiveWorkbook.Name
    StartBlatt = ActiveSheet.Name
    StartZeile = ActiveCell.Row
    StartSpalte = ActiveCell.Column
    AktPfad = CurDir 'Pfad wird ohne die beiden letzten Verästelungen angegeben: Warum?
'    AktOrdner = Windows.Activefolder.Name  Objekt unterstützt diese Eigenschaft nicht
    MELDUNG = ""  'kumulierenden Meldungstext zurücksetzen
    ABBRUCH = False 'Globalen Abbruch zurücksetzen
    TiT = "Mappe neu einrichten"
'2JN -------------------- StartBlatt = Kontenplan? -----------------------
    If StartBlatt <> "Kontenplan" Then
      A = MsgBox(prompt:= _
          "Die Einrichtung einer Mappe für das Folgejahr" & Chr(10) & _
          "kann nur von der Zelle E1 im Kontenplan gestartet werden," & Chr(10) & _
          "für ein neues Buchungsprojekt nur von Zelle D1 eines Kontenplans." & _
          Chr(10) & Chr(10) & "Mappeneinrichtung nicht gestartet", _
          Buttons:=vbOKOnly, Title:=TiT)
      Exit Sub
    End If
'3JN -------------------- Mappe für Folgejahr oder neues Projekt? -----------------------
'    If StartZeile = 1 And StartSpalte = 4 Then
'      NeuesBuchungsProjekt = True
'      FolgejahrMappe = False
'      GoTo EinrichtungsBeginn
'      A = MsgBox(prompt:= _
'          TiT & " abgebrochen." & Chr(10) & _
'          "Nur Startzelle E1 (Folgejahr) erlaubt, D1 (neues Projekt) nicht implementiert.", _
'          Buttons:=vbOKOnly, Title:=TiT)
'      Exit Sub
'    End If
    If StartZeile = 1 And StartSpalte = 5 Then
'      NeuesBuchungsProjekt = False
'      FolgejahrMappe = True
      GoTo EinrichtungsBeginn
    End If
    A = MsgBox(prompt:= _
          TiT & " abgebrochen." & Chr(10) & _
          "Nur Startzelle D1 (neues Projekt) oder E1 (Folgejahr) erlaubt.", _
          Buttons:=vbOKOnly, Title:=TiT)
    Exit Sub
'4JN -------------------- Mappe für Folgejahr oder neues Projekt? -----------------------
EinrichtungsBeginn:
  Sheets("Kontenplan").Activate
  Call KontenplanStruktur
  If ABBRUCH = True Then
    A = MsgBox(prompt:= _
          "Dieser Projekt-Kontenplan hat Strukturfehler, die " & _
          "vor Einrichtung einer Mappe für ein Buchungsprojekt " & _
          "beseitigt sein müssen.", _
          Buttons:=vbOKOnly, Title:=TiT)
    Exit Sub
  End If
'  If FolgejahrMappe = True Then   'erst, wenn neues Buchungsprojekt implementiert
    A = MsgBox(prompt:= _
          "Die Einrichtung einer Mappe für ein Buchungsprojekt " & _
          "setzt voraus, dass sie eine Kopie der Anwendungsmappe " & _
          "des alten Jahres/Projektes ist und auf den " & _
          "Namen umbenannt ist, den sie im neuen Jahr/Projekt tragen soll. " & _
          "Ist " & Chr(10) & Chr(10) & _
          MappenName & Chr(10) & Chr(10) & " eine solche Kopie?", _
          Buttons:=vbYesNo, Title:=TiT)
    If A = vbNo Then GoTo MappeUmbenennen
    If A = vbYes Then GoTo MappenNamenBestätigen
MappeUmbenennen:
    B = MsgBox(prompt:= _
        "Der Vorgang " & TiT & " wird abgebrochen." & Chr(10) & _
        "Zum Erzeugen der gewünschten Mappe alle Mappen schließen und mit" & _
        " Dateiverwaltungsfunktionen (Explorer) Mappe " & MappenName & _
        " kopieren und umbenennen, wie sie im neuen Jahr/neuem Projekt" & _
        " heissen soll.  Dann diese und " & ExAccVersion & _
        " öffnen und die Initialisierung vom Kontenplan der richtig " & _
        "benamsten Mappe aus erneut mit ''strg+j'' starten.", _
        Buttons:=vbOKOnly, Title:=TiT)
      Exit Sub
'  End If 'FolgejahrMappe = True
'5JN ------------------------ Mappennamen bestätigen -----------------------
MappenNamenBestätigen:
    A = MsgBox(prompt:= _
      "Der Name der neuen Mappe ist" & Chr(10) & Chr(10) & _
        "        ''" & MappenName & "''" & Chr(10) & Chr(10) & _
        "Ist das so gewünscht?", _
        Buttons:=vbYesNo, Title:="Jahreswechsel: neue Mappe")
    If A = vbYes Then
      MELDUNG = MELDUNG & Chr(10) & _
      TiT & ": Namen " & MappenName & " akzeptiert."
    End If
    If A = vbNo Then
      GoTo MappeUmbenennen
    End If
'5JN -------------------- Mappennamen auf letztem Stand ----------------------
    ExAccVersion = ThisWorkbook.Name   'ExAcc-Mappen-Name
    MappenName = ActiveWorkbook.Name  'Name der Anwendungsmappe
'6JN ------------------ Stand der Buchungen im alten Jahr ---------------------
    Sheets("ArProt").Activate
    Call JuengstBuDatum    'letzten Buchungsstand gesondert feststellen, -> B1
    A = MsgBox(prompt:= _
          "Die Initiierung von " & MappenName & " für das neue Jahr " & _
          "geht aus von einem Buchungsstand " & Chr(10) & Chr(10) & _
          "      " & Sheets("ArProt").Cells(1, 2) & " " & Cells(1, 5) & Chr(10) & Chr(10) & _
          "Ist das der gewünschte Jahresschnitt?  d.h. Soll die Mappe für das Folgejahr" & Chr(10) & _
          "mit den sich daraus ergebenden Bestandskonten-Überträgen eingerichtet werden?", _
          Buttons:=vbYesNo, Title:="Jahresabschluss")
    If A = vbYes Then
      MELDUNG = MELDUNG & Chr(10) & _
      ": Buchungsschnitt " & Sheets("ArProt").Cells(1, 2) & " " & AltesBuchjahr & " akzeptiert."
    End If
    If A = vbNo Then
      B = MsgBox(prompt:= _
          "Alle Mappen schließen und in der Altjahrs-Mappe " & _
          "die noch fehlenden Buchungen durchführen bzw. " & Chr(10) & _
          "die zuviel gebuchten stornieren, " & Chr(10) & _
          "im neuen Ordner " & MappenName & " durch die verbesserte Altjahrs-Mappe " & _
          " ersetzen, ggf. Namen der ersetzten korrigieren und erneut " & _
          "mit korrigierter Mappe den Jahreswechsel starten." & Chr(10) & _
          "Abbruch zur Verbesserung des Buchungsschnitts.", _
          Buttons:=vbOKOnly, Title:="Jahresawechsel")
      Exit Sub
    End If

'9JN ----------------- Buchjahr im Kontenplan festlegen (-> E1) ----------------------
    With Sheets("Kontenplan")
      .Activate
      Cells(1, 5).Value = Cells(1, 5).Value + 1  'Zeile 3 ist verborgen
BuchJahrBestätigen:
      A = MsgBox(prompt:= _
                "Das neue Buchungsjahr ist " & Chr(10) & Chr(10) & _
                "     " & Cells(1, 5) & Chr(10) & Chr(10) & _
                "Ist das richtig?", Buttons:=vbYesNo, Title:=TiT)
      If A = vbYes Then
        BuchJahr = Cells(1, 5)
        MELDUNG = MELDUNG & Chr(10) & _
        ": Buchungsjahr " & Cells(1, 5) & " akzeptiert."
      End If
      If A = vbNo Then
BuchJahrEingabe:
        BuchJahr = Application.InputBox _
            (prompt:="Gewünschtes Buchungsjahr eingeben." & Chr(10) & _
            "4 Ziffern >2021 und <2100", _
            Title:="Buchungsjahr festlegen", Default:=Cells(1, 5).Value, _
            Type:=6, Left:=450, Top:=300)
            'Type:=2 = String, Type:=4 = Boolean,Type:=6 =
        Cells(1, 5).Value = BuchJahr   'Public-Variable
        GoTo BuchJahrBestätigen
        If Cells(1, 5).Value < 2017 Or BuchJahr > 2099 Then
          A = MsgBox(prompt:="Die Eingabe" & Cells(1, 5) & Chr(10) & _
                         "ist ungültig. Nochmal versuchen?", Buttons:=vbYesNo, _
                         Title:="Buchungsjahr festlegen")
          If A = vbYes Then
            GoTo BuchJahrEingabe
          Else
            ABBRUCH = True
            MELDUNG = MELDUNG & Chr(10) & _
            "Buchungsjahr " & BuchJahr & " festlegen nicht gelungen." & Chr(10) & _
            "Jahreswechsel abgebrochen."
            GoTo JWFertigMeldung
          End If
        End If 'Cells(1, 5).Value <= 2016
      End If 'A = vbNo (Buchjahr nicht richtig)
    End With 'Sheets("Kontenplan")
'10JN ---------- Schalttag errechnen und in Kontenplan eintragen -----------------
    Sheets("Kontenplan").Cells(1, 6) = SchaltTag 'Function in Modul Datumtextscan
'11JN ---------- Headertexte im Kontenplan bestätigen/eintragen -----------------
'                      (Seite einrichten erst beim Drucken)
KopfZeilenBestätigen:
    A = MsgBox(prompt:="Sind die Kopfzeilentexte für auszudruckende Blätter richtig?" _
         & Chr(10) & Chr(10) & "Text links oben:" & Chr(10) & _
         "  ''" & Sheets("Kontenplan").Cells(1, 9) & "''" & Chr(10) & Chr(10) & _
         "Text rechts oben:" & Chr(10) & _
         "  ''" & Sheets("Kontenplan").Cells(2, 9) & "''" & Chr(10), _
        Buttons:=vbYesNo, Title:="Kopfzeilentexte festlegen")
    If A = vbYes Then GoTo KopfzeileFestgelegt
    LinkerHeader = Application.InputBox _
          (prompt:="Gewünschten linken Kopfzeilentext eingeben." & Chr(10), _
          Title:="Kopfzeilentext festlegen", Default:=Sheets("Kontenplan").Cells(1, 9).Value, _
          Type:=6, Left:=450, Top:=300) 'Type:=2 = String, Type:=4 = Boolean,Type:=6 =
    RechterHeader = Application.InputBox _
          (prompt:="Gewünschten rechten Kopfzeilentext eingeben." & Chr(10), _
          Title:="Kopfzeilentext festlegen", Default:=Sheets("Kontenplan").Cells(1, 11).Value, _
          Type:=6, Left:=450, Top:=300) 'Type:=2 = String, Type:=4 = Boolean,Type:=6 =
    Sheets("Kontenplan").Cells(1, 9) = LinkerHeader
    Sheets("Kontenplan").Cells(1, 11) = RechterHeader
    GoTo KopfZeilenBestätigen
KopfzeileFestgelegt:
    MELDUNG = MELDUNG & Chr(10) & _
    "Blattüberschriften festgelegt"
'12JN ----------------- Farbe wählen: alte Farbe beibehalten und Kontenplan färben -------------------
'ProjektFarbeFestlegen:
    Dim FarbKz As String, Fa As Integer
    Sheets("Kontenplan").Activate
    Cells(1, 4).Activate
    Call AktBlattFärben(Cells(1, 7))  'Kennzeichen in G1 stimmt mit Farbe überein
    A = MsgBox(prompt:= _
        "Ist diese Farbe die gewünschte für die Blätter-Kopfzeilen des Buchungsvorhabens?", _
        Buttons:=vbYesNo, Title:="Farbe Festlegen")
    If A = vbYes Then
      FarbKz = Sheets("Kontenplan").Cells(1, 7)
      MELDUNG = MELDUNG & Chr(10) & _
      "Vorhandene Farbe wird beibehalten"
      FarbeBeibehalten = True
      GoTo KontenplanAufräumen
    End If
'13JN ---------------- Farbe wählen: Aus Farbpalette wählen -----------------------
    Windows(ExAccVersion).Activate
    Sheets("Farbpalette").Select
    FarbKz = ""
    For Fa = 4 To 18 Step 2
      Cells(Fa, 4).Activate
      B = MsgBox(prompt:= _
            "Farbe " & Cells(Fa, 3) & " verwenden?", _
            Buttons:=vbYesNo, Title:="Farbwahl")
      If B = vbYes Then
        FarbKz = "F" & Cells(ActiveCell.Row, 2) & ""
        MELDUNG = MELDUNG & Chr(10) & _
        "Farbe " & Cells(ActiveCell.Row, 3) & " gewählt"
        Exit For
      End If
    Next Fa
    If FarbKz = "" Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Keine Farbe gewählt"
    End If
    Windows(MappenName).Activate
    Sheets("Kontenplan").Cells(1, 7) = FarbKz
'14JN ---------------- Kontenplan färben -------------------------------------
KontenplanFaerben:
    Workbooks(MappenName).Sheets("Kontenplan").Activate
    Call KontenplanStruktur
    Call AktBlattFärben(FarbKz)

'15JN --------------------------- Kontenplan aufräumen --------------------------
KontenplanAufräumen:
    Call KontenplanStruktur
    With Sheets("Kontenplan")
      If PersonKtoBereichVorhanden = True Then
        Range("H" & KPKZMitglieder & ":H" & KPKZEnde).Select
        Selection.ClearContents
      End If
      Cells(1, 2) = ""               'Sprungmarke löschen
      Cells(1, 1) = 1                'Kontenplanversion
      Cells(3, 5) = ""               'Vorjahr
      MELDUNG = MELDUNG & Chr(10) & _
      "Kontenplan: Versionsnummer zurückgesetzt"
      Range("A1").Select    'nur aus optischem Grund
   End With 'Sheets("Kontenplan")
'8JN ----------- BestandsKonten-Endstände im Alten Kontenplan sichern -----------
JahresAbschluss:
   Sheets("Kontenplan").Activate
   Call BestandsKontenEndStändeSichern   'Sub im Modul Jahreswechsel
'16JN ------------- ArProt durch Vorlage von ExAcc ersetzen -----------------------
   Windows(ExAccVersion).Activate
    Sheets("ArProtVorl").Select
    Sheets("ArProtVorl").Copy _
           after:=Workbooks(MappenName).Sheets("Kontenplan")
    Windows(MappenName).Activate
    ActiveWorkbook.Sheets("ArProt").Activate
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
    ActiveWorkbook.Sheets("ArProtVorl").Activate
    ActiveSheet.Name = "ArProt"
    ActiveWorkbook.ActiveSheet.Cells(1, 5) = BuchJahr
'17 ------------------- ArProt einfärben  Cells(4, 8) = "***"
     Call AktBlattFärben(FarbKz)
'    End If
    MELDUNG = MELDUNG & Chr(10) & _
    "ArProt initialisiert"
    Range("A1").Select   'nur aus optischem Grund
'17JN --------------- KntoVorl chronisieren -----------------------------
 Dim VorJahr As Integer
    VorJahr = BuchJahr - 1
    FarbKz = Sheets("Kontenplan").Cells(1, 7)
    Sheets("KntoVorl").Activate
    Cells(1, 8) = BuchJahr
    Cells(4, 3) = "31.12." & VorJahr
    Call AktBlattFärben(FarbKz)
    MELDUNG = MELDUNG & Chr(10) & _
    "KntoVorl initialisiert"
'17JN ======= Alle Blätter löschen/belassen/initialisieren ======
    Dim BlattKtoNr As Integer, BlattNamen As String, D3StZeil As Integer
    Dim BestKtoZei As Integer, JESaldo As Double, BlattUnten As Integer
    Dim BKUEBERTRAG As String
    BKUEBERTRAG = ""
    Sheets("Kontenplan").Activate
    For Each W In Worksheets
 '    With W
'       .Activate
       If W.Visible = xlSheetVeryHidden Then GoTo NaechstW
       '--------------- Kontenplan belassen ------------------------
       If W.Name = "Kontenplan" Then GoTo NaechstW
       '--------------- ArProt ist schon initialisiert --------------
       If W.Name = "ArProt" Then GoTo NaechstW
       '--------KntoVorl ist schon chronisiert, ggf.gefärbt----------
       If W.Name = "KntoVorl" Then GoTo NaechstW
       '------------- die Positionsanker belassen -------------------
       If W.Name = "PosAnkK" Then GoTo NaechstW
       If W.Name = "PosAnkS" Then GoTo NaechstW
       If W.Name = "PosAnkB" Then GoTo NaechstW
       If W.Name = "PosAnkF" Then GoTo NaechstW
       ' -----Bestandskonten initialisieren, andere löschen ---------------------
       '      Konten werden an Text "Kontoart" in Zelle B1 erkannt
       BlattNamen = W.Name
       Application.ScreenUpdating = True
       Sheets(BlattNamen).Activate
       '------- Nichtkontoblätter löschen ------------------------
       If Cells(1, 2) <> "Kontoart" Then
LoeschBlatt:
         Application.ScreenUpdating = True
         Sheets(W.Name).Activate
         Application.DisplayAlerts = False
         ActiveSheet.Delete
         Application.DisplayAlerts = True
         GoTo NaechstW
       End If
       '-------- Konten unterscheiden --------------------
       If Cells(1, 2) = "Kontoart" And BlattNamen <> "KntoVorl" Then
         BlattKtoNr = Cells(1, 4)        'aus Konto-Kopfzeile
         Call KtoKennDat(BlattKtoNr)
       '-------- Konten außer Bestandskonten löschen -------------
         If Not (AKtoArt = 1 Or AKtoArt = 6 Or AKtoArt = 7) Then
LoeschKontoBlatt:
           Application.DisplayAlerts = False
           Sheets(W.Name).Delete
           Application.DisplayAlerts = True
           Sheets("Kontenplan").Cells(AKtoKPZeil, 1) = ""   'E löschen
           If AKtoArt = 10 Or AKtoArt = 11 Then
             Sheets("Kontenplan").Cells(AKtoKPZeil, 8) = "" 'FA-Besch.-Kz löschen
           End If
           GoTo NaechstW
         End If 'Not(AktoArt=1/6/7)
       '--------BestandsKto Saldo sichern  -------------------
         If AKtoArt = 1 Or AKtoArt = 6 Or AKtoArt = 7 Then
           D3StZeil = Cells(1, 1)
           If Cells(Cells(1, 1), 2) = "***" And _
             Cells(Cells(1, 1) + 2, 5) = "Kontostand" Then 'Strukturkontrolle
             JESaldo = Cells(Cells(1, 1) + 2, 6)
           Else
             MELDUNG = MELDUNG & Chr(10) & _
             "Strukturfehler im Konto " & AKtoBeschr & Chr(10) & _
             "Jahresanfangsstand manuell in Zelle F4 eintragen!"
             JESaldo = 0  'Initialisierung deshalb nicht abgebrochen
           End If
       '------- Bestandskonto inititialisieren ------------------
           Sheets(BlattNamen).Cells(1, 8) = Sheets("KntoVorl").Cells(1, 8)
           BlattUnten = Sheets(BlattNamen).Cells(1, 1) + 2
       '------ Bei unverändertem Kto nur Übertragsdatum austauschen -------
           If BlattUnten <= 10 Then
             Sheets(BlattNamen).Cells(4, 3) = Sheets("KntoVorl").Cells(4, 3)
             GoTo NaechstW
           End If
       '------- Januar-mit KntoVorl überschreiben, Rest löschen ------------------
           Sheets("KntoVorl").Select
           Range("A4:I9").Select
           Selection.Copy
           Sheets(BlattNamen).Activate
           Range("A4:I9").Select
           ActiveSheet.Paste
          ' JESaldo = Cells(BlattUnten, 6)
           Sheets(BlattNamen).Select
           ActiveSheet.Range("A10:I" & BlattUnten).Select
           Application.CutCopyMode = False
           Selection.Delete shift:=xlUp
           ActiveSheet.Cells(4, 6) = JESaldo
           GoTo NaechstW
         End If 'AktoArt = 1/6/7
       End If 'Cells(1, 2) = "Kontoart"
       '------------ Alle nicht aufgeführten löschen ---------------
       Application.DisplayAlerts = False
       Sheets(W.Name).Delete
       Application.DisplayAlerts = True
NaechstW:
    Next W
  ' Sheets(ActiveSheet.Index + 1).Activate
   
   '---------------- KP bezüglich E aufräumen ---------------------
   '   für Fälle, wenn E bei einem nicht vorhandenen Blatt stand
Dim KoPlaZei As Integer
   Sheets("Kontenplan").Activate
   Call KontenplanStruktur
   For KoPlaZei = 5 To KPKZEnde
     If Cells(KoPlaZei, 2) = "" Then
       GoTo NaeKoPlaZeil
     End If
     Call KtoKennDat(Cells(KoPlaZei, 2))
     If Not (AKtoArt = 1 Or AKtoArt = 6 Or AKtoArt = 7) Then
       Cells(KoPlaZei, 1) = ""
     End If
NaeKoPlaZeil:
   Next KoPlaZei
'? End With 'Workbook(MappenName)
'19JN -------------------------- Fertigmeldung --------------------------------------
Dim Hinweis As String
JWFertigMeldung:
  Hinweis = "Hinweis 1:" & Chr(10) & _
    "Bei Bedarf kann der Jahresübertrag eines Bestandskontos in Zelle F4 des Kontenblatts" & _
    " manuell geändert werden." & Chr(10) & "(Andere Änderungen als in F4 sind verboten!)" & Chr(10) & _
    Chr(10) & "Hinweis 2:" & Chr(10) & _
    "Jetzt ist eine günstige Situation für gewünschte Kontenplanänderungen."
  If ABBRUCH = True Then
    MELDUNG = MELDUNG & Chr(10) & _
    "Die Initialisierung der Mappe ''" & MappenName & "'' wurde abgebrochen." & Chr(10) & _
    "Diese Mappe für einen erneuten Versuch nicht mehr verwendbar."
    GoTo FMDrucken
  End If
  MELDUNG = MELDUNG & Chr(10) & Chr(10) & _
    "Die Mappe  ''" & MappenName & "'' für das Jahr " & BuchJahr & " ist" _
    & Chr(10) & "mit den genannten Bestandskonten-Überträgen initialisiert."
    
FMDrucken:
    Call MsgBox(BKUEBERTRAG, vbOKOnly, TiT)
    Call MsgBox(MELDUNG, vbOKOnly, TiT)
    Call MsgBox(Hinweis, vbOKOnly, TiT)
  End With 'ActiveWindow
  End Sub  '========================= Ende Jahreswechsel ===============================
'13JN ---------- Bestandskonten-Überträge prüfen/korrigieren? ---------------
Sub UebertraegePrüfen()
   A = MsgBox(prompt:="Die Jahresüberträge der Bestandskonten" & Chr(10) & _
                      "prüfen? (mit Korrekturmöglichkeit)", Buttons:=vbYesNo, _
                      Title:="Übertrag")
   If A = vbYes Then
     Call KontenplanStruktur
     If ABBRUCH = True Then
       Call MsgBox(MELDUNG, vbOKOnly, TiT & " Jahreswechsel")
       Exit Sub
     End If
     With Sheets("Kontenplan")
       .Activate
       For KPZeile = KPKZBestand + 1 To KPKZBestand + SLZZBestand
 '        Call UebertragsFrage(KPZeile)
       Next KPZeile
       If KPKZFonds <> 0 Then
         For KPZeile = KPKZFonds + 1 To KPKZFonds + SLZZFonds
'           Call UebertragsFrage(KPZeile)
         Next KPZeile
       End If
       If KPKZVermögen <> 0 Then
         For KPZeile = KPKZVermögen + 1 To KPKZVermögen + SLZZVermögen
'           Call UebertragsFrage(KPZeile)
         Next KPZeile
       End If
     End With 'Sheets ("Kontenplan")
     MELDUNG = MELDUNG & Chr(10) & _
     "Mappe " & MappenName & ": Bestandskontenüberträge geprüft"
   End If 'A = vbYes
End Sub

Sub UebertragsFrage(KpZ As Integer)
  Dim UebertragBox As Variant
  If Sheets("Kontenplan").Cells(KpZ, 1) = "E" Then
    KtoNr = Sheets("Kontenplan").Cells(KpZ, 2)
    Call KtoKennDat(KtoNr)
    Sheets(AKtoBlatt).Activate
    Uebertrag = Sheets(AKtoBlatt).Cells(4, 6).Value
    If Uebertrag < 0 Then              'Vermeidung, dass durch das
      UebertragBox = "'" & Uebertrag   'Minuszeichen ExCel den
    Else                               'Uebertrag als Formel auffasst
      UebertragBox = Uebertrag
    End If
    Uebertrag = Application.InputBox(prompt:= _
         "Ist für das Konto ''" & KtoNr & "'' dieser Übertrag" & Chr(10) & _
         "vom Vorjahr richtig? Wenn nein, hier gewünschten" & Chr(10) & _
         "Betrag eingeben. Zahl mit 2 Stellen hinter dem Komma.", _
         Title:="Übertrag", Default:=UebertragBox, _
         Type:=6, Left:=450, Top:=300)
         'Type:=2 = String, Type:=4 = Boolean,Type:=6
    Sheets(AKtoBlatt).Cells(4, 6) = UebertragBox
  End If
End Sub 'UebertragsFrage

'14JN ---------------------------- Fertigmeldung -------------------------------

Sub BestandsKontenEndStändeSichern()                           '17.11.2017
  'Im vorhandenen Buchungsprojekt werden von den Bestandskonten die dem
  'Buchungsstand entsprechenden Endsalden in die Spalte I (9) des
  'Kontenplans gesichert; ebenso die Vertragstexte in Spalte J (10).
  Call KontenplanStruktur
  If ABBRUCH = True Then Exit Sub
  If KPKZBestand <> 0 Then
    BestKtoAnfZeile = KPKZBestand
    BestKtoEndZeile = KPKZBestand + SLZZBestand
    Call Bestandbereichschleife(BestKtoAnfZeile, BestKtoEndZeile)
  End If
  If KPKZFonds <> 0 Then
    BestKtoAnfZeile = KPKZFonds
    BestKtoEndZeile = KPKZFonds + SLZZFonds
    Call Bestandbereichschleife(BestKtoAnfZeile, BestKtoEndZeile)
  End If
  If KPKZVermögen <> 0 Then
    BestKtoAnfZeile = KPKZVermögen
    BestKtoEndZeile = KPKZVermögen + SLZZVermögen
    Call Bestandbereichschleife(BestKtoAnfZeile, BestKtoEndZeile)
  End If
End Sub 'BestandsKontenEndStändeSichern
  Sub Bestandbereichschleife(BestKtoAnfZeile, BestKtoEndZeile)
    Sheets("Kontenplan").Activate
    With Sheets("Kontenplan")
      For KPZeile = BestKtoAnfZeile To BestKtoEndZeile
        Cells(KPZeile, 2).Activate
        If ActiveCell = "" Or Cells(KPZeile, 1) <> "E" Then GoTo NäKPZeile
        Call KtoKennDat(Cells(KPZeile, 2))
        If AKtoArt = 1 Or AKtoArt = 6 Or AKtoArt = 7 Then
          If AKtoStatus >= KtoGanzLeer Then
            Cells(KPZeile, 10) = AktoVertragsText   'aus Blatt ermittelt, normaler-
            Cells(KPZeile, 9) = AktoEndSaldo        'weise nicht im Kontenplan
            MELDUNG = MELDUNG & Chr(10) & _
            "Bestands-Endsaldo ''" & AKtoBlatt & "'' gesichert"
          End If
        End If
NäKPZeile:
      Next KPZeile
    End With
  End Sub 'Betandbereichschleife   End With
   

    
Function ArProtEndZeile() 'Suche ArProtEnd-Tabellenzeile über TA-Nr.
  Dim APZeil As Integer, LetztAPZ As Integer, LetzteTAN As Integer
  Dim StartBlatt As String, StartZeile As Integer, StartSpalte As Integer
  Dim JüBuDatum As Integer
  'out: LetztesBuchungsDatum
  StartBlatt = ActiveSheet.Name
  StartZeile = ActiveCell.Row
  StartSpalte = ActiveCell.Column
  With Worksheets("ArProt")
    .Activate
    LetzteTAN = Cells(1, 1)
    If LetzteTAN <= 3 Then              'Vor den ersten drei Buchungszeilen
       ArProtEndZeile = Cells(1, 1) + 2   'keine Monatsüberschriftzeilen
      GoTo APEZEnde                       'erwartet.
    End If
    For APZeil = LetzteTAN + 2 To LetzteTAN + 100 'Differenz < 100 vorausgesetzt
      If Cells(APZeil, APCTANr) = "TA" Then
        APZeil = APZeil + 1
      End If
      If Cells(APZeil, APCTANr) = LetzteTAN Then
        Exit For
      End If
    Next APZeil
    If Cells(APZeil, APCDatum) = "" And IstDatum(Cells(APZeil - 1, APCDatum)) = True Then
      LetztAPZ = APZeil - 1
    End If
    If Cells(APZeil - 1, APCTANr) = "TA" And APZeil > 3 _
       And IstDatum(Cells(APZeil - 2, APCDatum)) = True Then
      LetztAPZ = APZeil - 2
    End If
    LetztesBuchungsDatum = Cells(1, 2) '.Activate
    ArProtEndZeile = Cells(1, 3)
APEZEnde:
  End With 'Worksheets("ArProt")
  With Sheets(StartBlatt)
    .Activate
    Cells(StartZeile, StartSpalte).Activate
  End With
End Function 'ArProtEndZeile     'LetzteArProtZeile

Sub JuengstBuDatum()  'Function DatumTZ(DText As String)
  'Durchsucht Spalte 2 von ArProt und schreibt das jüngste Datum in Zelle B1
  Dim JBuDat As String, APDatum As String, Zeile As Integer
  
  Sheets("ArProt").Cells(3, 2).Activate
  APDatum = ActiveCell
  JBuDat = ActiveCell
  For Zeile = 3 To Cells(1, 1) + 10
    Cells(Zeile, 2).Activate
    APDatum = ActiveCell
    If IstDatum(APDatum) = True Then
      If DatumTZ(APDatum) >= DatumTZ(JBuDat) Then
        JBuDat = APDatum
      End If
      Cells(1, 2) = JBuDat
    End If
NaechsteZeile:
  Next Zeile
  Cells(1, 2).Activate
End Sub 'JuengstBuDatum

Function IstDatum(DText) As Boolean  'ExAcc-interne Datumsform TT.MMM ohne Jahr
  Dim DT As String
  If Len(ActiveCell) > 6 Or Len(ActiveCell) < 5 Or ActiveCell = "" Then GoTo NoText
    DT = Right(ActiveCell, 4)
    If DT = ".Jan" Or DT = ".Feb" Or DT = ".Mrz" Or DT = ".Apr" _
      Or DT = ".Mai" Or DT = ".Jun" Or DT = ".Jul" Or DT = ".Aug" _
      Or DT = ".Sep" Or DT = ".Okt" Or DT = ".Nov" Or DT = ".Dez" _
      Then
      IstDatum = True
      GoTo Ex
    End If
NoText:
  IstDatum = False
Ex:
End Function




Sub AktBlattFärben(FarbNummer As String) 'Gemäß Blatt-Typ und FarbNummer
'----------------------- Blatt-Typ "Kontenplan"-------------------------
'Färbt das aktive Blatt von einem der Blattypen
'Kontenplan, ArProt, KontoVorl, Konto, SaliVorlVorl, Sali2VorlVorl, BeriVorlVorl,
'KoTaVorl ein mit der durch FarbNummer ("F1"/"F2"/"F3"/"F4"/"F5"/"F6"/"F7"/"F8")
'gegebenen Farbe.
Dim ZS1 As Integer, ZS2 As Integer, ZS3 As Integer

  If ActiveSheet.Name = "Kontenplan" Then
    ZS1 = Cells(1, 3) + 2
    Range("A1:K2,A3:A" & ZS1).Select
    If FarbNummer = "F1" Then GoTo F1
    If FarbNummer = "F2" Then GoTo F2
    If FarbNummer = "F3" Then GoTo F3
    If FarbNummer = "F4" Then GoTo F4
    If FarbNummer = "F5" Then GoTo F5
    If FarbNummer = "F6" Then GoTo F6
    If FarbNummer = "F7" Then GoTo F7
    If FarbNummer = "F8" Then GoTo F8
    If FarbNummer = "" Then GoTo OhneFarbe
  End If
'----------------------- Blatt-Typ "ArProt"-------------------------
  If ActiveSheet.Name = "ArProt" Then
      ZS1 = Cells(1, 3) + 3
      Range("A1:K2,A3:A" & ZS1 & ",J3:J" & ZS1).Select
    If FarbNummer = "F1" Then GoTo F1
    If FarbNummer = "F2" Then GoTo F2
    If FarbNummer = "F3" Then GoTo F3
    If FarbNummer = "F4" Then GoTo F4
    If FarbNummer = "F5" Then GoTo F5
    If FarbNummer = "F6" Then GoTo F6
    If FarbNummer = "F7" Then GoTo F7
    If FarbNummer = "F8" Then GoTo F8
    If FarbNummer = "" Then GoTo OhneFarbe
  End If
'----------------------- Blatt-Typ "KntoVorl"-------------------------
  If ActiveSheet.Name = "KntoVorl" Then
    Range("A1:I2,A3:A10,I3:I10,C4:E4,G4:H4,C7:H7,C8:E8,G8:H8").Select
    If FarbNummer = "F1" Then GoTo F1
    If FarbNummer = "F2" Then GoTo F2
    If FarbNummer = "F3" Then GoTo F3
    If FarbNummer = "F4" Then GoTo F4
    If FarbNummer = "F5" Then GoTo F5
    If FarbNummer = "F6" Then GoTo F6
    If FarbNummer = "F7" Then GoTo F7
    If FarbNummer = "F8" Then GoTo F8
    If FarbNummer = "" Then GoTo OhneFarbe
  End If
'----------------------- Blatt-Typ Konto-------------------------
  If ActiveSheet.Name <> "KntoVorl" And _
     ActiveSheet.Cells(1, 2) = "Kontoart" Then
    ZS1 = Cells(1, 1) + 4
    ZS2 = Cells(1, 1) + 1
    ZS3 = Cells(1, 1) + 2
    Range("A1:I2,A3:A" & ZS1 & ",I3:I" & ZS1 & "," & _
    "C4:E4,G4:H4,C" & ZS2 & ":H" & ZS2 & "," & _
    "C" & ZS3 & ":E" & ZS3 & ",G" & ZS3 & _
    ":H" & ZS3 & "").Select  'nur den letzten Saldoblock
    If FarbNummer = "F1" Then GoTo F1
    If FarbNummer = "F2" Then GoTo F2
    If FarbNummer = "F3" Then GoTo F3
    If FarbNummer = "F4" Then GoTo F4
    If FarbNummer = "F5" Then GoTo F5
    If FarbNummer = "F6" Then GoTo F6
    If FarbNummer = "F7" Then GoTo F7
    If FarbNummer = "F8" Then GoTo F8
    If FarbNummer = "" Then GoTo OhneFarbe
  End If
  '----------------------- Blatt-Typ "SaliVorlVorl"-------------------------
  If ActiveSheet.Name = "SaLiVorl" Then
'    Range("A1:M60").Select      'nirgendwo Farben
'     With Selection.Interior
'        .Pattern = xlSolid
'        .PatternColorIndex = xlAutomatic
'        .ThemeColor = xlThemeColorDark2       'Braun
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
'    End With
'    With Selection.Interior
'        .Pattern = xlNone
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
'      End With
    Range("A1:M1,A2,A3:M3,A6:M6,A7:A8,A9:D9,A11:M11,A12:A13,A14:D14," & _
    "A16:M16,A17:A18,A19:D19,A21:M21,A22:A23,A24:D24,A26:C27,E26:F27," & _
    "H26:I27,K26:L27,A29:M29,A30:A31,A32:D32,A34:M34,A35:A37,B37:D37," & _
    "A39:C40,E39:F40,H29:I40,K39:L40,A42:M42,A43:A44,A45:D45").Select
    If FarbNummer = "F1" Then GoTo F1
    If FarbNummer = "F2" Then GoTo F2
    If FarbNummer = "F3" Then GoTo F3
    If FarbNummer = "F4" Then GoTo F4
    If FarbNummer = "F5" Then GoTo F5
    If FarbNummer = "F6" Then GoTo F6
    If FarbNummer = "F7" Then GoTo F7
    If FarbNummer = "F8" Then GoTo F8
    If FarbNummer = "" Then GoTo OhneFarbe
  End If
'----------------------- Blatt-Typ "BeriVorlVorl" -------------------------
  If ActiveSheet.Name = "BeriVorl" Then 'And ZweiteHälfte = False Then
    Range("A2:J2,A3:B3,I3:J3,A4:J4," & _
        "B6:J6,B9,D9:F9,I9:J9," & _
        "B11:J11,B14,D14:F14,I14:J14," & _
        "B16:J16,B19,D19:F19,I19:J19," & _
        "B21:E21,G21:J21,B24,D24:E24,G24,I24:J24," & _
        "B26:E26,G26:J26,B29,D29:E29,G29,I29:J29," & _
        "B32:B35,C32:J33").Select
    If FarbNummer = "F1" Then GoTo F1
    If FarbNummer = "F2" Then GoTo F2
    If FarbNummer = "F3" Then GoTo F3
    If FarbNummer = "F4" Then GoTo F4
    If FarbNummer = "F5" Then GoTo F5
    If FarbNummer = "F6" Then GoTo F6
    If FarbNummer = "F7" Then GoTo F7
    If FarbNummer = "F8" Then GoTo F8
    If FarbNummer = "" Then GoTo OhneFarbe
   End If
 If ActiveSheet.Cells(1, 4) = "KONTENSTANDTABELLE" Then
   Range("A1:R2,A3:A" & Cells(1, 3) & "").Select
   If FarbNummer = "F1" Then GoTo F1
   If FarbNummer = "F2" Then GoTo F2
   If FarbNummer = "F3" Then GoTo F3
   If FarbNummer = "F4" Then GoTo F4
   If FarbNummer = "F5" Then GoTo F5
   If FarbNummer = "F6" Then GoTo F6
   If FarbNummer = "F7" Then GoTo F7
   If FarbNummer = "F8" Then GoTo F8
   If FarbNummer = "" Then GoTo OhneFarbe
 End If
  
    '--------- Wahl aus der Farbpalette ---------------
F1:
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1       'Grau
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
    GoTo EndeAktBlattFärben
F2:
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2       'Braun
        .TintAndShade = -9.99786370433668E-02
        .PatternTintAndShade = 0
    End With
    GoTo EndeAktBlattFärben
F3:
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2      'Blau
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    GoTo EndeAktBlattFärben
F4:
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255                          'Rot
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    GoTo EndeAktBlattFärben
F5:
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 49407                        'Orange
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    GoTo EndeAktBlattFärben
F6:
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6     'Rosa
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    GoTo EndeAktBlattFärben
F7:
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535                        'Gelb
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    GoTo EndeAktBlattFärben
F8:
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274                      'Grün
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    GoTo EndeAktBlattFärben
OhneFarbe:
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
EndeAktBlattFärben:
  Cells(4, 3).Activate
End Sub 'AktBlattFärben
'-----------Ende Modul JahrmappeNeu ---------------------------


