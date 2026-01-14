Attribute VB_Name = "KontPlanPflege"

'****************************************************************************************
'* Modul KontPlanPflege *  shortcut Strg+k von Kontenplan                               *
'****************************************************************************************
' Das Hauptprogramm unterstützt das Einfügen, Ändern, Löschen und Verschieben von
' Kontozeilen im Kontenplan und das Leeren eines Kontoblatts mittels betreffender
' Stornobuchungen, die im Arbeitsprotokoll als durchgestrichen ersichtlich sind.
' Weitere, von anderen Programmen benutzte Routinen zum Konto- und Zeit-Management
'Enthält:
' Die Globalen Variablen und Konstanten, die mit dem Kontenplan zusammenhängen.
' Sub KONTPLANBEARB() als Hauptroutine für die , die durch Tastenkombination Str+k aufgerufen wird
' Sub KontenplanStruktur() Ermittlung der Strukturparameter des Kontenplans
' Sub KontPlaZeileEinfügen(ActCel As Range)
' Function LöschKtoBlatt(KoPlaZelle As Range) As Boolean
' Function Kontoblatt(ByVal KtoKz As Variant) As String
' Sub KontoNichtDefiniert(KtoSoHa, KtoKennung, KPZeile)
' Function Eingerichtet(ByVal KtoKz As Variant) As Boolean
' Sub KtoKennDat
' Sub KontblattErstelln
' Sub SucheKonto(SuchText As String)
' Sub BuchJahrFestlegen()
'Function SchaltTag() As Integer im Modul DatumTextScan

'--------------------------- Globale Variablen und Konstanten --------------------------
Option Explicit
'Public ExAccVersion As String 'Name der Datei, die die Mappe mit ExAcc enthält. Wird am
                             'Anfang von Jahreswechsel und an anderen Stellen ermittelt
                             'und zur Übernahme der VorlVorl gebraucht
'Public MappenName As String  'der Name der Anwendungsmappe in JahrmappeNeu definiert

Public JahresBeginn As Boolean 'Zum Ende der Neueinrichtung True
'---------------------- Kontenplan: Struktur und Parameter ------------------------------
'------- Headerstruktur und Gesamtparameter:
Public Const KPJahresZahlR = 1, KPJahresZahlC = 5, KPCalTagC = 6
Public Const KPRErsteZeile = 5, MaxKoPlaLänge = 300  'zur Vermeidung von Endlosschleife
'------- Erkennungsmerkmale
Public LinkerHeader As String, RechterHeader As String
'------- Spaltenstruktur (z.T. kontoartabhängig redundant):
Public Const KPCE = 1, KPCKonto = 2, KPCArt = 3, KPCBeschr = 4, KPCBlattname = 5
Public Const KPCSamlKto = 6, KPCBZeile = 7, KPCBestätigung = 8, KPCBeriText = 8
Public Const KPCStraße = 9, KPCÜbertrag = 9, KPCLinkHeader = 9
Public Const KPCOrt = 10, KPCVertragText = 10
Public Const KPCRechtHeader = 11
'------- Bereichs-Zeilenstruktur  (out von KontenplanStruktur):
Public KPKZBereich2Vorhanden As Boolean, _
       KPKZBestand As Integer, KPKZAusgaben As Integer, KPKZEinnahmen As Integer, _
       KPKZAusgaben2 As Integer, KPKZEinnahmen2 As Integer, KPKZFonds As Integer, _
       KPKZVermögen As Integer, KPKZMitglieder As Integer, KPKZSpender As Integer, _
       KPKZAE As Integer, KPKZEnde As Integer  '=KPKZ für nächsten (nicht vorhandenen) Bereich
Public SLZZBestand As Integer, SLZZAusgaben As Integer, SLZZEinnahmen As Integer, _
       SLZZAusgaben2 As Integer, SLZZEinnahmen2 As Integer, SLZZFonds As Integer, _
       SLZZVermögen As Integer, SLZZMitglieder As Integer, SLZZSpender As Integer, _
       SLZZAE As Integer
Public BBZZBestand As Integer, BBZZAusgaben As Integer, BBZZEinnahmen As Integer, _
       BBZZAusgaben2 As Integer, BBZZEinnahmen2 As Integer, BBZZFonds As Integer, _
       BBZZVermögen As Integer
Public BereichAnzahl As Long, FondBereichVorhanden As Boolean, _
       KPErstePersonKtoZ As Long, KPLetzteErfolgsKtoZ As Long, _
       PersonKtoBereichVorhanden As Boolean
Public KPVersion As Integer, KPSUrteil As String, KPgeändert As Boolean
'------------------------- Konten: Arten, Struktur, Status ------------------------------
'------- Kontoarten (Kennzeichen für die Bereiche in Kontenplan und Berichten):
Public Const BestandKto = 1, AusgabKto = 2, EingabKto = 3, Ausgab2Kto = 4, Eingab2Kto = 5
Public Const FondsKto = 6, VermögenKto = 7, UndefKto = 8, SammelKto = 9
Public Const MitgliedKto = 10, SpenderKto = 11

'-------- Aktuelle (vom letzten KtoKennDat-Aufruf ermittelte) KontoKenndaten --------
Public AKtoKPZeil As Integer, AKtoEinricht As String, AKtoNr, AKtoArt As Integer, _
       AKtoBeschr As String, AKtoBlatt As String, AKtoSamlKto As Integer, _
       AktoBereichText As String, AKtoBeriZeile As Variant, AKtoBeriText As String, _
       AKtoStatus As Integer, _
       AKtoStrasse As String, AKtoOrt As String, _
       AktoÜbertrag As Double, AktoVertragsText As String, _
       AktoEndSaldo As Double, AktoStornoZahl As Integer, _
       AKtoZiB As Integer, AKtoBeriZeileAlt As Integer, AKto3SternZeile As Integer, _
       AKtoMeldZeile As Integer
'-------- Mehrstufige Dialoge ----------------------------------------------
Public BEENDEN As Boolean, FORTSETZEN As Boolean, UNVERÄNDERT As Boolean
Public AktVorgang As String
       'MELDUNG, ABBRUCH im Modul ERFASSEN deklariert
'------------ Konto-Status ------------------------------------------------------------
Public Const KtoUnbekannt = 0, KtoBlattFehlt = 1, KtoGanzLeer = 2, _
             KtoGeleert = 3, KtoLeerMitÜbertrag = 4, KtoHatBuchungen = 5, _
             KtoHatBuchungenUndÜbertrag = 6
'--------------------------------------------------------------------------------------
'------- Kontoblatt-Spaltenstruktur:
Public Const KoCTANr = 1, KoCDatum = 2, KoCBeleg = 3, KoCBlockDatum = 3, KoCGegKto = 4, _
             KoCBeschr = 5, KoCSaldo = 6, KoCSoll = 7, KoCHaben = 8, KoCBuID = 9
'---------------------------- SaldLi und Bericht: Struktur ------------------------------
'------- Zeilenstruktur
 Option Base 1
Public SaldLiBereichsAnfR(1 To 6) As Long, SaldLiBereichsEndR(1 To 6) As Long
Public BerichtBereichsAnfR(1 To 6) As Long, BerichtBereichsEndR(1 To 6) As Long
Public SaldLiBereichLänge(1 To 6) As Long, BerichtBereichLänge(1 To 6) As Long
Public Const BerBerZeiMax = 20  'Maximale Zeilenzahl in einem Berichtbereich
'----------------------- Von KP-Bearbeitungsfällen gebraucht -------------------------------
Dim Fall As Variant, A As VbMsgBoxStyle, B As VbMsgBoxStyle, C As VbMsgBoxStyle, I As Integer
Dim AktiZel As Range, AZei As Long, AZei1 As Long, ASpal As Long, KPName As String
Dim EndText As String ', BearbFallNurEINFÜGEN As Boolean
Const TiT = "ExAcc Kontenplanbearbeitung"
'------- Bearbfall KontoLeeren
Dim ZuLeerendesBlatt As String
Public KonTan As Integer, KonDat As String, KonZei As Integer, ArProZei As Integer
Public KonBuid As Integer, KPBearbeitung As String, BearbFall As String
Dim TabZeileLetzTanNr As Integer, ArProtEnde As Integer
Dim Bezugszelle As Range, KontoZ As Integer

Sub KONTPLANBEARB() '=====================================================================
Attribute KONTPLANBEARB.VB_ProcData.VB_Invoke_Func = "k\n14"
'Aufgerufen mit "Strg+k" von einer beliebigen Spalte einer Zeile im Kontenplan aus.
'Unterstützt das KONTOLEEREN und BLATTLÖSCHEN von Konten und das EINFÜGEN, ÄNDERN, LÖSCHEN
'und VERSCHIEBEN von Kontozeilen im Kontenplan.
'Die Verzweigung in die Bearbeitungsfälle geschieht in '6KpB--- und kann für mehrstufige
'Bearbeitungsfälle übersprungen (vererbt) werden; das geschieht in '5KpB--- .

'  MsgBox-Fenster Lageeinstellung nicht wirksam:
'  Application.Left = -0.5 'Abstand vom linken Rand des Bildschirms zum Excel-Hauptfenster
'  Application.Top = 7.75  'Abstand vom oberen Rand des Bildschirms zum Excel-Hauptfenster

Dim KPName As String, KPZeil As Integer
Dim A As VbMsgBoxStyle, B As VbMsgBoxStyle, KNr As Variant, GelöschtK As Long
Dim Erläuterung As String, EndText As String, AltesKonto As Integer
Dim StartBlatt As String, StartZeile As Integer ', StartSpalte As Integer
Dim IstBereichsHeader As Boolean

'1KpB ------------------------ Anfangssituation -----------------------------------
AnfangDerBearbeitung:
  ExAccVersion = ThisWorkbook.Name
   If ActiveSheet.Name <> "Kontenplan" Then
    A = MsgBox(prompt:= _
    "Zum Kontenplan ändern mit ''Strg+k'' die zu ändernde KontenplanzeileZeile" & Chr(10) & _
    "(im Falle EINFÜGEN die Zeile in der eingefügt werden soll)" & Chr(10) & _
    "aktivieren. Dorthin wechseln?", Buttons:=vbOKCancel, Title:=TiT)
    If A = vbCancel Then  'kein Aktivieren von Kontenplan, wenn
    End If                'mit Abbrechen quittiert wird
    If A = vbOK Then
      Sheets("Kontenplan").Activate
      
'-----Kontenplanpflege verbieten ----------------
    B = MsgBox(prompt:="Kontenplanpflege leider noch nicht implementiert." & Chr(10) & _
    "Eric fragen!", _
    Buttons:=vbOKOnly, Title:=TiT)
    Exit Sub
'-----Kontenplanpflege Verbieten Ende -----------

      B = MsgBox(prompt:= _
      "Die zu bearbeitende Zeile wählen ünd ''Strg+k'' tippen" & Chr(10) & _
      "(im Falle EINFÜGEN die Zeile in der eingefügt werden soll)", _
      Buttons:=vbOKOnly, Title:=TiT)
    End If
    UNVERÄNDERT = True
    BEENDEN = True
    FORTSETZEN = False
    Exit Sub
  End If  'ActiveSheet.Name <> "Kontenplan"
  StartZeile = ActiveCell.Row
'2KpB ----------------- In 2. Bearbeitungsstufe verzweigen -------------------------
  If Sheets("Kontenplan").Cells(1, 2) <> "" Then '(2.Schritt)
    If Sheets("Kontenplan").Cells(1, 2) = "A" Then GoTo AENDERN
    If Sheets("Kontenplan").Cells(1, 2) = "B" Then GoTo BLATTLÖSCHEN '-----Kontenplanpflege verbieten ----------------
    If Sheets("Kontenplan").Cells(1, 2) = "C" Then GoTo ZEILEKPLÖSCHEN
    If Sheets("Kontenplan").Cells(1, 2) = "D" Then GoTo VERSCHIEBEN
    If Sheets("Kontenplan").Cells(1, 2) = "E" Then GoTo EINFÜGEN
    If Sheets("Kontenplan").Cells(1, 2) = "F" Then GoTo KONTOLEEREN
  End If 'Cells(1, 2) <> ""
'3KpB ----------------------- Aktionen der 1. Stufe ----------------------------
With ActiveWindow 'Workbooks(MappenName)
  If Sheets("Kontenplan").Cells(1, 2) = "" Then
    MELDUNG = ""
    ABBRUCH = False
  End If
'3KpB ------------------ Blatt ist Kontenplan. Auftragsbeginn ------------------
    Range("A" & StartZeile & ":K" & StartZeile & "").Select
    A = MsgBox(prompt:= _
      "Soll die Kontenplanzeile " & StartZeile & " geändert werden " & Chr(10) & _
      "0der" & Chr(10) & _
      "an dieser Stelle eine Kontenplanzeile eingefügt werden," & Chr(10) & _
      "(wobei die Zeile mitsamt den folgenden nach unten geschoben würde)?", _
      Buttons:=vbYesNo, Title:=TiT)
    If A = vbNo Then GoTo Abschließen
'4KpB ------------------ Gewählte Zeile in Zeile 3 sichern ------------------
    If A = vbYes Then
      Range("A" & StartZeile & ":K" & StartZeile & "").Select
      Selection.Copy
      Range("A3:K3").Select
      Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, _
      SkipBlanks:=False, Transpose:=False
      ActiveSheet.Paste
      Application.CutCopyMode = False
      '------------- Zusätzliche Daten des Kontos in Zeile 4 sichern -------------
                       'Zeile ist Bereichs-Header
      Cells(4, 2) = StartZeile
      If Cells(StartZeile, 2) = "" Then IstBereichsHeader = True
      If Cells(StartZeile, 2) <> "" Then
        Call KtoKennDat(Cells(StartZeile, 2))
        Cells(4, 3) = AKtoArt
        Cells(4, 4) = AKtoStatus
        Cells(4, 5) = AKtoZiB
      End If
      GoTo AuftragsWahlS1
    End If
'6KpB --------------------- Bearbeitungsabsicht ermitteln -------------------------
AuftragsWahlS1:
  KNr = Sheets("Kontenplan").Cells(StartZeile, 2)
  With Worksheets("Kontenplan")
    Range("A" & StartZeile & ":K" & StartZeile & "").Select
    Fall = InputBox(prompt:= _
       "Welcher Bearbeitungsfall?" & Chr(10) & Chr(10) & _
       "A  AENDERN dieser Kontenplanzeile " & StartZeile & "      oder" & Chr(10) & _
       "B  BLATTLÖSCHEN des Blatts ''" & Cells(StartZeile, 5) & "''  oder" & Chr(10) & _
       "C  ZEILELÖSCHEN diese Kontenplanzeile oder" & Chr(10) & _
       "D  VERSCHIEBEN dieser KontenplanZeile oder" & Chr(10) & _
       "E  EINFÜGEN einer Zeile über dieser Zeile " & StartZeile & "    oder" & Chr(10) & _
       "F  KONTOLEEREN des Blatts ''" & Cells(StartZeile, 5) & "'' durch Storni" & _
       Chr(10) & Chr(10) & _
       "(Einen der Buchstaben A...F eingeben und OK drücken!)", _
       Title:=TiT, Default:="A") ', Type:=2, Left:=450, Top:=300)
                                 'Type:=2 = String, Type:=4 = Boolean,Type:=6 =
'7KPB -------------- Verzweigungsvorbereitung Bearbeitungsart 2,Stufe ------------------
    If Not (Fall = "A" Or Fall = "B" Or Fall = "C" Or _
           Fall = "D" Or Fall = "E" Or Fall = "F") Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Bearbeitungsfall-Wahl abgebrochen oder" & Chr(10) & _
      "Unverständlicher Kontenplan-Bearbeitungsfall" & Chr(10) & _
      "Nur Buchstabe A ... F erlaubt"
      GoTo Abschließen  'Leeren der Zwischenspeicher
    End If
    Sheets("Kontenplan").Cells(1, 2) = Fall  'Sprungkennzeichen setzen
    
'8KPB ------------------- Im Fall C Kontenplanzeile löschen ----------------------
'              ------------- Einstufig, mit Prüfung ------------
    If Fall = "C" Then
      If Cells(StartZeile, 1) = "" And Cells(StartZeile, 2) = "" And _
         Cells(StartZeile, 3) = "" And Cells(StartZeile, 4) = "" Then _
         GoTo ZeileLoeschen
      If IstBereichsHeader = True Then
        A = MsgBox("Beseitigen eines Bereichsheaders verändert die" & Chr(10) & _
                   "KontoArt der nachfolgenden Konten." & Chr(10) & _
                   "Ist das beabsichtigt?", _
                   vbYesNo, ExAccVersion & "KP-Bearbeitung" & BearbFall)
        If A = vbYes Then
ZeileLoeschen:
          Range("A" & StartZeile & ":K" & StartZeile & "").Select
          Selection.Delete
          Cells(1, 1) = Cells(1, 1) + 1 'Kontenplanversion inkrementieren
          Cells(1, 3) = Cells(1, 3) - 1 'Kontenplanlänge dekrementieren
          GoTo Abschließen
        End If 'A = vbYes
        If A = vbNo Then GoTo Abschließen
      End If 'IstBereichsHeader = True
      If AKtoStatus > KtoGeleert Then
        A = MsgBox("Das Konto " & AKtoNr & " hat Daten,die" & Chr(10) & _
                   "beim löschen der KP-Zeile " & StartZeile & " verloren gehen." & Chr(10) & _
                   "Sie können durch KONTOLEEREN in stornierten" & Chr(10) & _
                   " Buchungszeilen aufbewahrt werden.", _
                   vbOKOnly, ExAccVersion & " KP-Bearbeitung" & BearbFall)
        MELDUNG = MELDUNG & Chr(10) & AktVorgang & " abgebrochen."
        GoTo Abschließen
      End If 'AAKtoStatus > KtoGeleert
      If AKtoStatus <= KtoGeleert Then
        Range("A" & StartZeile & ":K" & StartZeile & "").Select
        Selection.Delete
        Cells(1, 1) = Cells(1, 1) + 1 'Kontenplanversion inkrementieren
        Cells(1, 3) = Cells(1, 3) - 1 'Kontenplanlänge dekrementieren
        GoTo Abschließen
      End If 'AKtoStatus <= KtoGeleert
    End If 'Fall = "C"
'8KPB ------------------- Im Fall E Leerzeile erzeugen ----------------------
'              ------- Einstufig, ohne weitere Prüfung ------------
    If Fall = "E" Then
      Range("A" & StartZeile & ":K" & StartZeile & "").Select
      A = MsgBox("An dieser Stelle eine Leerzeile einfügen?", _
                 vbYesNo, ExAccVersion & "KP-Bearbeitung" & BearbFall)
      If A = vbNo Then GoTo Abschließen
      Selection.Insert shift:=xlDown
      Cells(1, 1) = Cells(1, 1) + 1 'Kontenplanversion erhöhen
      Cells(1, 3) = Cells(1, 3) + 1 'Kontenplanlänge erhöhen
      MELDUNG = "dieseKontenplanzeile darf nicht leer bleiben!"
      GoTo Abschließen
    End If 'Fall = "E"
    Range("A" & StartZeile & ":K" & StartZeile & "").Select
'9KPB ---------------- 2. Schritt vorbereiten ------------------------------------
    Sheets("Kontenplan").Cells(1, 2) = Fall  'Sprungkennzeichen in den KP-Header
    A = MsgBox("Nach Bearbeitung dieser Kontenplanstelle" & Chr(10) & _
       "(Zeile" & StartZeile & ", keinesfalls in einer anderen Zeile!)" & Chr(10) & _
       "eine Zelle in der Zeile " & StartZeile & " aktivieren und" & Chr(10) & _
       "erneut ''Strg+k'' drücken!", _
       vbOKCancel, ExAccVersion & "KP-Bearbeitung" & BearbFall)
    If A = vbOK Then Exit Sub        'Nach Änderung 2.Schritt beginnen
    If A = vbCancel Then GoTo Rücksetzen 'mit Abschließen
'8KPB -------------------  ---------------------
Rücksetzen:
  With Sheets("Kontenplan")
    .Activate
    Range("A3:K3").Select
      Selection.Copy
    Range("A" & Cells(4, 2) & ":K" & Cells(4, 2) & "").Select
      Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, _
      SkipBlanks:=False, Transpose:=False
    ActiveSheet.Paste
    Application.CutCopyMode = False
  End With
Abschließen: '--------SprungKz u. Zwischenspeicher löschen --------------
  With Sheets("Kontenplan")
    Cells(1, 2) = ""
    Range("A2:L4").Select
    Selection.ClearContents
    Range("A" & StartZeile & ":K" & StartZeile & "").Select
    Cells(1, 1) = Cells(1, 1) + 1
    EndText = MELDUNG & Chr(10) & _
       AktVorgang & " abgeschlossen"
    Call MsgBox(EndText, vbOKOnly, TiT)
    AktVorgang = ""
  End With 'Sheets("Kontenplan")
  End With 'Sheets("Kontenplan")
  End With 'ActiveWindow
  Exit Sub
'A ===================== Kontoplanzeile Ändern ==================================
AENDERN:    '(KPZeil)
BearbFall = "AENDERN"
AktVorgang = "AENDERN"
With Sheets("Kontenplan")
  Cells(1, 2) = "A"   'Sprungkennzeichen für Stufe 2 in den KP-Header
  
End With 'Sheets("Kontenplan")
'An dieser Stelle ist die in dem ersten Strg+k - Aufruf angewählte Kontenplan-
'Zeile in der Kontenplanzeile 3 gesichert. In Zeile 4 ist zusätzlich die Zeilen-
'Nummer und die Kontoart abgelegt. An Stelle der ursprünglichen Zeile steht die
'vom Anwender geänderte Zeile.
'Der Code hier vergleicht nach dem gefortderten zweiten Strg+k - Aufruf die
'geänderte Zeile mit der in Zeile 3 befindlichen und entscheidet über die
'Zulässigkeit. Für zulässige Änderungen wird die Konsistenz mit den Kontoblättern
'wiederhergestellt.  Unzulässige Änderungen bewirken das Rücksetzen der Zeile in
'den ursprünglichen Zustand. In jedem Fall wird mit dem Löschen der Sprungmarke
'in Zelle B1 des Kontenplans und der Zeilen3 und 4 abgeschlossen.

'  Dim AlterBlattName As String, NeuerBlattName As String, NeueKontoart As Integer
'  Dim AlteKontoNr As Long, NeueKontoNr As Long, NeuerBezeichner As String
'  Dim KtoNrGut As Boolean, BlaNaGut As Boolean, KtoNr As Long, BlaNa As String
'  Dim AvbOK As Boolean, AvbCancel As Boolean, BvbOK As Boolean, BvbCancel As Boolean

'1A ------------------------------------------
  Dim AlterBlattName As String, NeuerBlattName As String, NeueKtoNr As Integer
  Dim NeuerKtoKopfText As String, NeuerÜbertrag As Double
    AktVorgang = " KP-ZEILE ÄNDERN "
    KPgeändert = False
  With Sheets("Kontenplan")
    .Activate
    KPZeil = Cells(4, 2)  'in 1.Schritt gesichert
'2A --------------- Ausschließen von Bereichskopfzeilen ------------------
    If Cells(KPZeil, 2) = "" Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Kontoart-Kopfzeilen können nicht geändert werden."
      GoTo Rücksetzen
    End If
'3A ------------------- Gleiche Zeile? ----------------------------
    Range("A" & KPZeil & ":K" & KPZeil & "").Select
       'Kontenplanzeile darf nicht gewechselt werden
    If AKtoKPZeil <> Cells(4, 2) Then  'gleiche Zeile erzwingen
      MELDUNG = MELDUNG & Chr(10) & _
        "Aktivierte Zeile stimmt nicht mit der ursprünglich gewählten, " & _
        Cells(4, 2) & ", überein!" & Chr(10) & _
        KPZeil & " manuell prüfen!"
      GoTo Rücksetzen
    End If
          ' Spaltenweise Prüfen und Verarbeiten der Änderungen
'4A ----------------- Spalte 1 "E" ----------------------------
    'Spalte 1 wird durch KtoKennDat auf den aktuellen Wert eingestellt
    If Cells(KPZeil, 1) <> Cells(3, 1) Then
      Cells(KPZeil, 1) = Cells(3, 1)
      MELDUNG = MELDUNG & Chr(10) & _
      "Spalte 1 im Kontenplan darf nur implizit über Wurde Operationen" & Chr(10) & _
      "wie LEEREN, LÖSCHEN, EINRICHTEN geändert werden. Wurde Rückgesetzt."
    End If
'5A ------------------------- Spalte 2 -----------------------
    If Cells(KPZeil, 2) <> Cells(3, 2) Then
      Call KtoKennDat(Cells(3, 2))
      If AKtoStatus > KtoLeerMitÜbertrag Then    'Status 4
        Cells(KPZeil, 2) = Cells(3, 2)
        MELDUNG = MELDUNG & Chr(10) & _
        "Die KontoNr " & Cells(3, 2) & " wurde schon in Buchungen verwendet;" _
        & Chr(10) & "sie darf nur nach LEEREN des Kontos geändert werden."
      End If
      If AKtoStatus <= KtoLeerMitÜbertrag And AKtoStatus >= KtoGanzLeer Then
        Sheets(Cells(3, 5)).Activate
        ActiveSheet.Cells(1, 4) = Sheets("Kontenplan").Cells(KPZeil, 2)
        MELDUNG = MELDUNG & Chr(10) & _
        "Die KontoNr " & Cells(3, 2) & " wurde in " & Cells(KPZeil, 2) & "geändert"
      End If
    End If 'Cells(KPZeil, 2) <> Cells(3, 2)
  End With
'6A------------------------ Spalte 3 -----------------------
 

'Ä2 3 ----------------- Zuerst Prüfen: Spalte 5 Blattname --------------------
    If Cells(KPZeil, 5) <> Cells(3, 5) Then
      AlterBlattName = Cells(3, 5)
      If Cells(3, 1) = "E" Then
        Sheets(AlterBlattName).Name = Cells(KPZeil, 5)
        Cells(KPZeil, 1) = "E"  'Durch Verschwinden des alten Blattnamens
      End If                    'von KtoKennDat beseitigt
      KPgeändert = True
    End If
'Ä2 4 ----------------- Spalte 2 Kontonummer --------------------
    If Cells(3, 2) <> Cells(KPZeil, 2) Then
      If Cells(4, 4) > KtoLeerMitÜbertrag Then
        MELDUNG = MELDUNG & Chr(10) & _
          "Nr eines Kontos mit Buchungen kann nicht geändert werden." & Chr(10) & _
          "ggf. mit KONTOLEEREN bearbeiten, dann umbenennen und dann " & Chr(10) & _
          "die durch das Leeren stornierten Buchungen mit neuer Nr. buchen."
        Cells(KPZeil, 2) = Cells(3, 2)
      End If
      If Cells(4, 4) <= KtoLeerMitÜbertrag Then
        NeuerBlattName = Cells(KPZeil, 5)  'd.h. nach Prüfen der Spalte 5
        NeueKtoNr = Cells(KPZeil, 2)
        Sheets(NeuerBlattName).Cells(1, 4) = NeueKtoNr
        Sheets("Kontenplan").Activate
        KPgeändert = True
      End If
    End If 'Spalte 2 geändert
'Ä2 5 ----------------- Spalte 3 Kontoart --------------------
    If Cells(3, 3) <> Cells(KPZeil, 3) Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Die Kontoart-Spalte kann nicht geändert werden." & Chr(10) & _
      "Ausgangszustand wiederhergestellt."
      Cells(KPZeil, 3) = Cells(3, 3)
    End If
'Ä2 6 ----------------- Spalte 4 Beschreibung --------------------
    If Cells(3, 4) <> Cells(KPZeil, 4) Then
      NeuerKtoKopfText = Cells(KPZeil, 4)
      NeuerBlattName = Cells(KPZeil, 5)  'd.h. nach Bearbeiten der Spalte 5
      Sheets(NeuerBlattName).Cells(1, 5) = NeuerKtoKopfText
      Sheets("Kontenplan").Activate
      KPgeändert = True
    End If
'Ä2 7 ----------------- Spalte 6 Sammelkonto ---------------------
    If Cells(3, 6) <> Cells(KPZeil, 6) Then
      KPgeändert = True
    End If
'Ä2 8 ----------------- Spalte 7 Berichtzeile --------------------
    If Cells(3, 7) <> Cells(KPZeil, 7) Then
      KPgeändert = True
    End If
'Ä2 9 ----------------- Spalte 8 Berichtzeilentext ---------------
    If Cells(3, 8) <> Cells(KPZeil, 8) Then
      KPgeändert = True
    End If
'Ä2 10 ----------------- Spalte 9 Kopfzeilentext ---------------
    If Cells(3, 9) <> Cells(KPZeil, 9) Then
      If AKtoArt = 1 Then
        NeuerKtoKopfText = Cells(KPZeil, 9)
        NeuerBlattName = Cells(KPZeil, 5)  'd.h. nach Bearbeiten der Spalte 5
        Sheets(NeuerBlattName).Cells(3, 5) = NeuerKtoKopfText
        Sheets("Kontenplan").Activate
        KPgeändert = True
      End If
    End If
'Ä2 11 ----------------- Spalte 11 Jahresübertrag ---------------
    If Cells(3, 11) <> Cells(KPZeil, 11) Then
      If AKtoArt = 1 Then
        NeuerÜbertrag = Cells(KPZeil, 11)
        NeuerBlattName = Cells(KPZeil, 5)  'd.h. nach Bearbeiten der Spalte 5
        Sheets(NeuerBlattName).Cells(3, 5) = NeuerÜbertrag
        Sheets("Kontenplan").Activate
        KPgeändert = True
      End If
    End If
'Ä2 12 --------------------- Vollzugstext --------------------
    MELDUNG = MELDUNG & Chr(10) & _
    "Erlaubte Änderungen an der Kontenplanzeile " & _
    KPZeil & " durchgeführt."
    BEENDEN = True
    GoTo EndeKontenBearbeitung
'  End With 'Worksheets("KontenPlan")
'11.3KpB Ende_________________KontenplanZeileÄndern_____________________________________


    



'7KpB ================== Bearbeitungsfall KontoLeeren ================ 19.3.2016==
KONTOLEEREN:  '(KPZeil As Integer) GoTo-Sequenz aus '5KpB---
'Stornierung der betreffenden Buchungen, beginnend mit der jüngsten bis zur Ältesten.
'Diese Buchungen sind dann nur noch in Form durchgestrichener Zeilen im ArProt
'bekannt und müssen irgendwann korrekturgebucht werden, mit oder ohne Änderungen.
'Am Ende des Leerens ist das Kontoblatt im Stand des 31. Dez. des Vorjahres und
'bleibt eingerichtet. Der Jahresanfangsstand eines Bestandskontos kann nicht geleert
'werden.
'Das transaktionsumfassende Löschen bewerkstelligt das Sub Einträgelöschen im
'Modul ArProtSchreiben, Makro ERFASSEN (siehe dort).
'------------ Konto-Status ------------------------------------------------------------
'Public Const KtoUnbekannt = 0, KtoBlattFehlt = 1, KtoGanzLeer = 2, _
'             KtoGeleert = 3, KtoLeerMitÜbertrag = 4, _
'             KtoHatBuchungen = 5, KtoHatBuchungenUndÜbertrag = 6
  Dim DreiSternZeile As Integer, ArProtZeil As Integer
'7.1KpB------------- Aktivierte Zelle für das gewünschte Konto? ---------------------
  AktVorgang = "KONTOLEEREN"
  MELDUNG = ""
  ABBRUCH = False
  With Worksheets("KontenPlan")
    .Activate
    Cells(KPZeil, KPCKonto).Activate
    Call KtoKennDat(Cells(KPZeil, 2))
    If AKtoStatus <= KtoLeerMitÜbertrag Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Konto ''" & Cells(KPZeil, KPCBeschr) & "'' hat keine" & _
                  " Buchungen."
      BEENDEN = True
      GoTo EndeKontenBearbeitung
    End If
    If AKtoStatus = KtoHatBuchungen Then
      A = MsgBox("Dieses Konto,  ''" & Cells(KPZeil, 2) & "'',  durch" & Chr(10) & _
          "Stornobuchungen leeren?" & Chr(10) & _
          "Hinweise:" & Chr(10) & _
          "1.Bei Stornobuchungen bleiben die Buchungstexte in ArProt erhalten;" & Chr(10) & _
          "  sie sind mit Durchstreichung gekennzeichnet und können für die" & Chr(10) & _
          "  erforderlichen Korrekturbuchungen, ggf. nach Abänderung, verwendet" & Chr(10) & _
          "  werden." & Chr(10) & _
          "2.Jahresanfangstände von Bestandskonten können nicht geleert werden.", _
                 vbYesNo, TiT & ": KONTOLEEREN")
      If A = vbNo Then
        MELDUNG = MELDUNG & Chr(10) & _
        "Eine Zelle in der Zeile des zu leerenden Kontos aktivieren " _
                & "und erneut ''Strg+k'' drücken!"
        ABBRUCH = True
        GoTo EndeKontenBearbeitung
      End If
    End If
  End With 'Worksheets("KontenPlan")
'7.2KpB ----- Hauptschleife: Identifizierung der Transaktionen des Kontoblatts ---------
  AnzahlStorni = 0  'im Modul ERFASSEN definiert
  With Worksheets(AKtoBlatt)
    .Activate
    ArProtEnde = ArProtEndZeile  'Function im Modul JahrMappeNeu
    DreiSternZeile = Cells(1, 1).Value
'7.3KpB------ Reihenfolge: von jüngeren zu älteren Transaktionen des Kontos ---------
    
    AnzahlStorni = 1
    For KonZei = DreiSternZeile - 1 To 6 Step -1
      Sheets(AKtoBlatt).Activate    'suche in Spalte "BuId" des Kontoblatts
      If Cells(KonZei, 1) <> "" Then
        KonTan = Cells(KonZei, 1)
        KonDat = Cells(KonZei, 2)
        KonBuid = Cells(KonZei, 9)
'-------- ArProtschleife:EinträgeLöschen von ArProtzeile der ermittelten BuId ---------
        With Worksheets("Arprot")
          .Activate
          For ArProtZeil = 3 To ArProtEnde 'vor der Hauptschleife ermittelt
            Cells(ArProtZeil, APCBuID).Activate
            If Cells(ArProtZeil, APCBuID) = KonBuid Then
              Cells(ArProtZeil, APCTANr).Activate
              Call EinträgeLöschen
              AnzahlStorni = AnzahlStorni + 1
              If MeldeStufe >= 3 Then
                A = MsgBox(prompt:="KONTOLEEREN abbrechen?", Buttons:=vbYesNo, Title:=TiT)
                If A = vbYes Then
                  ABBRUCH = True
                  GoTo EndeKontenBearbeitung
                End If
              End If
              Exit For
            End If
          Next ArProtZeil
        End With 'Worksheets("Arprot")
      End If
    Next KonZei
  End With 'Worksheets(ZuLeerendesBlatt)
  Dim Kto As Long
  With Worksheets("KontenPlan")
    .Activate
    Cells(KPZeil, KPCKonto).Activate
    Kto = ActiveCell.Value
    A = MsgBox("Das Konto ''" & ZuLeerendesBlatt & "'' ist durch " & AnzahlStorni & _
        " Stornobuchungen" & Chr(10) & "geleert, aber noch eingerichtet." & Chr(10) & _
        "Seine Einrichtung muß ggf. durch 'Strg+k/BLATTLÖSCHEN' zurückgenommen werden." _
        & Chr(10) & "Wenn es nicht mehr benutzt werden soll, muß in jeder der " & _
        AnzahlStorni & " Stornobuchungen das Buchungskonto  " & Chr(10) & _
        "''" & Kto & "'' durch ein anderes ersetzt werden.", _
        vbOKOnly, TiT & BearbFall)
  End With
  Call KtoKennDat(Cells(KPZeil, 2)) 'u.a.out: 1 von 4 Kontofüllzuständen
  KPBearbeitung = KPBearbeitung & "KONTOLEEREN"
  BEENDEN = True
GoTo EndeKontenBearbeitung   'End Sub 'KontoLeeren
'7KpB Ende_______________Bearbeitungsfall_KONTOLEEREN______________________________


'8KpB ================== Bearbeitungsfall KontenBlatt Löschen =====================
BLATTLÖSCHEN:
'Tut bei einem AKtoStatus < 2 (Konto unbekannt oder nicht eingerichtet) garnichts,
'verweigert bei einem AktoStatus > 2 das Löschen und macht auf die Möglichkeit des
'Kontoleerens aufmerksam.
'Public Const KtoUnbekannt = 0, KtoBlattFehlt = 1, KtoGanzLeer = 2, _
'             KtoLeerMitÜbertrag = 3, KtoHatBuchungen = 4, KtoMitStorni = 5
  BearbFall = "BLATTLÖSCHEN"
  KPZeil = 3
  Cells(KPZeil, 2).Activate
  Call KtoKennDat(Cells(KPZeil, 2))
  If AKtoStatus = KtoUnbekannt Then
    MELDUNG = MELDUNG & Chr(10) & _
    "Konto ''" & AKtoNr & "'' unbekannt."
    ABBRUCH = True
    GoTo EndeKontenBearbeitung
  End If
  If AKtoStatus = KtoBlattFehlt Then
    MELDUNG = MELDUNG & Chr(10) & _
    "Konto ''" & AKtoNr & "'' hat kein Kontoblatt." & Chr(10) & _
       "Blatt Löschen nicht erforderlich."
    ABBRUCH = True
    GoTo EndeKontenBearbeitung
  End If
  If AKtoStatus = KtoGanzLeer Or AKtoStatus = KtoGeleert Then
    Application.DisplayAlerts = False
    Sheets(AKtoBlatt).Delete
    Application.DisplayAlerts = True
 '   ABBRUCH = False  'darf nicht einseitig verfügt werden
    With Sheets("Kontenplan")
      .Activate
      Cells(AKtoKPZeil, 1) = ""
    End With
    MELDUNG = MELDUNG & Chr(10) & _
    "Kontoblatt von Konto ''" & AKtoNr & "'' gelöscht."
    If AktoStornoZahl <> 0 Then
      MELDUNG = MELDUNG & Chr(10) & _
        "Es hatte " & AktoStornoZahl & " storniert(e) Buchung(en), die" & _
        "gegebenenfalls nur noch in ArProt aufscheinen!"
    End If
    BEENDEN = True
    GoTo EndeKontenBearbeitung
  End If
  If AKtoStatus = KtoLeerMitÜbertrag Then
    MELDUNG = MELDUNG & Chr(10) & _
    "Konto ''" & AKtoNr & "'' ein Bestandskonto mit Jahresübertrag."
    ABBRUCH = True
    GoTo EndeKontenBearbeitung
  End If
  If AKtoStatus = KtoHatBuchungen Then
    MELDUNG = MELDUNG & Chr(10) & _
    "Konto ''" & AKtoNr & "'' hat Buchungsdaten." & Chr(10) & _
       "Das Blatt ''" & AKtoBlatt & "'' kann nur nach ''Str+k/KONTOLEEREN''" & _
       " gelöscht werden."
    ABBRUCH = True
    GoTo EndeKontenBearbeitung
  End If
EndeLöschen:
  BEENDEN = True
  GoTo EndeKontenBearbeitung
  '------------------nur von ZEILEKPLÖSCHEN hierher gebracht----------------------------------------
        Application.DisplayAlerts = False
         Sheets("KB" & Cells(KPZeil, 2) & "").Delete
         Application.DisplayAlerts = True
         With Sheets("Kontenplan")
           Range("A" & KPZeil & ":J" & KPZeil & "").Select
           Selection.Delete shift:=xlUp
           Range("B" & KPZeil & "").Select   'für die Endemeldung
           Cells(1, 3) = Cells(1, 3) - 1
           Cells(1, 1) = Cells(1, 1) + 1
           If Cells(1, 1) > 99 Then  '
             Cells(1, 1) = 1  'KP-Version hochzählen
           End If
         End With
         MELDUNG = MELDUNG & Chr(10) & _
           "Konto ''" & AltesKonto & "'' mit dazugehörigem" & _
           Chr(10) & "Blatt ''KB" & AltesKonto & "'' gelöscht." & Chr(10) & _
           "Bericht-Spalten im Bereich ''" & AktoBereichText & "'' noch passend?"
         BEENDEN = True
         GoTo EndeKontenBearbeitung
         If AKtoStatus = 4 Or AKtoStatus = 5 Then
         MELDUNG = MELDUNG & Chr(10) & _
           "Konto ''" & AltesKonto & "'' kann nicht" & _
           Chr(10) & "werden, da Blatt ''KB" & AltesKonto & "'' Daten enthält."
         ABBRUCH = True
         GoTo EndeKontenBearbeitung
       End If
'8KpB_Ende______________Bearbeitungsfall_KontenBlatt_Löschen______________________


'9KpB ============= Bearbeitungsfall Kontenplanzeile Löschen ===========17.3.2016=
ZEILEKPLÖSCHEN:   '(Zeile)
    Dim ZL As VbMsgBoxStyle, ZL1 As VbMsgBoxStyle
    AktVorgang = " KP-ZEILE LÖSCHEN "
    ABBRUCH = False
    MELDUNG = ""
    With Sheets("Kontenplan")
      Range("A" & KPZeil & ":K" & KPZeil & "").Select
      ZL = MsgBox(prompt:="Diese Zeile im Kontenplan löschen?", _
         Buttons:=vbYesNo, Title:=TiT & "KP-Zeile Löschen")
      If ZL = vbNo Then
        ABBRUCH = True
        GoTo EndeKontenBearbeitung
      End If
      If ZL = vbYes Then
        Cells(KPZeil, 2).Activate
        Call KtoKennDat(Cells(KPZeil, 2))
        If AKtoStatus > KtoBlattFehlt Then
          MELDUNG = MELDUNG & Chr(10) & _
          "Das Konto dieser Kontenplan-Zeile, " & ActiveCell & ", besitzt " & Chr(10) & _
          "ein Kontoblatt. Die Zeile kann nicht gelöscht werden." & Chr(10) & _
          "Zunächst von dieser Kontenplanzeile aus mit" & Chr(10) & _
          "''Strg+K / BLATTLÖSCHEN'' das Blatt entfernen."""
          ABBRUCH = True
          GoTo EndeKontenBearbeitung
        End If
        If AKtoStatus <= KtoBlattFehlt Then
          Range("A" & KPZeil & ":K" & KPZeil & "").Select
          Selection.Delete shift:=xlUp
          MELDUNG = MELDUNG & Chr(10) & _
          "Konto " & GelöschtK & " existiert nicht mehr." & Chr(10) & _
          "In Spalte F des Kontenplans prüfen, ob eine Leerzeile" & Chr(10) & _
          "im Bericht entsteht. Zeilennummern dort änderbar."
          Cells(1, 1) = Cells(1.1) + 1
          If Cells(1, 1) >= 100 Then
            Cells(1, 1) = 1
          End If
          BEENDEN = True
          GoTo EndeKontenBearbeitung
        End If
      End If 'ZL = vbYes
    End With 'Sheets("Kontenplan")
'9KpB Ende__________Bearbeitungsfall_Kontenplanzeile_löschen______________________



'10KpB ================== Bearbeitungsfall Kontozeile Einfügen ====================
EINFÜGEN:
'An dieser Stelle ist die in dem ersten Strg+k - Aufruf angewählte Kontenplan-
'Zeile in der Kontenplanzeile 3 gesichert. In Zeile 4 ist zusätzlich die Zeilen-
'Nummer und die Kontoart abgelegt. An Stelle der ursprünglichen Zeile ist die
'schon im ersten Dialogschritt bereitgestellte und im zweiten vom Anwender
'ausgefüllte Leerzeile.
'Der Code hier prüft die Zeile auf Plausibilität und beseitigt die im ersten
'Strg+k - Aufruf standardmäßig vorgenommenen Sicherungen.
    AktVorgang = " KP-ZEILE EINFÜGEN "
    StartZeile = Cells(4, 2)
    Range("B" & StartZeile & ":K" & StartZeile & "").Select
    A = MsgBox("Diese Kontozeile jetzt ausfüllen!", vbOKOnly, _
               "ExAcc Kontenplanbearbeitung")
    Cells(StartZeile, 1) = ""  'Das neue Konto hat noch kein Blatt
    Call KtoKennDat(Cells(StartZeile, 2))
    If ABBRUCH = True Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Die Daten der neu eingefügten KP-Zeile " & StartZeile & " sind nicht konsistent"
    End If
    GoTo Abschließen
'Ende _____________Bearbeitungsfall Kontozeile EINFÜGEN______________________
 
 
'12KpB ===================Bearbeitungsfall Verschieben========================
VERSCHIEBEN:
'Dialog zum Verschieben einer Kontozeile an eine andere Stelle im Kontenplan.
'Er erfordert mehrmaliges Aufrufen von KontPlanBearb. Die über aufeinander
'folgenden Dialogschritte hinwegzurettenden Daten werden in Kontenplan-
'Kopfzeilen zwischengespeichert: die Bearbeitungs-Fortsetzadresse in
'Zelle F1 und die die Nummer der Ursprungszeile in Zelle B1.
    AktVorgang = " KP-ZEILE VERSCHIEBEN "
    ABBRUCH = False
    MELDUNG = ""
    With Sheets("Kontenplan")
    Range("A" & KPZeil & ":J" & KPZeil & "").Select
    A = MsgBox("Soll diese Kontozeile verschoben werden?", vbYesNo, _
               "ExAcc Kontenplanbearbeitung")
    If A = vbNo Then
      MELDUNG = MELDUNG & Chr(10) & _
      "Eine Zelle in der richtigen Zeile aktivieren und" & _
                    Chr(10) & " erneut ''Strg+k'' drücken!"
      Cells(KPZeil, 2).Select
      ABBRUCH = True
      GoTo EndeKontenBearbeitung
    End If 'A = vbNo
'12.1KpB Versch----------Sichern der Daten für den 2. Dialogschritt -----------
    If A = vbYes Then
      Cells(1, 2) = KPZeil  'Ursprungszeile KPZeil des 1. Dialogschritts
      Cells(1, 6) = "V2"        'Fortsetzungkennzeichen
      MELDUNG = MELDUNG & Chr(10) & _
        "Zielzeile für die zu verschiebende Zeile markieren" & Chr(10) & _
        "(sie wird nicht überschrieben, sondern nach unten verschoben) und" & _
        Chr(10) & "erneut Str+k tippen!"
      FORTSETZEN = True
      GoTo EndeKontenBearbeitung
    End If
  End With 'sheets("Kontenplan")
'12.2KpB Versch---------------- Dialogschritt 2 ------------------------
VerschiebenSchritt2:
  BearbFall = "VERSCHIEBEN"
  With Sheets("Kontenplan")
    Cells(1, 6) = "" 'Fortsetzungskennzeichen löschen
    Range("A" & Cells(1, 2) & ":J" & Cells(1, 2) & "").Select
    Selection.Cut
    Range("A" & KPZeil & ":J" & KPZeil & "").Select 'KPZeil Zielzeile
    Selection.Insert shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A" & Cells(1, 2) & ":J" & Cells(1, 2) & "").Select
    Selection.Delete shift:=xlUp
  End With
   MELDUNG = MELDUNG & Chr(10) & _
    "Ehemalige Kontenplan-Zeile " & Cells(1, 2) & _
    "ist jetzt Zeile" & KPZeil & ""
  BEENDEN = True
  GoTo EndeKontenBearbeitung
'----------------------------------------------------------------------------------
EndeKontenBearbeitung:
AktVorgang = ""
End Sub 'KontenplanBearbeitung'12.3KpB Ende___________________Kontenplanzeile Verschieben_______________________________



Sub KontenplanStruktur()
'Liefert die Bereichskopf-Zeilennummern für bis zu 9 Kontoarten in 10 Globalvariable
'KPKZnn mit nn = Bestand/Ausgaben/Einnahmen/Ausgaben2/Einnahmen2/Fonds/Vermögen/
'Mitglieder/Spender/Ende. Die Bereiche (Kontoarten) Bestand, Ausgabe, und Eingabe
'sind obligatorisch vorhanden. Die nicht vorhandenen haben KPKZnn=0.
'Liefert ebenso die zugehörigen Zeilenzahlen in den Saldlibereichen SBZZnn und in den
'Berichtbereichen BBZZnn. Die Kontoarten Mitglieder und Spender haben keine Variablen
'SBZZnn und BBZZnn. Nicht vorhandene Bereiche haben in der Regel SBZZnn=0 und
'BBZZnn=0. Das ist aber ohne Belang.
'Diese 26 Variablen sind als Public definiert im Modul KontPlanPflege.
'Prüft oder ermittelt die Länge des Kontenplans (KPKZEnde) und stellt sicher, dass sie
'richtig in Zelle A1 eingetragen ist und dass die "KopfZeile" der letzten Zeile das
'Zeichen "***" enthält.
  Dim KPSAbbruch
  Dim A, B, Zeile As Integer, BereichNr As Integer, BZeile As Integer
  Dim AktBlatt As String, KoPLaenge As Integer ', AktZell As Range,
  Dim BereichNrAktuell As Integer, BereichNrNaechst As Integer
  Dim KPKZAktuell As Integer, NaechsteKopfZeile As Integer
  Dim SLZZAktuell As Integer, BBZZAktuell As Integer, MaxBeriZeil As Integer
  Dim AnzahlBereiche As Integer, SuchAnfZeile As Integer
  Dim KPPersonBereichFehlt As Boolean
  Dim ABlaNam As String, ARow As Long, ACol As Long '
  Const TiT = "KontenplanStruktur"
 
 '1 KPS --------------------- Aufbewahren Aufrufsituation --------------------------
  With ActiveSheet          '
    ABlaNam = ActiveSheet.Name
    ARow = ActiveCell.Row
    ACol = ActiveCell.Column
  End With
  KPSAbbruch = False  'unabhängig von der übergeordneten ABBRUCH-Situation
'2 KPS ------- Kontenplanende. Notfalls berichtigen. Endekriterium: 3 Leerzeilen -----------
  With Sheets("Kontenplan")
    .Activate
    If Cells(5, 3) <> 1 Or Cells(5, 4) <> "Bestand" Then
      KPSAbbruch = True
      A = MsgBox(prompt:="Inkorrekte Kontenplanstruktur:" & Chr(10) & _
                         "Bestandsbereich beginnt nicht in Zeile 5", _
                 Buttons:=vbOKOnly, Title:=TiT)
      GoTo KPSBeenden
    End If
    If Cells(Cells(1, 3).Value, 3).Value = "***" Then 'Zelle Kontenplanende richtig
      KPKZEnde = Cells(1, 3).Value
    Else
      For Zeile = 4 To Cells(1, 3).Value + 200  'Fehler bis höchstens 200 angenommen
        If Cells(Zeile, 3).Value = "***" Or _
             (Cells(Zeile, 2).Value = "" And Cells(Zeile + 1, 2).Value = "" And _
              Cells(Zeile + 2, 2).Value = "") Then  '3 aufeinander folgende ""
          Cells(Zeile, 3).Value = "***"             'signalisieren Ende Kontenplan
          Cells(1, 3).Value = Zeile
          KPKZEnde = Zeile '----- = out-Variable, public ------
          Cells(KPKZEnde, 3).Activate
          A = MsgBox(prompt:="Ist das das korrekte Kontenplanende?", _
                     Buttons:=vbYesNo, Title:=TiT)
          If A = vbYes Then Exit For
          If A = vbNo Then
            MELDUNG = MELDUNG & Chr(10) & _
            AktVorgang & " abgebrochen. Kontenplanstruktur sanieren!"
            KPSAbbruch = True
            GoTo KPSBeenden
          End If
        End If
      Next Zeile
    End If 'Else Kontenplanende nicht richtig
  End With 'Sheets("Kontenplan")
'3 KPS ----------------------Kontenplan-BereichsKopfzeilen ----------------------
    'Syntax für Kopfzeilen: Spalten 1 und 2 = "" und Spalte 3 enthält eine Zahl
    'zwischen 1 und 11 oder "***". Public-Variablen im Modul KontPlanPflege
    KPKZBestand = 0
    KPKZAusgaben = 0
    KPKZEinnahmen = 0
    KPKZAusgaben2 = 0
    KPKZEinnahmen2 = 0
    KPKZFonds = 0
    KPKZVermögen = 0
    KPKZMitglieder = 0
    KPKZSpender = 0
    KPKZAE = 0
    
    SLZZBestand = 0     'Die im Folgenden ebenso erzeugten globalen Variablen
    SLZZAusgaben = 0    'SLZZnn (Anzahl der verschiedenen Konten im Bereich und
    SLZZEinnahmen = 0   'BBZZnn (die sich aus der größten Zahl der Spalte (6),
    SLZZAusgaben2 = 0   '"Berichtzeile" innerhalb des Bereichs ergebende Zeilenzahl
    SLZZEinnahmen2 = 0  'in der Berichtsdarstellung) werden im Falle der Kontoarten
    SLZZFonds = 0       '"Mitglieder" und "Spender" nicht für die Berichterstellung
    SLZZVermögen = 0    'gebraucht, aber für die Prüfung der KP-Struktur und für.
    SLZZMitglieder = 0  'die Feststellung, ob eine geeignete FA-Bescheinigungs-
    SLZZSpender = 0     'Vorlage vorhanden ist.
    SLZZAE = 0
    
    BBZZBestand = 0
    BBZZAusgaben = 0
    BBZZEinnahmen = 0
    BBZZAusgaben2 = 0
    BBZZEinnahmen2 = 0
    BBZZFonds = 0
    BBZZVermögen = 0
    '3 KPS --------- Vorhandensein und Ort der Kontoart-Header im Kontenplan  ----------
       'Für jede KtoArt Zeilenzahl in Kplan und SaldLi: SLZZnn, in Bericht: BBZZnn
    With Sheets("Kontenplan")
      .Activate
    NaechsteKopfZeile = 5  'Zeile 3 und 4 (meist ausgeblendet) KP-Änderungs-Zwischen-Puffer
    AnzahlBereiche = 0
    KPKZBereich2Vorhanden = False
BereicheScan:
    For Zeile = 5 To KPKZEnde - 1 'überspringt For wenn KPKZAktuell >= KPKZEnde
      Cells(Zeile, KPCArt).Activate  'nur für Test
      If Cells(Zeile, KPCE) = "" And Cells(Zeile, KPCKonto) = 0 _
                                 And Cells(Zeile, KPCArt) = 0 Then
        Zeile = Zeile + 1   'eine Leerzeile überspringen
      End If
      If Cells(Zeile, KPCE) = "" And Cells(Zeile, KPCKonto) = 0 _
                                 And Cells(Zeile, KPCArt) <> 0 Then
        If Cells(Zeile, KPCArt) = "***" Then
          GoTo NextBereichZeile 'Überspringen, Typen unkompatibel
        End If
      End If
'3.1 KPS ----------------------Konto-Art 1: Bestandskonto ------------------------
        If Cells(Zeile, KPCArt) = 1 Then
          AnzahlBereiche = AnzahlBereiche + 1  'für die Strukturprüfung
          KPKZBestand = Zeile
          MaxBeriZeil = 0
          For BZeile = KPKZBestand + 1 To KPKZEnde
            If Cells(BZeile, KPCKonto) = "" Then Exit For
 '           Cells(BZeile, KPCBZeile).Activate
            If Cells(BZeile, KPCBZeile) > MaxBeriZeil Then
              MaxBeriZeil = Cells(BZeile, KPCBZeile)
            End If
          Next BZeile
          NaechsteKopfZeile = BZeile
          Zeile = BZeile
          SLZZBestand = BZeile - KPKZBestand - 1
          BBZZBestand = MaxBeriZeil
        End If
'3.2 KPS ----------------------Konto-Art 2: Ausgaben ------------------------
        If Cells(NaechsteKopfZeile, KPCArt) <> 2 Then
          MELDUNG = MELDUNG & Chr(10) & _
          "Zeile " & NaechsteKopfZeile & "im Kontenplan ist illegal "
          KPSAbbruch = True
          GoTo KPSBeenden
        End If
        If Cells(NaechsteKopfZeile, KPCArt) = 2 Then
          AnzahlBereiche = AnzahlBereiche + 1  'für die Strukturprüfung
          KPKZAusgaben = NaechsteKopfZeile
          MaxBeriZeil = 0
          For BZeile = KPKZAusgaben + 1 To KPKZEnde
            If Cells(BZeile, KPCKonto) = "" Then
              If Cells(BZeile, KPCArt) <> 3 Then
                MELDUNG = MELDUNG & Chr(10) & _
                "Zeile " & BZeile & "im Kontenplan ist illegal "
                KPSAbbruch = True
                GoTo KPSBeenden
              End If
              Exit For
            End If
            If Cells(BZeile, KPCBZeile) > MaxBeriZeil Then
              MaxBeriZeil = Cells(BZeile, KPCBZeile)
            End If
          Next BZeile
          NaechsteKopfZeile = BZeile
          Zeile = BZeile
          SLZZAusgaben = BZeile - KPKZAusgaben - 1
          BBZZAusgaben = MaxBeriZeil
        End If
'3.3 KPS ----------------------Konto-Art 3: Einnahmen -------------------------
        If Cells(NaechsteKopfZeile, KPCArt) <> 3 Then
          MELDUNG = MELDUNG & Chr(10) & _
          "Zeile " & NaechsteKopfZeile & "im Kontenplan ist illegal "
          KPSAbbruch = True
          GoTo KPSBeenden
        End If
        If Cells(NaechsteKopfZeile, KPCArt) = 3 Then
          AnzahlBereiche = AnzahlBereiche + 1  'für die Strukturprüfung
          KPKZEinnahmen = NaechsteKopfZeile
          MaxBeriZeil = 0
          For BZeile = KPKZEinnahmen + 1 To KPKZEnde
            If Cells(BZeile, KPCKonto) = "" Then Exit For
'            Cells(BZeile, KPCBZeile).Activate
            If Cells(BZeile, KPCBZeile) > MaxBeriZeil Then
              MaxBeriZeil = Cells(BZeile, KPCBZeile)
            End If
          Next BZeile
          NaechsteKopfZeile = BZeile
          Zeile = BZeile
          SLZZEinnahmen = BZeile - KPKZEinnahmen - 1
          BBZZEinnahmen = MaxBeriZeil
        End If
'3.4 KPS ----------------------Konto-Art 4: Ausgaben2-Konto -------------------------
        If Cells(NaechsteKopfZeile, KPCArt) = 4 Then
          AnzahlBereiche = AnzahlBereiche + 1  'für die Strukturprüfung
          KPKZBereich2Vorhanden = True     'für Wahl der SaldLi2Vorlage
          KPKZAusgaben2 = NaechsteKopfZeile
          MaxBeriZeil = 0
          For BZeile = KPKZAusgaben2 + 1 To KPKZEnde
            If Cells(BZeile, KPCKonto) = "" Then Exit For
 '           Cells(BZeile, KPCBZeile).Activate
            If Cells(BZeile, KPCBZeile) > MaxBeriZeil Then
              MaxBeriZeil = Cells(BZeile, KPCBZeile)
            End If
          Next BZeile
          NaechsteKopfZeile = BZeile
          Zeile = BZeile
          SLZZAusgaben2 = BZeile - KPKZAusgaben2 - 1
          BBZZAusgaben2 = MaxBeriZeil
        End If
'3.5 KPS ----------------------Konto-Art 5: Einnahmen2-Konto -------------------------
        If Cells(NaechsteKopfZeile, KPCArt) = 5 Then
          AnzahlBereiche = AnzahlBereiche + 1  'für die Strukturprüfung
          KPKZBereich2Vorhanden = True   'für Wahl der SaldLi2Vorlage
          KPKZEinnahmen2 = NaechsteKopfZeile
          MaxBeriZeil = 0
          For BZeile = KPKZEinnahmen2 + 1 To KPKZEnde
            If Cells(BZeile, KPCKonto) = "" Then Exit For
'            Cells(BZeile, KPCBZeile).Activate
            If Cells(BZeile, KPCBZeile) > MaxBeriZeil Then
              MaxBeriZeil = Cells(BZeile, KPCBZeile)
            End If
          Next BZeile
          NaechsteKopfZeile = BZeile
          Zeile = BZeile
          SLZZEinnahmen2 = BZeile - KPKZEinnahmen2 - 1
          BBZZEinnahmen2 = MaxBeriZeil
        End If
'3.6 KPS ----------------------Konto-Art 6: Fondskonto -------------------------
        If Cells(NaechsteKopfZeile, KPCArt) = 6 Then
          AnzahlBereiche = AnzahlBereiche + 1  'für die Strukturprüfung
          KPKZFonds = NaechsteKopfZeile
          MaxBeriZeil = 0
          For BZeile = KPKZFonds + 1 To KPKZEnde
            If Cells(BZeile, KPCKonto) = "" Then Exit For
 '           Cells(BZeile, KPCBZeile).Activate
            If Cells(BZeile, KPCBZeile) > MaxBeriZeil Then
              MaxBeriZeil = Cells(BZeile, KPCBZeile)
            End If
          Next BZeile
          NaechsteKopfZeile = BZeile
          Zeile = BZeile
          SLZZFonds = BZeile - KPKZFonds - 1
          BBZZFonds = MaxBeriZeil
        End If
'3.7 KPS ----------------------Konto-Art 7: Immobilienkonto -------------------------
        If Cells(NaechsteKopfZeile, KPCArt) = 7 Then
          AnzahlBereiche = AnzahlBereiche + 1  'für die Strukturprüfung
          KPKZVermögen = NaechsteKopfZeile
          MaxBeriZeil = 0
          For BZeile = KPKZVermögen + 1 To KPKZEnde
            If Cells(BZeile, KPCKonto) = "" Then Exit For
'            Cells(BZeile, KPCBZeile).Activate
            If Cells(BZeile, KPCBZeile) > MaxBeriZeil Then
              MaxBeriZeil = Cells(BZeile, KPCBZeile)
            End If
          Next BZeile
          NaechsteKopfZeile = BZeile
          Zeile = BZeile
          SLZZVermögen = BZeile - KPKZVermögen - 1
          BBZZVermögen = MaxBeriZeil
        End If
'3.8 KPS ----------------------Konto-Art 8: Nicht definiert -------------------------
         If Cells(NaechsteKopfZeile, KPCArt) = 8 Then
          MELDUNG = MELDUNG & Chr(10) & _
          "Kontenplan Zeile " & NaechsteKopfZeile & ": Kontoart 8 ist nicht definiert. " & Chr(10) & _
              "Kontenplan korrigieren!"
              ABBRUCH = True
          GoTo KPSBeenden
        End If
'3.9 KPS ----------------------Konto-Art 9: Nicht definiert -------------------------
         If Cells(NaechsteKopfZeile, KPCArt) = 9 Then
          MELDUNG = MELDUNG & Chr(10) & _
         "Kontenplan Zeile " & NaechsteKopfZeile & ": Kontoart 9 ist nicht definiert. " & Chr(10) & _
              "Kontenplan korrigieren!"
              KPSAbbruch = True
          GoTo KPSBeenden
        End If
'3.10 KPS ----------------------Konto-Art 10: Mitgliederkonto -------------------------
        If Cells(NaechsteKopfZeile, KPCArt) = 10 Then
          AnzahlBereiche = AnzahlBereiche + 1  'für die Strukturprüfung
          KPKZMitglieder = NaechsteKopfZeile
          For BZeile = KPKZMitglieder + 1 To KPKZEnde
            If Cells(BZeile, KPCKonto) = "" Then Exit For
          Next BZeile
          NaechsteKopfZeile = BZeile
          Zeile = BZeile
          SLZZMitglieder = BZeile - KPKZMitglieder - 1
        End If
'3.11 KPS ----------------------Konto-Art 11: Spenderkonto -------------------------
        If Cells(NaechsteKopfZeile, KPCArt) = 11 Then
          AnzahlBereiche = AnzahlBereiche + 1  'für die Strukturprüfung
          KPKZSpender = NaechsteKopfZeile
          For BZeile = KPKZSpender + 1 To KPKZEnde
            If Cells(BZeile, KPCKonto) = "" Then Exit For
          Next BZeile
          NaechsteKopfZeile = BZeile
          Zeile = BZeile
          SLZZSpender = BZeile - KPKZSpender - 1
        End If
'3.11 KPS ----------------Konto-Art 12: Aufwandsentschädigungskonto ------------------
        If Cells(NaechsteKopfZeile, KPCArt) = 12 Then
          AnzahlBereiche = AnzahlBereiche + 1  'für die Strukturprüfung
          KPKZAE = NaechsteKopfZeile
          For BZeile = KPKZAE + 1 To KPKZEnde
            If Cells(BZeile, KPCKonto) = "" Then Exit For
          Next BZeile
          NaechsteKopfZeile = BZeile
          Zeile = BZeile
          SLZZAE = BZeile - KPKZAE - 1
        End If
'3.12 KPS ----------------- nicht implementierte Kontoarten -------------------------
        If Cells(NaechsteKopfZeile, KPCArt) > 12 And _
           Cells(NaechsteKopfZeile, KPCArt) <> "***" Then
          MELDUNG = MELDUNG & Chr(10) & _
          "Der Kontenplan enthält eine Bereichkopfzelle der Art " & _
          Cells(NaechsteKopfZeile, KPCArt) & Chr(10) & _
          "Kontoarten > 12 sind nicht definiert. " & Chr(10) & _
          "Abbruchgrund. Kontenplan sanieren!"
          KPSAbbruch = True
          GoTo KPSBeenden
        End If
NextBereichZeile:
    Next Zeile  'Kopfzeilen-Scan
  End With 'Sheets("Kontenplan")'
KoPlaScanEnde:
'4 KPS ---------------------------- Prüfungen --------------------------------------
'4.1 KPS ----------- Obligatorische Bereiche vorhanden? ----------------------------
    If KPKZBestand = 0 Or KPKZAusgaben = 0 Or KPKZEinnahmen = 0 Then
      MELDUNG = MELDUNG & Chr(10) & _
          "Die Bereiche 1 ''Bestand'', 2 ''Ausgaben'' und 3 ''Einnahmen''" & Chr(10) & _
          "müssen mit mindestens je einem Konto vertreten sein." & Chr(10) & _
          "Kontenplan sanieren!"
          KPSAbbruch = True
       GoTo KPSBeenden
    End If 'KPKZBestand = 0 ...
'4.2 KPS --------------- Bereich "Bestand" bei Zeile 5 beginnend? ------------------
    If KPKZBestand <> 5 Then
      Sheets("Kontenplan").Activate
      MELDUNG = MELDUNG & Chr(10) & _
          "Zeile 5 des Kontenplans muss " & Chr(10) & _
          "in der Spalte 3 eine 1 enthalten und" & Chr(10) & _
          "die Spalten 1 und 2 müssen leer sein." & Chr(10) & _
          AktVorgang & " abgebrochen.    Kontenplan sanieren!"
          KPSAbbruch = True
      GoTo KPSBeenden
    End If 'KPKZBestand <> 5
'4.3 KPS ----------- Keine Leerzeilen, Keine doppelt vorhandenen Bereiche ------------

    KoPLaenge = KPKZBestand + SLZZBestand + SLZZAusgaben + SLZZEinnahmen + SLZZAusgaben2 _
      + SLZZEinnahmen2 + SLZZFonds + SLZZVermögen + SLZZMitglieder + SLZZSpender + SLZZAE _
      + AnzahlBereiche
    If KoPLaenge <> KPKZEnde Then
      Sheets("Kontenplan").Activate
      MELDUNG = MELDUNG & Chr(10) & _
          "Im Kontenplan sind keine Leerzeilen oder " & Chr(10) & _
          "doppelt vorhandene Bereiche erlaubt!" & Chr(10) & _
          AktVorgang & " abgebrochen.    Kontenplan sanieren!"
          KPSAbbruch = True
      GoTo KPSBeenden
    End If
'4.4 KPS --------- Eindeutige Kontonummern und Blattnamen --------------------------
  With Sheets("Kontenplan")
    SuchAnfZeile = 5
    For Zeile = SuchAnfZeile To KPKZEnde - 1
      If Cells(Zeile, KPCKonto) = 0 Then
        Zeile = Zeile + 1
      End If
      If Zeile >= KPKZEnde Then
        Exit For
      End If
      If KtoUndBlatEindeutig(Zeile) = False Then
        MELDUNG = MELDUNG & Chr(10) & _
        "Im Kontenplan ist die Kontonummer " & Cells(Zeile, KPCKonto) & Chr(10) & _
        "oder der Blattname " & Cells(Zeile, KPCBlattname) & Chr(10) & _
        " nicht eindeutig, d.h. mehrfach vorhanden." & Chr(10) & _
        AktVorgang & " abgebrochen.    Kontenplan sanieren!"
        KPSAbbruch = True
        GoTo KPSBeenden
      End If
    Next Zeile
'5 KPS ----------------- ggf. global benötigte, abgeleitete Parameter ------------
MitgliederBereich:
  If KPKZMitglieder <> 0 Then
    PersonKtoBereichVorhanden = True
    KPLetzteErfolgsKtoZ = KPKZMitglieder - 1
    KPErstePersonKtoZ = KPKZMitglieder + 1
  Else
    PersonKtoBereichVorhanden = False
    KPLetzteErfolgsKtoZ = KPKZEnde - 1
  End If
'6 KPS ----------- Transaktjahr und Version public hinterlegen -----------------
  KPVersion = Sheets("Kontenplan").Cells(1, 1)
  TransaktJahr = Sheets("Kontenplan").Cells(1, 5)
  GoTo KPSBeenden
  End With 'Kontenplan
'6 KPS------------------ Wiederherstellen Aufrufsituation -------------------------------
KPSBeenden:
  If KPSAbbruch = True Then
    ABBRUCH = True  'in den übergeordneten ABBRUCH integrieren
    Call MsgBox(MELDUNG, vbOKOnly, KPVersion)
  End If
  Cells(4, 1).Activate
  Sheets(ABlaNam).Activate
  ActiveSheet.Cells(ARow, ACol).Activate
End Sub 'KontenplanStruktur()

Function KtoUndBlatEindeutig(AnfZeil As Integer) As Boolean
'Prüfen, ob die in (KoPlaZeil, Spalte2) befindliche Kontonummer zwischen dieser
'Zeile und dem Kontenplanende nicht nochmal vorkommt
  Dim KtoNr As Long, Z As Integer, BlaName As String, ARow As Long, ACol As Long
  Dim KoPlaEnde As Long
  
  With Sheets("Kontenplan")
    KtoNr = Cells(AnfZeil, KPCKonto)       'Konto und Blatt, die in ihren
    BlaName = Cells(AnfZeil, KPCBlattname) 'jeweiligen Spalten abwärts nicht
    KoPlaEnde = Cells(1, 3)                          'vorkommen sollen
    KtoUndBlatEindeutig = True
    For Z = AnfZeil + 1 To KoPlaEnde
      If Cells(Z, KPCKonto) = 0 Then GoTo NZ  'Vergleich überspringen
      If Cells(Z, KPCKonto) = KtoNr Or _
         (Cells(Z, KPCBlattname) = BlaName And _
         BlaName <> "") Then   'blanke Zellen werden nicht verglichen
        KtoUndBlatEindeutig = False
        Exit For
      End If
NZ:
    Next Z
  End With 'Sheets("Kontenplan")
End Function
Function KontoNrVomBlatt(Blatt As String) As Integer
  Dim AktivBlatt As String, AktivZeile As Integer, AktivSpalte As Integer
  Dim BZ As Integer
  AktivBlatt = ActiveSheet
  AktivZeile = ActiveCell.Row
  AktivSpalte = ActiveCell.Column
  Sheets("Kontenplan").Activate
  With Sheets("Kontenplan")
    For BZ = 6 To Cells(1, 3)
      If Cells(BZ, 5) = "" Then GoTo NaechstBZ
      If Cells(BZ, 5) = Blatt Then
        KontoNrVomBlatt = Cells(BZ, 2)
        GoTo AnfZustand
      End If
NaechstBZ:
    Next BZ
  End With
  KontoNrVomBlatt = 0   'Blatt nicht gefunden
AnfZustand:
  Sheets(AktivBlatt).Activate
  ActiveSheet.Cells(AktivZeile, AktivSpalte).Activate
End Function
'==============================================================================
Sub KtoKennDat(AktKtoNr)          '***neu*** statt Function KtoPlaDat
'out: AKtoEinricht As String, AKtoNr, AKtoArt As Integer, AKtoBeschr As String,
'     AKtoBlatt, AKtoSamlKto, AKtoBeriZeile As Variant, AKtoBeriText As String,
'     AKtoStrasse As String, AKtoOrt As String, (im Fall Personenkonto)
'     AKtoÜbertrag As Double, AktoVertragsText as String,(im Fall Bestandskonto)
'     AKtoEndSaldo As Double
'     MELDUNG as String, ABBRUCH as Boolean
'     AKtoStatus:  KtoUnbekannt = 0, KtoBlattFehlt = 1, KtoGanzLeer = 2, _
                   KtoGeleert = 3, KtoLeerMitÜbertrag = 4, KtoHatBuchungen = 5, _
                   KtoHatBuchungenUndÜbertrag = 6
'Liefert die KennDaten der "AktKtoNr" aus der zu "AKtoNr" gehörenden Zeile des
'Kontenplans in die unter "out:" genannten globalen Variablen. Die Kenndaten
'"KtoArt", "KtoBlatt" und "AKtoStatus, die nicht explizit im Kontenplan dargestellt
'sind, werden folgendermaßen ermittelt:
'"KtoArt" wird vom Header des Kontogruppe entnommen.,
'"KtoBlatt" wird aus der Kontonummer mit vorangestelltem "KB" zusammengesetzt.
'"AKtoStatus"  wird durch Analyse des Kontenblatts ermittelt. Die erste Kontenplan-
'Spalte ("E") wird geprüft und ggf. korrigiert.'
'Aktives Blatt und aktive Zelle des aufrufenden Programms bleiben erhalten.

Dim KtoPlaZeile As Integer, ZiG As Integer, AktZell As Range, AktBlatt As String
Dim Sp As Integer, KpSpZahl As Integer, A As VbMsgBoxResult, W, DSZ As Integer
Dim KKDBlattName As String, EKorrekt As Boolean
Dim AZrow As Integer, AZcolumn As Integer, MeldeText As String, KZeile As Integer
Dim KKDAbbruch As Boolean
  With ActiveSheet          'Aufbewahren Aufrufsituation
    AktBlatt = ActiveSheet.Name
    AZrow = ActiveCell.Row
    AZcolumn = ActiveCell.Column
    End With                  '----------------------------
'  MELDUNG = ""
  KKDAbbruch = False  'Keine Beeinflussung der ABBRUCH-Situation
  '------------- Alte Globalwerte rücksetzen (wegen Aussprüngen)---------------
  AKtoKPZeil = 0
  AKtoBlatt = ""
  AKtoArt = 0
  AKtoZiB = 0
  AktoBereichText = ""
  AKtoBeschr = ""
  AKtoBeriZeile = 0
  AKtoBeriText = ""
  AKtoStrasse = ""       'nur Personenkonto
  AKtoOrt = ""           'nur Personenkonto
  AktoÜbertrag = 0       'nur Bestandskonto
  AktoVertragsText = ""  'nur Bestandskonto
  AktoEndSaldo = 0
  '--------------------- Kontenplan-Zeile der AktKtonr suchen -------------------
  With Sheets("KontenPlan")
    .Activate
    If KPKZEnde = 0 Then
      Call KontenplanStruktur
      If ABBRUCH = True Then
        MELDUNG = MELDUNG & Chr(10) & _
        "Kontenplanstruktur defekt"
        GoTo Ausgang
      End If
    End If
    KtoPlaZeile = 5
    Do While KtoPlaZeile < KPKZEnde 'Sheets("Kontenplan").Cells(1, 2).Value
      Cells(KtoPlaZeile, KPCKonto).Activate
      If Cells(KtoPlaZeile, KPCKonto) = AktKtoNr Then
        GoTo KontoNrGefunden
      End If
      KtoPlaZeile = KtoPlaZeile + 1  'interne Variable
    Loop
    AKtoStatus = KtoUnbekannt    'KontoNichtGefunden, Konto nicht im Kontenplan
    MELDUNG = MELDUNG & Chr(10) & _
          "Das Konto " & AktKtoNr & " ist nicht definiert." & Chr(10) & _
          "Um es in den Kontenplan einzufügen, vom Kontenplan aus" & Chr(10) & _
          "Stg+k / EINFÜGEN verwenden!"
    KKDAbbruch = True
    GoTo Ausgang
    '-------------------Kenndaten in die gobalen Variablen eintragen ------------------
KontoNrGefunden:
    AKtoKPZeil = KtoPlaZeile
    '-------------Spalte 1: AKtoEinricht prüfen ggf. Korrigieren -------------------
    AKtoEinricht = Sheets("Kontenplan").Cells(KtoPlaZeile, KPCE)
    AKtoBlatt = Sheets("Kontenplan").Cells(KtoPlaZeile, KPCBlattname)
 '   MELDUNG = ""
    If AKtoEinricht = "" Then
      For Each W In Worksheets            'E-Konsistenz-Korrektur
        If W.Name = AKtoBlatt Then        'E ergänzen, da fälschlicherweise
          Cells(AKtoKPZeil, KPCE) = "E"   'nicht vorhanden
          AKtoEinricht = "E"
          MELDUNG = MELDUNG & Chr(10) & "Hinweis: " & Chr(10) & _
                      "Da Blatt ''" & AKtoBlatt & "'' vorhanden," _
                      & Chr(10) & "im Kontenplan in Zeile " _
                      & AKtoKPZeil & " die 1. Spalte ''E'' gesetzt."
          GoTo EKonsistenzGeprüft
        End If
      Next W
    End If 'AKtoEinricht = ""  'ggf. gemäß Prüfung abgewandelt
    If AKtoEinricht = "E" Then
      EKorrekt = False
      For Each W In Worksheets     'E-Konsistenz-Korrektur
        If W.Name = AKtoBlatt Then
          EKorrekt = True
          Exit For
        End If
      Next W
      If EKorrekt = False Then
        Cells(KtoPlaZeile, KPCE) = ""     'E Löschen, da Kontoblatt nicht
        AKtoEinricht = ""                 'vorhanden
        MELDUNG = MELDUNG & Chr(10) & "Hinweis: " & Chr(10) & _
                    "Da Blatt ''" & KKDBlattName & "'' nicht eingerichtet," _
                    & Chr(10) & "im Kontenplan in Zeile " _
                    & AKtoKPZeil & " die 1. Zelle gelöscht."
      End If
    End If  'AktoEinricht = "E"   'ggf. gemäß Prüfung abgewandelt
EKonsistenzGeprüft:
    '-------------Spalte 2: AktoNr ----------------------------
    AKtoNr = AktKtoNr 'Aufrufparameter
    '---------------- Spalte 3: KPCArt: Vererbung von Bereichskopf ------------------
    ZiG = KtoPlaZeile   'Zeile in der Gruppe
    Do While Cells(ZiG, KPCKonto) <> "" '"" in KPCKonto-Spalte = Kennzeichen für
      ZiG = ZiG - 1                     'Bereichskopf
    Loop
    AKtoArt = Sheets("Kontenplan").Cells(ZiG, KPCArt)  'Art aus Bereich-Kopf
    '--------------- Zeile im Bereich (Abstand vom Bereichskopf) --------------
    AKtoZiB = KtoPlaZeile - ZiG  'Zeile innerhalb Gruppe gleicher Kontoart
    '-------------------Spalte 3 ---------------------------------------------
'   AktoBereichText = unten ermittelt
    '-------------------Spalte 4 ---------------------------------------------
    AKtoBeschr = Cells(KtoPlaZeile, KPCBeschr)
    AktoBereichText = Cells(ZiG, KPCBeschr)
    '-------------------Spalte 5 Blattname ------------------------------------
    If AKtoNr <> 0 And Cells(KtoPlaZeile, KPCBlattname) = "" Then
      AKtoBlatt = "KB" & Cells(KtoPlaZeile, KPCKonto)    'ersatzweise vergeben,
      Cells(KtoPlaZeile, KPCBlattname) = AKtoBlatt       'aus Spalte 2 erzeugt
      AKtoEinricht = ""
      MELDUNG = MELDUNG & Chr(10) & "Hinweis: " & Chr(10) & _
      "Fehlender Blattname " & AKtoBlatt & " wurde von KtoKennDat vergeben"
    End If
    If Cells(KtoPlaZeile, KPCBlattname) <> "" Then
      AKtoBlatt = Cells(KtoPlaZeile, KPCBlattname)
    End If
    '-------------------Spalte 6 ---------------------------------------------
    AKtoSamlKto = Cells(KtoPlaZeile, KPCSamlKto)     'Spalte 6
    '-------------------Spalte 7 Zeile im Bericht ---------------------------------------------
    If Not (AKtoArt = 10 Or AKtoArt = 11) Then
      AKtoBeriZeile = Cells(KtoPlaZeile, KPCBZeile)  'Spalte 7
    End If
    '-------------------Spalte 8 ---------------------------------------------
    AKtoBeriText = Cells(KtoPlaZeile, KPCBeriText)   'Spalte 8
    '-------------------Spalte 9 und 10 --------------------------------------
    If AKtoArt = 10 Or AKtoArt = 11 Then
      AKtoStrasse = Cells(KtoPlaZeile, KPCStraße)      'Spalte 9
      AKtoOrt = Cells(KtoPlaZeile, KPCOrt)             'Spalte 10
    End If
    If AKtoArt = 1 Or AKtoArt = 6 Or AKtoArt = 7 Then
      AktoÜbertrag = Cells(KtoPlaZeile, KPCÜbertrag)         'Spalte 9
      AktoVertragsText = Cells(KtoPlaZeile, KPCVertragText) 'Spalte 10
    End If
  End With 'Sheets("KontenPlan")
    
'-------------------------- Konto-Status --------------------------
'Public Const KtoUnbekannt = 0, KtoBlattFehlt = 1, KtoGanzLeer = 2,
'             KtoGeleert = 3, KtoLeerMitÜbertrag = 4, _
'             KtoHatBuchungen = 5, KtoHatBuchungenUndÜbertrag = 6
'----------------- AktoStatus (von 1 bis 6)  ----------------------
  If AKtoEinricht = "" Then
    AKtoStatus = KtoBlattFehlt  '(AktoStatus 1)
    AktoStornoZahl = 0
    AKto3SternZeile = 0
    MELDUNG = ""
    KKDAbbruch = False
    GoTo Ausgang
  End If
  With Sheets(AKtoBlatt)
    .Activate
    If Cells(4, 6) = 0 And Cells(6, 2) = "***" And Cells(1, 9) <= 0 Then
      AKtoStatus = KtoGanzLeer  '(=2)
      AKto3SternZeile = 6
      GoTo DreiSternZeilePrüfen
    End If
    If Cells(Cells(1, 1), 2) = "***" And _
       Cells(Cells(1, 1), 2).Offset(2, 4) = 0 Then  'Stornozahl in (1,9)
      AKtoStatus = KtoGeleert   '(=3)
      AktoStornoZahl = Cells(1, 9)
      AKto3SternZeile = Cells(1, 1)
      GoTo DreiSternZeilePrüfen
    End If
    If Cells(4, 6) <> 0 And Cells(6, 2) = "***" Then
      AKtoStatus = KtoLeerMitÜbertrag '(=4)
      AktoÜbertrag = Cells(4, 6)
      AKto3SternZeile = 6
      GoTo DreiSternZeilePrüfen
    End If
    If Cells(1, 1) > 6 And Cells(4, 6) = 0 Then
      AKtoStatus = KtoHatBuchungen  '(=5)
      GoTo DreiSternZeilePrüfen
    End If
    If Cells(1, 1) > 6 And Cells(4, 6) <> 0 Then
      AKtoStatus = KtoHatBuchungenUndÜbertrag  '(=6)
      GoTo DreiSternZeilePrüfen
    End If
DreiSternZeilePrüfen:

'--------------- Dreisternzeile korrekt vermerkt ---------------
    If Cells(Cells(1, 1), KoCDatum) = "***" And _
       Cells(Cells(1, 1) + 1, KoCDatum + 3) = "Umsatz Periode" And _
       Cells(Cells(1, 1) + 2, KoCDatum + 3) = "Kontostand" Then
      AKto3SternZeile = Cells(1, 1)
      GoTo DSZrichtig
    End If
'---------------------- Dreisternzeile suchen ------------------
    For DSZ = 6 To Cells(1, 1) + 10
      Cells(DSZ, KoCDatum).Activate
      If Cells(DSZ, KoCDatum) = "***" Then
        Cells(1, 1) = DSZ
        GoTo Umgebung3SternZeile
      End If
    Next DSZ
KeineDreisternZeile:
    MELDUNG = MELDUNG & Chr(10) & _
    "Blatt ''" & AKtoBlatt & "'', Konto " & AktKtoNr & _
    " hat einen unbrauchbaren Abschluss. ''***''-Zeile korrigieren!"
    KKDAbbruch = True
    GoTo Ausgang
    '--------------- Im Kontoblatt Umgebung 3SternZeile prüfen ----------------
Umgebung3SternZeile:
    AKto3SternZeile = DSZ
    Sheets(AKtoBlatt).Cells(DSZ + 1, 5).Activate
    If ActiveCell.Value <> "Umsatz Periode" Then GoTo AktoAbbrMeldung
    Sheets(AKtoBlatt).Cells(DSZ + 2, 5).Activate
    If ActiveCell <> "Kontostand" Then
      ActiveCell = "Kontostand"
    End If
 '   GoTo AktoAbbrMeldung
    If Cells(DSZ + 1, 2) = "" And _
       Cells(DSZ + 2, 2) = "" And _
       Cells(DSZ + 3, 3) = "" And _
       Cells(DSZ + 4, 3) = "" And _
       Cells(DSZ + 5, 3) = "" And _
       Cells(DSZ + 6, 3) = "" Then
       GoTo DSZrichtig
    Else
      GoTo AktoAbbrMeldung
    End If
AktoAbbrMeldung:
    MELDUNG = MELDUNG & Chr(10) & _
    "Im Kontoblatt ''" & AKtoBlatt & "'' (Konto " & AKtoNr & ") die Umgebung der" & _
    "''***''-Zeile prüfen! Abbruchgrund."
    KKDAbbruch = True
    GoTo Ausgang
  End With 'Sheets(AKtoBlatt)
 '     For KZeile = 6 To DSZ - 1            'etwa vorhandene
 '       If Cells(KZeile, 2) = "***" Then   'falsche *** löschen
 '         Cells(KZeile, 2) = ""
 '         meldung = meldung & Chr(10) & _
 '         "KontoKennDat hat in Kontoblatt " & AKtoBlatt & _
 '         " Zeile " & KZeile & " falsches ''***'' gelöscht."
 '       End If
 '     Next KZeile
 '     For KZeile = DSZ + 1 To DSZ + 20     'etwa vorhandene
 '       If Cells(KZeile, 2) = "***" Then   'falsche *** löschen
 '         Cells(KZeile, 2) = ""
 '         meldung = meldung & Chr(10) & _
 '         "KontoKennDat hat in Kontoblatt " & AKtoBlatt & _
 '         " Zeile " & KZeile & " falsches ''***'' gelöscht."
 '       End If
 '     Next KZeile
 '     AKto3SternZeile = DSZ
 '     Cells(1, 1) = DSZ
 '     GoTo DSZrichtig
 '   Else
 '     GoTo KeineDreisternZeile
 '   End If
 ' End With 'Sheets(AKtoBlatt)
'DSZinA1Falsch:
'    meldung = meldung & Chr(10) & _
'    "In Kontoblatt " & AKtoBlatt & " ist in Zelle A1 die Zeilenangabe " & _
'    Sheets(AKtoBlatt).Cells(1, 1) & " für das ***-Zeichen falsch"
'    GoTo KeineDreisternZeile
'DSZselbstFalsch:
'    meldung = "DSZselbstFalsch"
'    GoTo KeineDreisternZeile
'Eindeutigkeit der TA- und der BuID-Spalten prüfen ----------
DSZrichtig:
'--------------- Endsaldo ---------------
  AktoEndSaldo = Cells(Cells(1, 1) + 2, 6)
'------------- Verzichten auf Eindeutigkeitsprüfung wegen Zeitaufwand --------
  GoTo Ausgang
'-----------Prüfung auf Eindeutigkeit der Ta-Nr (und der BuID) -----------
  Dim Z As Long, VerglZ As Long, Prüfling1 As Long, Prüfling9 As Long
  Dim BeginnZeile As Long
  
  With Sheets(AKtoBlatt)
    .Activate
    Cells(6, 1).Activate
    BeginnZeile = ActiveCell.Row
    For Z = BeginnZeile To AKto3SternZeile
      If Cells(Z, 1) <> "" And Cells(Z, 9) <> "" Then
        Cells(Z, 1).Activate    'für Test
        Prüfling1 = Cells(Z, 1)
        Prüfling9 = Cells(Z, 9)
        For VerglZ = Z + 1 To AKto3SternZeile
          Cells(VerglZ, 1).Activate           'für Test
          If Cells(VerglZ, 1) = Prüfling1 Or _
             Cells(VerglZ, 9) = Prüfling9 Then
    MELDUNG = MELDUNG & Chr(10) & _
    "Im Kontenblatt " & AKtoBlatt & " haben die beiden Zeilen " & Z & _
    " und " & VerglZ & " die gleiche Transaktionsnummer " & Cells(Z, 1) & Chr(10) & _
    "Dieses und die anderen an der Transaktion beteiligten Kontoblätter" & _
    Chr(10) & "spätestens vor Berichterstellungen korrigieren!"
            KKDAbbruch = True
            GoTo Ausgang    'keine weiteren Prüfungen
          End If
        Next VerglZ
      End If
'      Z = Z + 1
      If Z > AKto3SternZeile Then
        GoTo Ausgang
      End If
    Next Z
  End With 'Sheets(AKtoBlatt)
Ausgang:
  If KKDAbbruch = True Then
    ABBRUCH = True    'Für andere Programme verständlich
  End If
  With Worksheets(AktBlatt) 'Wiederherstellen Aufrufsituation ---------------
    .Activate
    Cells(AZrow, AZcolumn).Activate
  End With
End Sub 'KtoKennDat

Function SpalteEindeutig(Spalte As Integer, GuteDSZ As Long) As Boolean
'Prüft von der aktiven Zelle des Aktiven Blatts aus abwärts bis zu
'der als (geprüft vorausgesetzten) Dreisternzeile, ob in der Spalte
'der aktiven Zelle eine Zahl (Long) wiederholt vorkommt. Die Zeile, bei
'der dies auftritt, wird in der Globalvariablen DoppelInZeile abgelegt
'(die sonst 0 ist) und SpalteEindeutig false gesetzt.
  Dim Z As Long, Z2 As Long, Prüfling As Long
  Dim BeginnZeile As Long, BeginnZ As Long, DoppelInZeile As Integer
  With ActiveSheet
    BeginnZeile = ActiveCell.Row
    Spalte = ActiveCell.Column
    DoppelInZeile = 0
    
    BeginnZ = BeginnZeile
    For Z = BeginnZ To GuteDSZ
LeerZeile:
      If Cells(Z, Spalte) = "" Or Cells(Z, Spalte) = "BuID" _
                               Or Cells(Z, Spalte) = "TA" Then
        Z = Z + 1            'Leerzeilen oder andere überspringen
        If Z < GuteDSZ Then GoTo LeerZeile
      Else
        Prüfling = Cells(Z, Spalte)
        BeginnZ = Z + 1
        For Z2 = BeginnZ To GuteDSZ
LeerZeile2:
          If Cells(Z2, 9) = "" Or Cells(Z2, Spalte) = "BuID" _
                               Or Cells(Z2, Spalte) = "TA" Then
            Z2 = Z2 + 1
            If Z2 < GuteDSZ Then GoTo LeerZeile2
          Else
            If Cells(Z2, Spalte) = Prüfling Then
              DoppelInZeile = Z2
              GoTo DoppelGefunden
            End If
          End If
        Next Z2
      End If
    Next Z
    SpalteEindeutig = True
    GoTo FunctionEnd
DoppelGefunden:
    SpalteEindeutig = False
    GoTo FunctionEnd
FunctionEnd:
  End With 'ActiveSheet
End Function 'SpalteEindeutig

'======================================im Modul Kontenplanpflege =================
Sub KontoBlattEinrichten(KontoNr As Integer)
' Setzt einen vorangegangenen KtoKennDat-Aufruf voraus und benutzt dessen
' Akto-Daten.
' Erstellt mit dem Blatt "KntoVorlage" als Muster ein neues Kontoblatt mit
' den Werten, die in der Zeile der aktiven Zelle stehen und fügt es hinter
' dem letzten schon erstellten Blatt ein, das im Kontenplan dieser Zeile
' vorausgeht.  Setzt voraus, daß das Blatt "Kontenplan" und darin eine Zelle
' in der Spalte "Konto" aktiv ist. Wenn von Function Kontoblatt aufgerufen,
' ist dies sichergestellt.

Dim StartBlatt As String, StartZeile As Integer, StartSpalte As Integer
Dim VorausgehBlatt As String  ', KontoNr As Integer
Dim Ein As VbMsgBoxStyle, I As Integer
'AKtoStatus: KtoUnbekannt = 0, KtoBlattFehlt = 1, KtoGanzLeer = 2, _
'            KtoGeleert = 3, KtoLeerMitÜbertrag = 4, KtoHatBuchungen = 5, _
'            KtoHatBuchungenUndÜbertrag = 6
'1 KBE ------------------------- Anfangsbedingungen -----------------------------
  ABBRUCH = False
  StartBlatt = ActiveSheet.Name
  StartZeile = ActiveCell.Row
  StartSpalte = ActiveCell.Column
'  KontoNr = Cells(StartZeile, StartSpalte)
'2 KBE -------------------- ABZAbbruch schon gemeldet? ---------------------------
  If ABBRUCH = True Then
    MELDUNG = MELDUNG & Chr(10) & _
    "KontoBlattEinrichten übersprungen"
    GoTo EndeBlattEinrichten
  End If
'3 KBE -------------------- Blatt definiert? ---------------------------
  If AKtoStatus = 0 Then
    MELDUNG = MELDUNG & Chr(10) & _
    "Das Konto " & AKtoNr & " ist unbekannt. Es muss erst mit" & Chr(10) & _
    "''Strg+k / EINFÜGEN'' im Kontenplan definiert werden."
    ABBRUCH = True
    GoTo EndeBlattEinrichten
  End If
'3 KBE ---------------- Blatt schon eingerichtet? -------------------------  If AKtoStatus >= KtoGanzLeer Then      'KtoGanzLeer = 2
  If Sheets("Kontenplan").Cells(AKtoKPZeil, 1) = "E" Then 'zur Sicherheit
    MELDUNG = MELDUNG & Chr(10) & _
    "Das Konto " & AKtoNr & " hat laut Kontenplan bereits ein Blatt eingerichtet"
    GoTo EndeBlattEinrichten
  End If
'3 KBE -------------------- Blatt erzeugen, positionieren --------------------
  With Sheets("Kontenplan")
    .Activate
    If AKtoStatus = KtoBlattFehlt Then
      If AKtoKPZeil <= 5 Then
        VorausgehBlatt = "ArProt"
        GoTo Positionieren
      End If
    End If
    If AKtoKPZeil > 5 Then
      For I = AKtoKPZeil To 4 Step -1
        If Cells(I, KPCE).Value = "E" Then
          VorausgehBlatt = Cells(I, KPCBlattname).Value
          GoTo Positionieren
        End If
      Next I
    End If
    VorausgehBlatt = "ArProt"
Positionieren:
    If AKtoBlatt <> "" Then
      Sheets(VorausgehBlatt).Select
      ActiveWindow.ScrollWorkbookTabs Sheets:=1
      Sheets("KntoVorl").Copy after:=Sheets(VorausgehBlatt)
      Sheets("KntoVorl (2)").Select
      Sheets("KntoVorl (2)").Name = AKtoBlatt
      Sheets("Kontenplan").Cells(AKtoKPZeil, 1) = "E"
      AKtoStatus = KtoGanzLeer
    End If
  End With 'Sheets("Kontenplan")
'4 KBE ------------------------ Blatt mit Fetinfos füllen ---------------
    If AKtoBlatt <> "" Then
    Call KtoKennDat(KontoNr)  'Aufruf oben gilt noch bis auf AKtoStatus
      With Sheets(AKtoBlatt)
        Cells(1, 3) = AKtoArt
        Cells(1, 4) = AKtoNr
        Cells(1, 5) = AKtoBeschr
        Cells(3, 5) = AKtoStrasse
        Cells(4, 6) = AktoÜbertrag
      End With
    End If
'5 KBE---------- Kontenplan-Versionszähler inkrementieren ------------
  With Sheets("Kontenplan")
    .Activate
    Cells(1, 1) = Cells(1, 1) + 1
    If Cells(1, 1) > 99 Then
      Cells(1, 1) = 1
    End If
  End With
  GoTo EndeBlattEinrichten
EndeBlattEinrichten:
  Sheets(StartBlatt).Activate
  Cells(StartZeile, StartSpalte).Activate
End Sub 'KontoblattErstellen'=================================================================================

Sub SucheKonto(SuchText As String)
'Aufruf von ERFASSEN aus mit ArProt als aktivem Blatt und eine Zelle in der Sollkonto-
'oder Habenkontospalte aktiv. Sucht im Kontenplan in der Spalte "Beschreibung"
'nach dem String "SuchText", von der ersten Zeile an bis BBZZEnde. Bietet in einer
'MessageBox das gefundene Konto an mit der Wahl, es zu akzeptieren, weiterzusuchen
'oder abzubrechen. Wird das Konto akzeptiert, schreibt SucheKonto die Kontonummer
'in die aktive Zelle (des Blattes ArProt)und aktiviert die rechts daneben liegende
'Zelle.

Dim Länge As Integer, AktBlatt As String, AktZell As Range, Zeile As Integer
Dim Kto As Integer, Becr As String, Blat As String, Art As Integer, A
Dim Status As Integer, SuchAnfZeile As Integer
  With ActiveSheet          'Aufbewahren Aufrufsituation
    AktBlatt = ActiveSheet.Name
    Set AktZell = ActiveCell
  End With                  '---------------------------
  Länge = Len(SuchText)
  SuchAnfZeile = 4
Such:
  With Worksheets("Kontenplan")
    .Activate
    For Zeile = SuchAnfZeile To Cells(1, 3) - 1
      If Left(Cells(Zeile, KPCBlattname), Länge) = SuchText Then
        Exit For
      End If
    Next Zeile
    SuchAnfZeile = Zeile + 1
    If Zeile <= Cells(1, 3).Value Then
      Kto = Cells(Zeile, KPCKonto)
      Call KtoKennDat(Kto)
      Becr = AKtoBeschr
      Blat = AKtoBlatt
      Art = AKtoArt
      Status = AKtoStatus
    Else
      Kto = 0
      Becr = ""
      Blat = ""
      Art = 0
    End If
  End With
  With Worksheets(AktBlatt) 'Wiederherstellen Aufrufsituation
    .Activate
    AktZell.Activate
    If Kto <> 0 Then
      A = MsgBox("     " & Kto & "    " & Becr & "     " & Blat & Chr(10) & Chr(10) & _
            "Schaltfläche ''Nein'', falls weitergesucht werden soll", 35, _
            "Ist das Konto gemeint?")
      If A = vbCancel Then Exit Sub
      If A = vbYes Then
        ActiveCell.Value = Kto
        ActiveCell.Offset(0, 1).Activate
        If ActiveCell.Column = APCText And (Art = MitgliedKto Or Art = SpenderKto) Then
          If Art = MitgliedKto Then
            Sheets("ArProt").Cells(ActiveCell.Row, APCText) = _
            "Beitrag " & Sheets("Kontenplan").Cells(Zeile, KPCBeschr)
          End If
          If Art = SpenderKto Then
            Sheets("ArProt").Cells(ActiveCell.Row, APCText) = _
            "Spende " & Sheets("Kontenplan").Cells(Zeile, KPCBeschr)
          End If
          ActiveCell.Offset(0, 1).Activate
        End If
        Exit Sub
      End If
      If A = vbNo Then
        GoTo Such
        End If
      End If
    If Kto = 0 Then
      A = MsgBox("Mit einer anderen Zeichenfolge versuchen" & Chr(10) & Chr(10), 0, _
                 "Kein Konto gefunden")
    End If
  End With                  '
End Sub  'SucheKonto
'=================================================================================

'==================================================================================

Sub KopfZeilenÜberSchrift(AktBlatt As String)
  With Sheets("Kontenplan")
    .Activate
    LinkerHeader = Cells(1, 9) '.Value
    RechterHeader = Cells(1, 11) '.Value
  End With 'Sheets("Kontenplan")
  With Sheets(AktBlatt)
   .Activate
   .PageSetup.LeftHeader = LinkerHeader
   .PageSetup.RightHeader = RechterHeader
  End With
End Sub 'KopfZeilenÜberSchrift

'Modul KontPlanBearb =================== Ende =====================================


