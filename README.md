# excelaccountant
Software, basierend auf VBA Excel Makros für die doppelte Buchfürung. Ursprünglich um 2000 herum für die Arbeit des Schatzmeisters eines Vereins entwickelt. Mehrere Personen haben über die Jahre "ihren Senf" dazugegeben. Inzwischen ist es von einem einfachen Makro zu einem recht komplizierten Konstrukt gewachsen. Die Arbeit hat es stets erfüllt. Im Zuge der Modernisierung wird es nun nicht mehr gebraucht. Zum Löschen ist es zu Schade. Eventuell kann jemand was damit anfangen. Viel Spass :)


Funktion:
Mit dem Buchungsprogramm ExAcc kann der verantwortliche Schatzmeister 
(Anwender) die Bewegungen des Finanzkörpers einfach erfassen und übersichtlich darstellen. 
Der Finanzkörper ist organisiert in eine Anzahl von Konten, die der Herkunft, der Verwendung oder 
dem Verbleib von  Geldern zugeordnet sind. Eine Bewegung im Finanzkörper, Transaktion genannt, 
ist nach der Methode der doppelten Buchführung, die ExAcc verwendet, immer der Übertrag eines   
Betrages von einem für die Anwendung definierten Herkunftskonto (Habenkonto) in ein 
ebensolches Zielkonto (Sollkonto). Die Buchungen beschränken sich auf ein Kalenderjahr und sind
innerhalb des Jahres nach Monaten gegliedert.

Makros:
ExAcc baut auf den Funktionen des Tabellenkalkulationsprogramms Excel auf, die  
durch sogenannte Makros ergänzt sind. Ein Makro fasst eine (manchmal längere) Reihe von 
Tabellenoperationen zusammen, die beim Buchen, der Eingabe einer Transaktion, erforderlich sind.  
Ein Makro wird durch eine Tastendruck-Kombination geweckt, erleichtert so dem Schatzmeister  
das Buchen, gewährleistet die Vollständigkeit der Eingabeoperation und erübrigt tiefe
Kenntnis von Excel. 

Mappen:
Eine ExAcc-Anwendung besteht aus zwei Excel-Mappen. 
Die eine Mappe ist die ExAcc-Mappe (System-Mappe), bestehend aus den Makros 
und generischen Blättern, die der anderen Mappe, der Anwendungsmappe, im Verlauf der 
Buchung als Vorlagen dienen. Beim Buchen müssen beide Mappen geöffnet sein, jedoch nur die 
Anwendungsmappe im aktiven Fenster (im Vordergrund) stehen. Zugriffe auf die Systemmappe
geschehen unbemerkt für den Anwender. Die Anwendungsmappe erhält einen Namen, den  
der Anwender bestimmt. z.B. "Buch 2025.xls"

Blätter (Tabellen):
Die Anwendungsmappe enthält als erste Tabelle (als erstes Blatt der Mappe) den Kontenplan. 
Er enthält alle Angaben, die für die Anwendung erforderlich sind und durch die sie definiert ist.  
Die zweite Tabelle (das zweite Blatt der Mappe) ist das Arbeitsprotokoll. Dies ist das Blatt, in dem      
der Schatzmeister seine Buchungen einträgt. In die anderen Blätter der Anwendungsmappe wird nur
nur über Makroaufrufe von ExAcc aus eingetragen. Der Schatzmeister kann sich die anderen Tabellen   
(die anderen Blätter) der Anwendungsmappe anschauen, darf sie aber nicht verändern, wenn er
die Datenkonsistenz nicht gefährden will. Sie zeigen Konten, Saldenlisten, Berichte u.a.
Erstellt wird ein Kontoblatt mittels der Angaben im Kontenplan erst bei der ersten Benutzung, ein.
Berichts- oder Saldenblatt erst durch Anforderung eines Berichts.

Anpassung an die Buchungsaufgabe:
Der Benutzer muss zuerst den Kontenplan erstellen in dem Format, das im Blatt Kontenplan 
vorgegeben ist, oder ihn beim Jahreswechsel von der Vorjahresbuchung  übernehmen.  
Der Kontenplan ist in mehrere Bereiche eingeteilt mit verschiedenen Kontoarten, die von ExAcc 
verschieden behandelt werden (siehe Blatt Kontoarten)..   
Der Kontenplan enthält alle für das Buchungsprojekt (Applikation) spezifischen Informationen  
und konfiguriert die Anwendermappe (passt sie an die vorliegende Buchungsaufgabe an).
ExAcc erzeugt die notwendigen Blätter bei der ersten Benutzung.
Die Vorlagen SaldLiVorl, BerichtVorl und FABestVorl sind zu überprüfen, ob ihre Texte der 
Aufgabe entsprechend richtig formuliert sind.  Die Texte können nach Bedarf abgeändert 
werden, wenn dabei die Struktur des Blatts erhalten bleibt.  Das Arbeitsprotokoll ist das Tabellenblatt, 
mit dem alle Buchungsaufgaben durchgeführt werden, nachdem das System im Blatt Kontenplan 
konfiguriert wurde. 

Ablauf einer Buchung:
Die Benutzung von ExAcc geschieht hauptsächlich durch Buchen im Blatt Arbeitsprotokoll. Der Buchungs- 
Vorgang besteht im Ausfüllen der anstehenden Zeile (der letzten Zeile des Protokolls), und Auslösung   
der Buchung durch Drücken der Tastenkombination (TK) Strg+b  am Ende der Zeile (Spalte H) .
In diese anstehende Arbeitsprotokoll-Zeile schreibt der Benutzer rechts von der  fortlaufenden,
von ExAcc vorgegebenen Transaktionsnummer (in der ersten Spalte, Spalte A) 
Transaktionsdatum,  Belegkennzeichen, Sollkonto, Habenkonto, eine textliche Beschreibung 
der Buchung und den  Betrag.  Anschließend aktiviert er die Zelle mit dem von ExAcc vorgegebenen
Buchungskennzeichen  *** (in manchen Fällen  ****  oder  *****) in Spalte "gebucht", Spalte H,
und drückt die Tastenkombination Strg+b.  ExAcc überträgt daraufhin den genannten Betrag in der
Reihenfolge des genannten Transaktionsdatums (also chronologisch) von dem genannten Haben-
Konto in das genannte Soll-Konto und fügt noch einige Daten hinzu,  die der internen Sicherung der 
Datenkonsistenz dienen, oder als Querverweise dem Buchprüfer die Arbeit erleichtern.
Für Buchungen, an denen mehr als zwei Konten beteiligt sind (Buchungen 
mit Sammelkonten, oder nicht realen Konten wie durchlaufenden Posten, Fonds) verlangt ExAcc  im
Dialog die zusätzlich benötigte Information, wenn sie nicht schon im Kontenplan in der Spalte 
 SamlKto (Spalte E) gegeben ist.   
Außer durch explizites Schreiben der Daten in die Buchungszeile können Eingabehilfen durch
Drücken der Taste b  bei gedrückter Taste Strg veranlasst werden, die das Buchen beschleunigen.
(siehe auch Blatt "ShortCuts")
Aber auch vorher, d.h. während des  Ausfüllens der Buchungszeile, bieten die mit der TK  Strg+b
angesteuerten Funktionen vielfache Erleichterung. Die Wirkung der TK  Strg+b  ist kontextsensitiv,
hängt also von der Spalte der aktiven Zelle ab. Welche Wirkung zu erwarten ist, kann mittels der 
TK Strg+h oder einfach durch ausprobieren ermittelt werden. Das Ausprobieren
ist in allen Zellen außer den in Spalte H  folgenlos in dem Sinne, dass die in der Zelle erscheinende
Information mit der Entfernen-Taste wieder gelöscht werden oder manuell überschrieben werden
kann und erst durch Starten der Buchung seine Wirkung entfaltet. 

Berichte:
Aus den durch die Buchungen erzeugten Kontoblättern können auf Anforderung folgende Berichte 
erstellt werden:  
Kontenstandstabelle
Eine Auflistung der Monatsendstände aller Sachkonten von Jahresbeginn bis zum aktuellen
Buchungsstand
Summen- und Saldenliste
Auflistung der Anfangs- und Endstände der Sachkonten und ihre Änderung für eine wählbare, 
nach Monatsschnitten strukturierte Periode; mit Kontrollrechnung 
Kurz-Bericht
Darstellung des Bestands, der Abgänge und Zugänge nach Kontoarten getrennt in komprimierter  
Form für Sachkonten einer wählbaren, nach Monatsschnitten strukturierten Periode.
Finanzamt-Bescheinigungen
 Für Anwendungen mit Personenkonten:  Personenbezogene formgerechte Beitragsbescheinigungen 
für das Kalenderjahr differenziert nach Mitgliederbeiträgen und Spenden

Die ExAcc-Funktionen werden mit Tastenkombinationen "Strg+Buchstabe" aktiviert; 
siehe Blatt "ShortCuts".
