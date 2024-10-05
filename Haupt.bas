Attribute VB_Name = "Haupt"
Option Explicit
Public FNr&
Public Const vNS$ = vbNullString
Public Const MTBeg% = 1
Public Enum TbTyp
 tbma = MTBeg ' mitarbeiter
 tbdp ' dienstplan
 tbwp ' wochenplan
 tbar ' arten
 tbpr ' protokoll
 tbbi ' bilanzen
 tbul ' userlist
 tbei ' einstellungen
 tbab ' ausbez
 tbende
End Enum
Public Const Az0% = 25
Public Enum azgtyp ' für MfGTyp
 azgnix = Az0 ' nix anzeigen
 azgdp ' Dienstplan
 azgma ' Mitarbeiter
 azgar ' Arten
 azgwp ' Wochenplan
 azgpr ' Protokoll
 azgul ' userlist
 azgab ' Ausbezahlungen
 azgende
End Enum
Public FPos&
Public Const obdebug% = False
'Public Const ConString$ = "DRIVER={MySQL ODBC 5.1 Driver};server=linux;uid=praxis;pwd=...;database=dp;option=" & opti
'Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rsc As New ADODB.Recordset
Public catx As New ADOX.Catalog
Public AltInhalt$
Const LVobMySQL% = True
Public Const liName$ = "linux1"
Public pVerz$

'Struktur für Feiertage
Public Type feiert
    name As String
    KuNa As String
    Datum As Date
    obhalb As Boolean
    wday As Integer
End Type

'alle Feiertage und Halbfeiertage des Jahres in München
Public ftag(17) As feiert
Type HSL
      Hue As Double     '    As Long
      Saturation As Double  '  As Long
      Luminance As Double '   As Long
End Type

Public Type dpSatz
 Tag As Date
 ausbez As Double
 urlhaus As Double
 artnr As String
 Stdn As Double
End Type ' dpSatz

Const GradZ% = 240 ' wird im CommonDialog verwendet, in der Algorithmusquelle: 100
Const GradZC% = 240 ' wird im CommonDialog verwendet, in der Algorithmusquelle: 360
Public uVerz$ ' Ausgangsverzeichnis für Dateisuche bei  DBVerb
Public Const doppelteWeg$ = "DELETE FROM dienstplan WHERE EXISTS(SELECT 0 FROM dienstplan d WHERE d.tag=dienstplan.tag AND d.persnr=dienstplan.PersNr AND d.artnr=dienstplan.artnr AND d.id<dienstplan.id)"

Public obStart% ' ob Startvorgang (da in ConstrFestleg DateiÖffnen-Dialog für die MDB-Datei zeigen und dbcn nicht verbinden)


#If doppelt Then
Public Function syscmd(art%, Optional Inhalt$)
 On Error Resume Next
 Select Case art
  Case 4 ' acSysCmdSetStatus
   MDI.cnLab = Inhalt
'   Forms(0).Überschrift = Inhalt
  Case 5 ' acSysCmdClearStatus
   MDI.cnLab = vNS
'   Forms(0).Überschrift = vNS
 End Select
' Debug.Print Inhalt
 Err.Clear
 DoEvents
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in syscmd/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' syscmd(art%, Optional Inhalt$)
#End If

Function min(a, b)
 If a < b Then min = a Else min = b
End Function ' min(a,b)

Function MAX(a, b)
 If a > b Then MAX = a Else MAX = b
End Function ' max(a,b)

Public Function RGBToHSL02(ByVal RGBValue As Long) As HSL
' by Donald (Sterex 1996), donald@xbeat.net, 20011116
  Dim r As Long, g As Long, b As Long
  Dim lMax As Long, lMin As Long
  Dim q As Single

  r = RGBValue And &HFF
  g = (RGBValue And &HFF00&) \ &H100&
  b = (RGBValue And &HFF0000) \ &H10000

  If r > g Then
    lMax = r: lMin = g
  Else
    lMax = g: lMin = r
  End If
  If b > lMax Then
    lMax = b
  ElseIf b < lMin Then
    lMin = b
  End If

  RGBToHSL02.Luminance = lMax * GradZ * 0.5 / 255
  
  If lMax > lMin Then
    RGBToHSL02.Saturation = (lMax - lMin) * GradZ / lMax
    q = GradZC / 6 / (lMax - lMin)
    Select Case lMax
    Case r
      If b > g Then
        RGBToHSL02.Hue = q * (g - b) + GradZC
      Else
        RGBToHSL02.Hue = q * (g - b)
      End If
    Case g
      RGBToHSL02.Hue = q * (b - r) + 1 / 3 * GradZC
    Case b
      RGBToHSL02.Hue = q * (r - g) + 2 / 3 * GradZC
    End Select
  End If
  If obdebug Then Debug.Print RGBToHSL02.Hue, RGBToHSL02.Luminance, RGBToHSL02.Saturation
End Function ' RGBToHSL02
Public Function Dunkler&(Farbe&, Faktor!)
 Dim zwi As HSL
 zwi = RGBToHSL02(Farbe)
 zwi.Luminance = zwi.Luminance - (zwi.Luminance - 0) * ((Faktor - 1) / Faktor)
' Dunkler = HSLToRGB02a(Zwi.Hue * 360 / GradZC, Zwi.luminance  * 100 / GradZ, Zwi.Luminance * 100 / GradZ)
 Dunkler = HSL_to_RGB((zwi.Hue), (zwi.Saturation), (zwi.Luminance))
End Function ' Dunkler&(Farbe&, Faktor!)

Public Function Heller&(Farbe&, Faktor!)
 Dim zwi As HSL
 zwi = RGBToHSL02(Farbe)
#If False Then
 Dim r#, b#, g#, z#
 r = Farbe Mod 256
 g = ((Farbe - r) / 256) Mod 256
 b = (((Farbe - r) / 256) - g) / 256 Mod 256
 Call RGB_to_HSL(r / 256, g / 256, b / 256, zwi.Hue, zwi.Saturation, zwi.Luminance)
 zwi.Hue = zwi.Hue * GradZC
 zwi.Luminance = zwi.Luminance * GradZ
 zwi.Saturation = zwi.Saturation * GradZ
#End If
 zwi.Luminance = zwi.Luminance + (GradZ - zwi.Luminance) * ((Faktor - 1) / Faktor)
' If zwi.Luminance + (GradZ - zwi.Luminance) * ((Faktor - 1) / Faktor) <> zwi.Luminance Then Stop
' Heller = HSLToRGB02a(Zwi.Hue * 360 / GradZC, Zwi.Saturation * 100 / GradZ, Zwi.Luminance * 100 / GradZ)
 
 Heller = HSL_to_RGB((zwi.Hue), (zwi.Saturation), (zwi.Luminance))
End Function ' Heller&(Farbe&, Faktor!)

Public Function RGB_to_HSL(r#, g#, b#, h#, s#, l#)
'RGB_to_HSL  (r,g,b,h,s,l)
    Dim V#, m#, vm#
    Dim r2#, g2#, b2#
    Dim zwi#

    V = MAX(r, g)
    V = MAX(V, b)
    m = min(r, g)
    m = min(m, b)
    l = (m + V) / 2#
    If l <= 0 Then Exit Function
    vm = V - m
    s = vm
    If s > 0 Then
     If l < 0.5 Then
      zwi = V + m
     Else
      zwi = 2# - V - m
     End If
     s = s / zwi
    Else
     Exit Function
    End If

    r2 = (V - r) / vm
    g2 = (V - g) / vm
    b2 = (V - b) / vm

    If r = V Then
     If g = m Then
      h = 5 + b2
     Else
      h = 1# - g2
     End If
    ElseIf g = V Then
     If b = m Then
      h = 1# + r2
     Else
      h = 3# - b2
     End If
    Else
     If r = m Then
      h = 3# + g2
     Else
      h = 5# - r2
     End If
     h = h / 6
    End If
End Function ' RGB_to_HSL

Public Function HSL_to_RGB(h&, s&, l&)
 Dim r#, g#, b#
 Dim V#
    If l + l <= GradZ Then
     V = l * (1 + s / GradZ)
    Else
     V = l + s - l * s / GradZ
    End If
    V = V / GradZ
    If V <= 0 Then
     r = 0
     g = 0
     b = 0
    Else
     Dim m#, sV#
     Dim sextant%
     Dim fract#, vsf#, mid1#, mid2#
     m = l / GradZ + l / GradZ - V
     sV = (V - m) / V
     h = h * 6
     sextant = Int(h / GradZC)
     fract = h / GradZC - sextant
     vsf = V * sV * fract
     mid1 = m + vsf
     mid2 = V - vsf
     Select Case sextant
            Case 0: r = V: g = mid1: b = m
            Case 1: r = mid2: g = V: b = m:
            Case 2: r = m: g = V: b = mid1
            Case 3: r = m: g = mid2: b = V
            Case 4: r = mid1: g = m: b = V
            Case 5: r = V: g = m: b = mid2
    End Select
   End If
   HSL_to_RGB = RGB(Int(r * 255), Int(g * 255), Int(b * 255))
End Function ' HSL_to_RGB(h&, s&, l&)

Function FTbeleg(Jahr%) As Boolean
'Feiertage werden mit Daten belegt
 Dim OstDat As Date     'Tag und Monat des Ostertermins
 Dim i%
 If Jahr < 1901 Or Jahr > 2078 Then ' Gültigkeitszeitraum der Formel
   MsgBox "Ostern " & Str(Jahr) & " läßt sich nicht sicher errechnen. Breche ab"
   FTbeleg = False
   Exit Function
 End If
 OstDat = OsterDatum(Jahr)
 Call do_FTbeleg(Jahr, OstDat)
 For i = 0 To 17
  ftag(i).wday = Weekday(ftag(i).Datum)
 Next
 FTbeleg = True     'Feiertagsbelegung erfolgreich
End Function ' FTbeleg(jahr%) As Boolean

Function do_FTbeleg(Jahr%, Ostern As Date)
'Feiertage werden mit Namen und Datum belegt, wenn Ostern bekannt ist (Mondkalender fehlt im Programm)
 On Error GoTo fehler
 Dim jahrstr$
 jahrstr = Str(Jahr)
    ftag(0).KuNa = "NJ": ftag(0).name = "Neujahr": ftag(0).Datum = CDate("1.1." + jahrstr)
    ftag(1).KuNa = "Hl3": ftag(1).name = "Hl.3 Könige": ftag(1).Datum = CDate("6.1." + jahrstr)
    ftag(2).KuNa = "Fasc": ftag(2).name = "Faschingsdi": ftag(2).Datum = Ostern - 47: ftag(2).obhalb = True
    ftag(3).KuNa = "Karf": ftag(3).name = "Karfreitag": ftag(3).Datum = Ostern - 2
    ftag(4).KuNa = "OSo": ftag(4).name = "Ostersonntag": ftag(4).Datum = Ostern
    ftag(5).KuNa = "OMo": ftag(5).name = "Ostermontag": ftag(5).Datum = Ostern + 1
    ftag(6).KuNa = "Maif": ftag(6).name = "Maifeiertag": ftag(6).Datum = CDate("1.5." + jahrstr)
    ftag(7).KuNa = "ChHF": ftag(7).name = "Christ.Hlf.": ftag(7).Datum = Ostern + 39
    ftag(8).KuNa = "PfiS": ftag(8).name = "Pfingstsonn": ftag(8).Datum = Ostern + 49
    ftag(9).KuNa = "PfiM": ftag(9).name = "Pfingstmont": ftag(9).Datum = Ostern + 50
    ftag(10).KuNa = "Frlm": ftag(10).name = "Fronleichnm": ftag(10).Datum = Ostern + 60
    ftag(11).KuNa = "MaHF": ftag(11).name = "Mariä Hlft.": ftag(11).Datum = CDate("15.8." + jahrstr)
    ftag(12).KuNa = "TdDE": ftag(12).name = "Tag d.dt.Ei": ftag(12).Datum = CDate("3.10." + jahrstr)
    ftag(13).KuNa = "AllH": ftag(13).name = "Allerheilig": ftag(13).Datum = CDate("1.11." + jahrstr)
    ftag(14).KuNa = "HlAb": ftag(14).name = "Heilig.Aben": ftag(14).Datum = CDate("24.12." + jahrstr)
    ftag(15).KuNa = "1.WF": ftag(15).name = "1.Weihnacht": ftag(15).Datum = CDate("25.12." + jahrstr)
    ftag(16).KuNa = "2.WF": ftag(16).name = "2.Weihnacht": ftag(16).Datum = CDate("26.12." + jahrstr)
    ftag(17).KuNa = "Silv": ftag(17).name = "Silvester": ftag(17).Datum = CDate("31.12." + jahrstr)
    If Jahr < 91 Or (Jahr > 1970 And Jahr < 1990) Then ftag(17).obhalb = True
    Exit Function
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in do_ftbeleg/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function 'do_FTbeleg(jahr%, Ostern As Date)

Function OsterDatum(Jahr%) As Date
' Berechnet Ostertermine von 1901 -2078, nach Gauß
 Dim a%, b%, c%, d%, e%, Tag%, monat%
 a = Jahr Mod 19
 b = Jahr Mod 4
 c = Jahr Mod 7
 d = (19 * a + 24) Mod 30
 e = (2 * b + 4 * c + 6 * d + 5) Mod 7
 Tag = 22 + d + e
 monat = 3
 If Tag > 31 Then
    Tag = d + e - 9
    monat = 4
 ElseIf Tag = 26 And monat = 4 Then
    Tag = 19
 ElseIf Tag = 25 And monat = 4 And d = 28 And e = 6 And a > 10 Then
    Tag = 18
 End If
 OsterDatum = DateSerial(Jahr, monat, Tag)
End Function ' OsterDatum(Jahr%) As Date

Sub ViewsErstellen()
 Dim sql$
 sql = "DROP FUNCTION IF EXISTS `obft`"
 Call DBVerb.cnVorb("dp", tbm(tbwp), tbm(tbdp)) '("", "--multi", vns)
 myEFrag (sql)
 sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `obft`(dt DATE) RETURNS FLOAT" & Chr$(13) & _
    "NO SQL DETERMINISTIC" & vbCrLf & _
    "BEGIN " & Chr$(13) & _
    " DECLARE a,b,c,d,e,tag,monat,jahr INT;" & vbCrLf & _
    " DECLARE erg DATE;" & vbCrLf & _
    " SET jahr=YEAR(dt);" & vbCrLf & _
    " SET a=jahr MOD 19; SET b=jahr MOD 4; SET c=jahr MOD 7;" & vbCrLf & _
    " SET d=(19*a+24)MOD 30; SET e=(2*b+4*c+6*d+5)MOD 7; SET Tag=22+d+e; SET monat=3;" & vbCrLf & _
    " IF Tag>31 THEN SET Tag=d+e-9; SET monat=4; ELSEIF Tag=26 && monat=4 THEN SET Tag=19;" & vbCrLf & _
    " ELSEIF Tag=25&&monat=4&&d=28&&e=6&&a>10 THEN SET Tag=18; END IF;" & vbCrLf & _
    " SET erg = STR_TO_DATE(CONCAT(jahr,'-',monat,'-',Tag),'%Y-%m-%d');" & vbCrLf & _
    " SET tag=DAY(dt);" & vbCrLf & _
    " SET monat=MONTH(dt);" & vbCrLf & _
    " IF (tag=1&&monat IN(1,5,11))||(tag=6&&monat=1)||(tag=15&&monat=8)||(tag=3&&monat=10)||(tag IN(24,25,26)&&monat=12)||DATEDIFF(dt,erg)IN(-2,0,1,39,49,50,60)THEN RETURN 1; END IF;" & vbCrLf & _
    " IF DATEDIFF(dt,erg)=-47 THEN RETURN 0.5; END IF;" & vbCrLf & _
    " RETURN 0;" & vbCrLf & _
    "END "
    myEFrag (sql)
End Sub ' Viewserstellen


Public Function tbm$(akttb As TbTyp)
 Select Case akttb
  Case tbma: tbm = "mitarbeiter"
  Case tbdp: tbm = "dienstplan"
  Case tbwp: tbm = "wochenplan"
  Case tbar: tbm = "arten"
  Case tbpr: tbm = "protok"
  Case tbbi: tbm = "bilanzen"
  Case tbul: tbm = "user"
  Case tbei: tbm = "einstellungen"
  Case tbab: tbm = "ausbez"
 End Select ' Case akttb
End Function ' tbm$(mfgtyp As TbTyp)

Public Function azm$(ztyp As azgtyp)
 Select Case ztyp
  Case azgnix: azm = "Nix"
  Case azgdp: azm = "Dienstplan"
  Case azgma: azm = "Mitarbeiter"
  Case azgar: azm = "Arten"
  Case azgwp: azm = "Wochenplan"
  Case azgpr: azm = "Protokoll"
  Case azgul: azm = "Userlist"
 End Select ' Case ztyp
End Function ' azm$(ztyp As AzgTyp)

Public Function azt(az As azgtyp) As TbTyp ' Anzeige zu Tabelle
 Select Case az
  Case azgnix: azt = MTBeg
  Case azgdp: azt = tbdp
  Case azgwp: azt = tbwp
  Case azgma: azt = tbma
  Case azgar: azt = tbar
  Case azgpr: azt = tbpr
  Case azgul: azt = tbul
  Case azgab: azt = tbab
 End Select ' case az
End Function ' azt(azgtyp) As TbTyp

#If alt Then
' aufgerufen in: Datenausgeben_Click, MFGRefresh
Function UrlAnsp(ByVal PNr&, ByVal Bervon As Date, ByVal Berbis As Date, ByVal Cn, ByRef UAA!, Optional ByRef UAB!, Optional ByVal mitdruck%) ' Urlaubsanspruch aktuell, Urlaubsanspruch bisher, mitdruck=1: nur aktuellen Urlaub, 2= auch alten urlaub
 Dim rs0 As New ADODB.Recordset
' If IsNull(cn) Then cn = "DRIVER={MySQL ODBC 3.51 Driver};server=mitte;uid=praxis;pwd=...;database=dp;option=11"
 Dim austr As Date, ueberschr%
 On Error GoTo fehler
 austr = Cn.Execute("SELECT REPLACE(COALESCE(`aus`,0),'0000-00-00','1899-12-30') austr FROM `" & tbm(tbma) & "` WHERE `persnr` = " & PNr)!austr ' Austritt
 Dim aktab As Date, aktUA!, Ansprvon As Date, Ansprbis As Date ' altes ab-Datum, alter Urlaubsanspruch
 Dim WAZ!, WAZv! ' Wochenarbeitszeit vorher
 UAB = 0
 UAA = 0
 aktab = 0
 aktUA = 0
 Set rs0 = Nothing
 rs0.Open "SELECT ab, Urlaub, WAZ FROM `" & tbm(tbwp) & "` WHERE persnr = " & PNr & " AND ab < '" & Format(Berbis, "yyyymmdd") & "' ORDER BY ab", Cn, adOpenDynamic, adLockReadOnly
 Do While Not rs0.EOF
  aktab = rs0!ab
  aktUA = rs0!urlaub
  If rs0!WAZ <> 0 Then WAZv = rs0!WAZ
  rs0.Move 1
  If rs0.EOF Then
   Ansprbis = Berbis
  Else
   Ansprbis = rs0!ab
   WAZ = rs0!WAZ
  End If
  If austr <> 0 And austr < Ansprbis Then Ansprbis = austr
  If Ansprbis < Bervon Then Ansprvon = Ansprbis Else Ansprvon = Bervon
  If aktab < Ansprvon And WAZv <> 0 Then
   If mitdruck = 2 Then
    Print #323, "Urlaubsanspruchberechnung bisher:"
    Print #323, "Urlaubsanspruch bisher = Urlaubsanspruch bisher * WAZ / WAZv + (Zeitraumende - Zeitraumbeginn) / 365 * Jahresanspruch in Zeitraum"
    Print #323, CStr(Round(UAB + (Ansprvon - aktab) / 365 * aktUA, 2)) & " = " & CStr(Round(UAB, 2)) & " * " & CStr(WAZ) & " / " & CStr(WAZv) & " + (" & Ansprvon & " - " & aktab & ") / 365 * " & aktUA
   End If ' mitdruck = 2 Then
   UAB = UAB * WAZ / WAZv + (Ansprvon - aktab) / 365 * aktUA
  End If ' aktab < Ansprvon Then
  If aktab > Bervon Then Ansprvon = aktab Else Ansprvon = Bervon
  If Ansprbis >= Bervon And WAZv <> 0 Then
   If mitdruck > 0 Then
    If Not ueberschr Then
     Print #323, "Urlaubsanspruchberechnung aktuell:"
     Print #323, "Urlaubsanspruch aktuell = Urlaubsanspruch aktuell * WAZ / WAZv  + (Zeitraumende - Zeitraumbeginn) / (Jahresende - Jahresbeginn) * Jahresanspruch in Zeitraum"
     ueberschr = True
    End If ' Not ueberschr Then
    Print #323, CStr(Round(UAA + (Ansprbis - Ansprvon) / 365 * aktUA, 2)) & " = " & CStr(Round(UAA, 2)) & " * " & CStr(WAZ) & " / " & CStr(WAZv) & " + (" & Ansprbis & " - " & Ansprvon & ") / (" & Berbis & " - " & Bervon & ") * " & aktUA
   End If ' mitdruck > 0 Then
   UAA = UAA * WAZ / WAZv + (Ansprbis - Ansprvon) / 365 * aktUA
  End If ' Ansprbis >= Bervon Then
 Loop ' While Not rs0.EOF
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in UrlAnsp/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' UrlAnsp
#End If

Public Sub Sendschluessel(Text As Variant, Optional Wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.SendKeys CStr(Text), Wait
   Set WshShell = Nothing
End Sub ' Sendschluessel(Text As Variant, Optional Wait As Boolean = False)

Function datform(DaT) ' for vb-Datumsformat oder vb-double (#)
 On Error GoTo fehler
 If IsNull(DaT) Then
  datform = "null"
 ElseIf (LVobMySQL) Then
  If DaT - Int(DaT) = 0 Then
   datform = "'" + Format(DaT, "yyyy-mm-dd") + "'"
  Else
   datform = "'" + Format(DaT, "yyyy-mm-dd hh:mm:ss") + "'"
  End If
 Else
  datform = "#" + Format(DaT, "mm\/dd\/yy hh:mm:ss") + "#"
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in datForm/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' datForm

Function ProgEnde(Optional frm)
 If Not IsMissing(frm) Then
  On Error Resume Next
  If frm.dbv.wCn Is Nothing Or frm.dbv.wCn.State = 0 Then
   frm.dbv.wCn.Open frm.dbv.CnStr
   DBCnS = frm.dbv.CnStr
  End If
  frm.dbv.wCn.Execute doppelteWeg
 End If
 End
End Function ' ende

Function ausgeb(Ausgabe$)
End Function ' ausgeb(Ausgabe$)


Public Function Key(KeyCode%, Shift%, frm As Form)

End Function
