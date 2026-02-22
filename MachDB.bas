Attribute VB_Name = "Module1"
'Bauanleitung für eine Datenbank wie `//linux1/dp` vom 6.10.24 00:56:59
Option Explicit
Dim cnzCStr$ ' da unter Vista der Connectionstring jetzt nicht mehr aussagekräftig ist
Dim cnz As New ADODB.Connection, FNr&, lErrNr& ' letzter Fehler bei doEx
Dim obProt% ' ob Protokollierung stattfindet, da Protokolldatei zu öffnen
Dim Str(1, 8, 29) As New CString, ArtZ&(3, 8)
Dim hDBn$ ' hiesiger Datenbankname


Sub FüllStr0()
 Str(0, 0, 0) = "arten"
 Str(0, 0, 1) = "`ArtNr`"
 Str(0, 0, 2) = "`erkl`"
 Str(0, 0, 3) = "`Stdn`"
 Str(0, 0, 4) = "`Farbe`"
 Str(0, 0, 5) = "`zusatz`"
 Str(0, 0, 6) = "`ArtNr`"
 ArtZ(0, 0) = 5
 ArtZ(1, 0) = 1
 Str(1, 0, 0) = "CREATE TABLE `arten` ("
 Str(1, 0, 1) = " `ArtNr` varchar(30) NOT NULL COMMENT 'Artnr'"
 Str(1, 0, 2) = " `erkl` varchar(30) DEFAULT NULL COMMENT 'Erklärung'"
 Str(1, 0, 3) = " `Stdn` double(3,1) unsigned NOT NULL COMMENT 'Stunden'"
 Str(1, 0, 4) = " `Farbe` int(4) unsigned NOT NULL"
 Str(1, 0, 5) = " `zusatz` tinyint(1) unsigned DEFAULT NULL"
 Str(1, 0, 6) = "  PRIMARY KEY (`ArtNr`)"
 Str(1, 0, 7) = " ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_german2_ci ROW_FORMAT=DYNAMIC COMMENT='Dienstarten; InnoDB free: 261120 kB'"
End Sub ' FüllStr0

Sub FüllStr1()
 Str(0, 1, 0) = "ausbez"
 Str(0, 1, 1) = "`ID`"
 Str(0, 1, 2) = "`tag`"
 Str(0, 1, 3) = "`PersNr`"
 Str(0, 1, 4) = "`ausbez`"
 Str(0, 1, 5) = "`urlhaus`"
 Str(0, 1, 6) = "`ID`"
 Str(0, 1, 7) = "`Persnr`"
 Str(0, 1, 8) = "`tag`"
 ArtZ(0, 1) = 5
 ArtZ(1, 1) = 3
 Str(1, 1, 0) = "CREATE TABLE `ausbez` ("
 Str(1, 1, 1) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 1, 2) = " `tag` date NOT NULL"
 Str(1, 1, 3) = " `PersNr` int(4) unsigned NOT NULL COMMENT 'Bezug auf Mitarbeiter'"
 Str(1, 1, 4) = " `ausbez` double(5,1) NOT NULL DEFAULT 0.0 COMMENT 'Stunden ausbezahlt'"
 Str(1, 1, 5) = " `urlhaus` double(5,1) NOT NULL DEFAULT 0.0 COMMENT 'Urlaubsstd. ausbezahlt'"
 Str(1, 1, 6) = "  PRIMARY KEY (`ID`)"
 Str(1, 1, 7) = "  KEY `Persnr` (`PersNr`,`tag`) USING BTREE"
 Str(1, 1, 8) = "  KEY `tag` (`tag`)"
 Str(1, 1, 9) = " ENGINE=InnoDB AUTO_INCREMENT=50 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_german2_ci ROW_FORMAT=DYNAMIC COMMENT='Dienstplan'"
End Sub ' FüllStr1

Sub FüllStr2()
 Str(0, 2, 0) = "bilanzen"
 Str(0, 2, 1) = "`PersNr`"
 Str(0, 2, 2) = "`Jahr`"
 Str(0, 2, 3) = "`Urlaub`"
 Str(0, 2, 4) = "`UrlStd`"
 Str(0, 2, 5) = "`Überstunden`"
 Str(0, 2, 6) = "`Fortbildung`"
 Str(0, 2, 7) = "`Planstunden`"
 Str(0, 2, 8) = "`PersNr`"
 Str(0, 2, 9) = "`PersNr2`"
 ArtZ(0, 2) = 7
 ArtZ(1, 2) = 1
 ArtZ(2, 2) = 1
 Str(1, 2, 0) = "CREATE TABLE `bilanzen` ("
 Str(1, 2, 1) = " `PersNr` int(5) unsigned NOT NULL DEFAULT 0"
 Str(1, 2, 2) = " `Jahr` int(4) unsigned NOT NULL"
 Str(1, 2, 3) = " `Urlaub` double(5,1) NOT NULL COMMENT 'Tage'"
 Str(1, 2, 4) = " `UrlStd` double(6,1) NOT NULL COMMENT 'Urlaub in Stunden'"
 Str(1, 2, 5) = " `Überstunden` double(5,1) NOT NULL COMMENT 'Stunden'"
 Str(1, 2, 6) = " `Fortbildung` double(4,1) NOT NULL COMMENT 'Tage'"
 Str(1, 2, 7) = " `Planstunden` double(7,1) NOT NULL COMMENT 'Stunden pro Jahr laut Plan'"
 Str(1, 2, 8) = "  PRIMARY KEY (`PersNr`,`Jahr`) USING BTREE"
 Str(1, 2, 9) = "  CONSTRAINT `PersNr2` FOREIGN KEY (`PersNr`) REFERENCES `mitarbeiter` (`PersNr`)"
 Str(1, 2, 10) = " ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_german2_ci ROW_FORMAT=DYNAMIC COMMENT='Urlaubs- und Überstundenbilanzen'"
End Sub ' FüllStr2

Sub FüllStr3()
 Str(0, 3, 0) = "dienstplan"
 Str(0, 3, 1) = "`ID`"
 Str(0, 3, 2) = "`tag`"
 Str(0, 3, 3) = "`PersNr`"
 Str(0, 3, 4) = "`ArtNr`"
 Str(0, 3, 5) = "`ID`"
 Str(0, 3, 6) = "`Artnr`"
 Str(0, 3, 7) = "`Persnr`"
 Str(0, 3, 8) = "`tag`"
 Str(0, 3, 9) = "`ArtNr`"
 Str(0, 3, 10) = "`Persnr`"
 ArtZ(0, 3) = 4
 ArtZ(1, 3) = 4
 ArtZ(2, 3) = 2
 Str(1, 3, 0) = "CREATE TABLE `dienstplan` ("
 Str(1, 3, 1) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 3, 2) = " `tag` date NOT NULL"
 Str(1, 3, 3) = " `PersNr` int(4) unsigned NOT NULL COMMENT 'Bezug auf Mitareiter'"
 Str(1, 3, 4) = " `ArtNr` varchar(30) DEFAULT NULL COMMENT 'Bezug auf Arten'"
 Str(1, 3, 5) = "  PRIMARY KEY (`ID`)"
 Str(1, 3, 6) = "  KEY `Artnr` (`ArtNr`)"
 Str(1, 3, 7) = "  KEY `Persnr` (`PersNr`,`tag`) USING BTREE"
 Str(1, 3, 8) = "  KEY `tag` (`tag`)"
 Str(1, 3, 9) = "  CONSTRAINT `ArtNr` FOREIGN KEY (`ArtNr`) REFERENCES `arten` (`ArtNr`)"
 Str(1, 3, 10) = "  CONSTRAINT `Persnr` FOREIGN KEY (`PersNr`) REFERENCES `mitarbeiter` (`PersNr`)"
 Str(1, 3, 11) = " ENGINE=InnoDB AUTO_INCREMENT=36965 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_german2_ci ROW_FORMAT=DYNAMIC COMMENT='Dienstplan'"
End Sub ' FüllStr3

Sub FüllStr4()
 Str(0, 4, 0) = "einstellungen"
 Str(0, 4, 1) = "`Einstellung`"
 Str(0, 4, 2) = "`Wert`"
 Str(0, 4, 3) = "`Einstellung`"
 ArtZ(0, 4) = 2
 ArtZ(1, 4) = 1
 Str(1, 4, 0) = "CREATE TABLE `einstellungen` ("
 Str(1, 4, 1) = " `Einstellung` varchar(70) CHARACTER SET latin1 COLLATE latin1_german2_ci NOT NULL"
 Str(1, 4, 2) = " `Wert` varchar(70) CHARACTER SET latin1 COLLATE latin1_german2_ci DEFAULT NULL"
 Str(1, 4, 3) = "  PRIMARY KEY (`Einstellung`)"
 Str(1, 4, 4) = " ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_swedish_ci COMMENT='Einstellungen'"
End Sub ' FüllStr4

Sub FüllStr5()
 Str(0, 5, 0) = "mitarbeiter"
 Str(0, 5, 1) = "`PersNr`"
 Str(0, 5, 2) = "`Kuerzel`"
 Str(0, 5, 3) = "`Nachname`"
 Str(0, 5, 4) = "`Vorname`"
 Str(0, 5, 5) = "`Aus`"
 Str(0, 5, 6) = "`KalDB`"
 Str(0, 5, 7) = "`KalTab`"
 Str(0, 5, 8) = "`KalDatSp`"
 Str(0, 5, 9) = "`KalSp1`"
 Str(0, 5, 10) = "`KalSp2`"
 Str(0, 5, 11) = "`KalSp3`"
 Str(0, 5, 12) = "`KalSp4`"
 Str(0, 5, 13) = "`Beruf`"
 Str(0, 5, 14) = "`PersNr`"
 Str(0, 5, 15) = "`Aus`"
 ArtZ(0, 5) = 13
 ArtZ(1, 5) = 2
 Str(1, 5, 0) = "CREATE TABLE `mitarbeiter` ("
 Str(1, 5, 1) = " `PersNr` int(5) unsigned NOT NULL AUTO_INCREMENT COMMENT 'Personal-Nummer'"
 Str(1, 5, 2) = " `Kuerzel` varchar(10) DEFAULT NULL COMMENT 'Kürzel'"
 Str(1, 5, 3) = " `Nachname` varchar(50) DEFAULT NULL"
 Str(1, 5, 4) = " `Vorname` varchar(50) DEFAULT NULL"
 Str(1, 5, 5) = " `Aus` date DEFAULT NULL COMMENT 'Austritt'"
 Str(1, 5, 6) = " `KalDB` varchar(45) DEFAULT NULL"
 Str(1, 5, 7) = " `KalTab` varchar(45) DEFAULT NULL"
 Str(1, 5, 8) = " `KalDatSp` varchar(45) DEFAULT NULL"
 Str(1, 5, 9) = " `KalSp1` varchar(45) DEFAULT NULL"
 Str(1, 5, 10) = " `KalSp2` varchar(45) DEFAULT NULL"
 Str(1, 5, 11) = " `KalSp3` varchar(45) DEFAULT NULL"
 Str(1, 5, 12) = " `KalSp4` varchar(45) DEFAULT NULL"
 Str(1, 5, 13) = " `Beruf` char(1) DEFAULT NULL COMMENT 'm=MFA,d=Diabetesberaterin,a=Ärztin/Arzt'"
 Str(1, 5, 14) = "  PRIMARY KEY (`PersNr`)"
 Str(1, 5, 15) = "  KEY `Aus` (`Aus`)"
 Str(1, 5, 16) = " ENGINE=InnoDB AUTO_INCREMENT=115 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_german2_ci"
End Sub ' FüllStr5

Sub FüllStr6()
 Str(0, 6, 0) = "protok"
 Str(0, 6, 1) = "`ID`"
 Str(0, 6, 2) = "`tag`"
 Str(0, 6, 3) = "`PersNr`"
 Str(0, 6, 4) = "`ArtNrV`"
 Str(0, 6, 5) = "`ArtNr`"
 Str(0, 6, 6) = "`AendDat`"
 Str(0, 6, 7) = "`AendPC`"
 Str(0, 6, 8) = "`AendUser`"
 Str(0, 6, 9) = "`user`"
 Str(0, 6, 10) = "`ID`"
 Str(0, 6, 11) = "`Persnr`"
 Str(0, 6, 12) = "`Artnr`"
 Str(0, 6, 13) = "`user`"
 Str(0, 6, 14) = "`ProtArtnr`"
 Str(0, 6, 15) = "`ProtPersnr`"
 ArtZ(0, 6) = 9
 ArtZ(1, 6) = 4
 ArtZ(2, 6) = 2
 Str(1, 6, 0) = "CREATE TABLE `protok` ("
 Str(1, 6, 1) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 6, 2) = " `tag` date NOT NULL"
 Str(1, 6, 3) = " `PersNr` int(5) unsigned NOT NULL"
 Str(1, 6, 4) = " `ArtNrV` varchar(30) DEFAULT NULL"
 Str(1, 6, 5) = " `ArtNr` varchar(30) DEFAULT NULL"
 Str(1, 6, 6) = " `AendDat` datetime DEFAULT NULL"
 Str(1, 6, 7) = " `AendPC` varchar(20) DEFAULT NULL"
 Str(1, 6, 8) = " `AendUser` varchar(25) DEFAULT NULL"
 Str(1, 6, 9) = " `user` varchar(45) NOT NULL"
 Str(1, 6, 10) = "  PRIMARY KEY (`ID`)"
 Str(1, 6, 11) = "  KEY `Persnr` (`PersNr`)"
 Str(1, 6, 12) = "  KEY `Artnr` (`ArtNr`)"
 Str(1, 6, 13) = "  KEY `user` (`user`)"
 Str(1, 6, 14) = "  CONSTRAINT `ProtArtnr` FOREIGN KEY (`ArtNr`) REFERENCES `arten` (`ArtNr`)"
 Str(1, 6, 15) = "  CONSTRAINT `ProtPersnr` FOREIGN KEY (`PersNr`) REFERENCES `mitarbeiter` (`PersNr`)"
 Str(1, 6, 16) = " ENGINE=InnoDB AUTO_INCREMENT=46323 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_german2_ci COMMENT='Änderungsprotokoll für Dienstplan'"
End Sub ' FüllStr6

Sub FüllStr7()
 Str(0, 7, 0) = "user"
 Str(0, 7, 1) = "`user`"
 Str(0, 7, 2) = "`Passwort`"
 Str(0, 7, 3) = "`hinzugefügt`"
 Str(0, 7, 4) = "`geändert`"
 Str(0, 7, 5) = "`ID`"
 Str(0, 7, 6) = "`ID`"
 Str(0, 7, 7) = "`user`"
 ArtZ(0, 7) = 5
 ArtZ(1, 7) = 2
 Str(1, 7, 0) = "CREATE TABLE `user` ("
 Str(1, 7, 1) = " `user` varchar(45) NOT NULL"
 Str(1, 7, 2) = " `Passwort` blob NOT NULL"
 Str(1, 7, 3) = " `hinzugefügt` datetime NOT NULL"
 Str(1, 7, 4) = " `geändert` datetime NOT NULL DEFAULT '0000-00-00 00:00:00'"
 Str(1, 7, 5) = " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT"
 Str(1, 7, 6) = "  PRIMARY KEY (`ID`)"
 Str(1, 7, 7) = "  UNIQUE KEY `user` (`user`)"
 Str(1, 7, 8) = " ENGINE=InnoDB AUTO_INCREMENT=36 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_german2_ci COMMENT='Benutzer für Datenänderungen'"
End Sub ' FüllStr7

Sub FüllStr8()
 Str(0, 8, 0) = "wochenplan"
 Str(0, 8, 1) = "`PersNr`"
 Str(0, 8, 2) = "`ab`"
 Str(0, 8, 3) = "`Mo`"
 Str(0, 8, 4) = "`Di`"
 Str(0, 8, 5) = "`Mi`"
 Str(0, 8, 6) = "`Do`"
 Str(0, 8, 7) = "`Fr`"
 Str(0, 8, 8) = "`Sa`"
 Str(0, 8, 9) = "`So`"
 Str(0, 8, 10) = "`WAZ`"
 Str(0, 8, 11) = "`Urlaub`"
 Str(0, 8, 12) = "`PersNr`"
 Str(0, 8, 13) = "`Mo`"
 Str(0, 8, 14) = "`Di`"
 Str(0, 8, 15) = "`Mi`"
 Str(0, 8, 16) = "`Do`"
 Str(0, 8, 17) = "`Fr`"
 Str(0, 8, 18) = "`Sa`"
 Str(0, 8, 19) = "`So`"
 Str(0, 8, 20) = "`DiArt`"
 Str(0, 8, 21) = "`DoArt`"
 Str(0, 8, 22) = "`FrArt`"
 Str(0, 8, 23) = "`MiArt`"
 Str(0, 8, 24) = "`MoArt`"
 Str(0, 8, 25) = "`Persnr1`"
 Str(0, 8, 26) = "`SaArt`"
 Str(0, 8, 27) = "`SoArt`"
 ArtZ(0, 8) = 11
 ArtZ(1, 8) = 8
 ArtZ(2, 8) = 8
 Str(1, 8, 0) = "CREATE TABLE `wochenplan` ("
 Str(1, 8, 1) = " `PersNr` int(5) unsigned NOT NULL DEFAULT 0 COMMENT 'Personal-Nr.'"
 Str(1, 8, 2) = " `ab` date NOT NULL DEFAULT '2007-01-01' COMMENT 'Gültigkeitsbeginn'"
 Str(1, 8, 3) = " `Mo` varchar(10) DEFAULT NULL COMMENT 'Montag'"
 Str(1, 8, 4) = " `Di` varchar(10) DEFAULT NULL COMMENT 'Dienstag'"
 Str(1, 8, 5) = " `Mi` varchar(10) DEFAULT NULL COMMENT 'Mittwoch'"
 Str(1, 8, 6) = " `Do` varchar(10) DEFAULT NULL COMMENT 'Donnerstag'"
 Str(1, 8, 7) = " `Fr` varchar(10) DEFAULT NULL COMMENT 'Freitag'"
 Str(1, 8, 8) = " `Sa` varchar(10) DEFAULT NULL COMMENT 'Samstag'"
 Str(1, 8, 9) = " `So` varchar(10) DEFAULT NULL COMMENT 'Sonntag'"
 Str(1, 8, 10) = " `WAZ` double(3,1) NOT NULL DEFAULT 38.5"
 Str(1, 8, 11) = " `Urlaub` int(2) NOT NULL DEFAULT 26 COMMENT 'Urlaubstage pro Jahr'"
 Str(1, 8, 12) = "  PRIMARY KEY (`PersNr`,`ab`)"
 Str(1, 8, 13) = "  KEY `Mo` (`Mo`)"
 Str(1, 8, 14) = "  KEY `Di` (`Di`)"
 Str(1, 8, 15) = "  KEY `Mi` (`Mi`)"
 Str(1, 8, 16) = "  KEY `Do` (`Do`)"
 Str(1, 8, 17) = "  KEY `Fr` (`Fr`)"
 Str(1, 8, 18) = "  KEY `Sa` (`Sa`)"
 Str(1, 8, 19) = "  KEY `So` (`So`)"
 Str(1, 8, 20) = "  CONSTRAINT `DiArt` FOREIGN KEY (`Di`) REFERENCES `arten` (`ArtNr`)"
 Str(1, 8, 21) = "  CONSTRAINT `DoArt` FOREIGN KEY (`Do`) REFERENCES `arten` (`ArtNr`)"
 Str(1, 8, 22) = "  CONSTRAINT `FrArt` FOREIGN KEY (`Fr`) REFERENCES `arten` (`ArtNr`)"
 Str(1, 8, 23) = "  CONSTRAINT `MiArt` FOREIGN KEY (`Mi`) REFERENCES `arten` (`ArtNr`)"
 Str(1, 8, 24) = "  CONSTRAINT `MoArt` FOREIGN KEY (`Mo`) REFERENCES `arten` (`ArtNr`)"
 Str(1, 8, 25) = "  CONSTRAINT `Persnr1` FOREIGN KEY (`PersNr`) REFERENCES `mitarbeiter` (`PersNr`)"
 Str(1, 8, 26) = "  CONSTRAINT `SaArt` FOREIGN KEY (`Sa`) REFERENCES `arten` (`ArtNr`)"
 Str(1, 8, 27) = "  CONSTRAINT `SoArt` FOREIGN KEY (`So`) REFERENCES `arten` (`ArtNr`)"
 Str(1, 8, 28) = " ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_german2_ci COMMENT='Wochenplan'"
End Sub ' FüllStr8

Function doEx&(sql$, obtolerant%) ' SQL-Befehl ausführen, Fehler anzeigen
 Dim rAF&, FMeld$
 Dim lErrNr&, fDesc$
 On Error Resume Next
 cnz.DefaultDatabase = hDBn
 If obtolerant Then On Error Resume Next Else On Error GoTo fehler
 myEFrag sql, rAF, cnz, True, lErrNr, fDesc
' lErrNr = Err.Number
 FMeld = IIf(lErrNr = 0, "Kein Fehler", "Err.Nr " & lErrNr & ", " & fDesc) & ", rAf: " & rAF & " bei " & sql
 On Error GoTo fehler
 Debug.Print "doEx, FMeld: " & FMeld
 Debug.Print fDesc
 If obProt Then Print #302, FMeld
 DoEvents
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.name
#Else
 AnwPfad = App.Path
#End If
Select Case Err.Number
 Case -2147467259
  If InStrB(Err.Description, "nicht erzeugen") Then ' 'Kann Tabelle 'testDB1.faxe' nicht erzeugen (Fehler: 150)
   doEx = 150
   Exit Function
  ElseIf InStrB(Err.Description, "is not BASE TABLE") <> 0 Then
   doEx = 151
   Exit Function
  ElseIf InStrB(Err.Description, "MySQL server has gone away") Then
   cnz.Close
   cnz.Open
   Call doEx("USE `" & hDBn & "`", 0)
   Resume
  End If
End Select
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) & vbCrLf & "LastDLLError: " & CStr(Err.LastDllError) & vbCrLf & "Source: " & IIf(IsNull(Err.source), "", CStr(Err.source)) & vbCrLf & "Description: " & Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doEx/" & AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function ' doEx

Function SplitN&(ByRef q$, Sep$, erg$()) ' da Split() Speicher fraß
 Dim p1&, p2&, Slen&, obExit%, runde&
 On Error GoTo fehler
 If Not IsNull(q) Then
  Slen = Len(Sep)
  For runde = 1 To 2
   p2 = 0
   Do
    p1 = p2
    p2 = InStr(p1 + Slen, q, Sep)
    If p2 = 0 Then p2 = Len(q) + 1: obExit = True
    If p2 <> 0 Then
     If runde = 2 Then
      erg(SplitN) = Mid$(q, p1 + Slen, p2 - p1 - Slen)
     End If
     SplitN = SplitN + 1
    End If
    If obExit Then Exit Do
   Loop
   If runde = 1 Then
    ReDim erg(SplitN - 1)
    SplitN = 0
    obExit = 0
   End If
  Next runde
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in SplitN/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' aufSplit

' in calldoGenMachDB_Click
Public Function doMach_dp(DBn$, DBCn As ADODB.Connection, Optional Server$, Optional obStumm% = True) ' Datenbankname
 Dim rsc As New ADODB.Recordset, sct$, Spli$(), tStr$, TMt As New CString, TabEig$
 Dim i&, p1&, p2&, p3&, CLen&, CLen1&, obLT%, ep$, pwp&, pwd$, mpwd$
 Dim Index$()
 On Error Resume Next
 hDBn = DBn
 Open App.Path & "\MachDB.bas_prot.txt" For Output As #302
 obProt = (Err.Number = 0)
 On Error GoTo fehler
 ep = DBCn.Properties("Extended Properties")
 pwp = InStr(UCase$(ep), "PWD=") + 4
 If pwp <> 0 Then pwd = Mid$(ep, pwp, InStr(pwp, ep, ";") - pwp)
 mpwd = MachDatenbank.setzmpwd()
 If mpwd = "" Then Exit Function
 If LenB(Server) = 0 Then Server = GetServr(DBCn)
 cnzCStr = "PROVIDER=MSDASQL;driver={MySQL ODBC 8.0 Unicode Driver};server=" & Server & ";uid=mysql;pwd=" & mpwd & ";"
 Set cnz = Nothing
 cnz.Open cnzCStr & mpwd & ";"
 Call doEx("CREATE DATABASE IF NOT EXISTS `" & DBn & "` CHARACTER SET utf8mb4 COLLATE utf8mb4_german2_ci;", 0)
 Call doEx("GRANT ALL ON `" & DBn & "`.* TO 'praxis'@'%' IDENTIFIED BY '" & pwd & "' WITH GRANT OPTION", 0)
 Call doEx("GRANT ALL ON `" & DBn & "`.* TO 'praxis'@'localhost' IDENTIFIED BY '" & pwd & "' WITH GRANT OPTION", 0)
 Call doEx("USE `" & DBn & "`", 0)
 Call doEx("SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", 0)
 FüllStr0
 FüllStr1
 FüllStr2
 FüllStr3
 FüllStr4
 FüllStr5
 FüllStr6
 FüllStr7
 FüllStr8
 Call doEx("SET FOREIGN_KEY_CHECKS = 0", 0)

 Dim j&, ZZ&, Tbl$, sql As New CString
 For i = 0 To 8
  If InStrB(Str(1, i, 0), "CREATE TABLE") <> 0 Then
   Tbl = Str(0, i, 0)
   ZZ = ArtZ(0, i) + ArtZ(1, i)
   sql = "CREATE TABLE IF NOT EXISTS `" & Tbl & "` (" & vbLf
   For j = 1 To ZZ
    If Str(1, i, j) <> "" Then
     sql.Append Str(1, i, j)
     If j < ZZ Then sql.AppVar Array(",", vbLf)
    End If ' If Str(1, i, j) <> "" Then
   Next j
   ZZ = ZZ + ArtZ(2, i) + 1
   sql.AppVar Array(vbLf, ")")
   sql.Append Str(1, i, ZZ)
   FNr = doEx(sql.Value, 0)
   Do
    Set rsc = Nothing
    myFrag rsc, "SHOW CREATE TABLE `" & Tbl & "`", adOpenStatic, cnz, adLockReadOnly
    sct = rsc.Fields(1)
    If InStrB(sct, "CREATE ALGORITHM") = 1 Then
     FNr = doEx("DROP VIEW `" & Tbl & "`", 0)
     FNr = doEx(sql.Value, 0)
    Else
     Exit Do
    End If
   Loop
   If InStrB(AIoZ(sct), AIoZ(Str(1, i, ZZ))) = 0 Then
    Call doEx("ALTER TABLE `" & Tbl & "`" & Str(1, i, ZZ), 0)
   End If
   TMt.Clear
   SplitN sct, vbLf, Spli
   For j = 1 To ArtZ(0, i) ' Tabellenfelder
    Dim k&, enthalten%, genau%, Posi$
    enthalten = 0
    genau = 0
    k = 0
    ' SET rsc = Nothing
    myFrag rsc, "SHOW columns FROM `" & Tbl & "` WHERE field = '" & Mid$(Str(0, i, j), 2, Len(Str(0, i, j)) - 2) & "'", adOpenStatic, cnz, adLockReadOnly
    enthalten = Not rsc.BOF
    If enthalten Then
     genau = (InStrB(sct, Str(1, i, j)) <> 0)
     If Not genau Then
      CLen = -1 ' Column-Length nicht kürzen
      obLT = (InStrB(sct, Str(0, i, j) & " longtext") <> 0)
      If Not obLT Then
       p1 = InStr(sct, "(")
       p2 = InStr(p1, sct, Str(0, i, j)) 'zCat.Tables(Tbl).Columns(k).Name & "`")
       If p2 = 0 Then p2 = InStr(p1, LCase$(sct), LCase$(Str(0, i, j)))
       p1 = InStr(p2, sct, "(")
       p3 = InStr(p2, sct, ",")
       If p3 = 0 Then p3 = InStr(p2, sct, vbLf & ")")
       If p1 <> 0 And p1 < p3 Then
        p2 = InStr(p1, sct, ")")
        CLen = Mid$(sct, p1 + 1, p2 - p1 - 1)
       End If
      End If
     End If
    End If
    If Not enthalten Or Not genau Then
     If j = 1 Then
      Posi = " FIRST,"
     Else
      Posi = " AFTER " & Str(0, i, j - 1) & ","
     End If
     If Not enthalten Then
      TMt.AppVar (Array(" add ", Str(1, i, j), Posi))
     ElseIf Not genau Then
      If CLen <> -1 Or obLT Then
       p1 = InStr(Str(1, i, j), "(")
       If p1 <> 0 Then
        p2 = InStr(p1, Str(1, i, j), ")")
        If p2 <> 0 Then
         CLen1 = Mid$(Str(1, i, j), p1 + 1, p2 - p1 - 1)
         If obLT Then
          Str(1, i, j).REPLACE "varchar(" & CLen1 & ")", "longtext"
         ElseIf CLen1 < CLen Then
          Str(1, i, j).REPLACE "(" & CLen1 & ")", "(" & CLen & ")"
         End If ' obLT THEN ELSE
         genau = (InStrB(sct, Str(1, i, j)) <> 0)
        End If ' p2 <> 0 THEN"
       End If ' p1 <> 0 THEN"
      End If ' CLen <> -1 OR obLT THEN"
      If Not genau Then
       TMt.AppVar (Array(" MODIFY ", Str(1, i, j), Posi))
      End If ' Not genau THEN
     End If ' Not enthalten THEN"
    End If ' Not enthalten OR Not genau THEN"
   Next j
   For j = ArtZ(0, i) + 1 To ArtZ(0, i) + ArtZ(1, i) ' Indices
    If InStrB(sct, Str(1, i, j)) = 0 Then
     If InStrB(Str(1, i, j).Value, "PRIMARY") <> 0 Then
      If InStrB(sct, "PRIMARY KEY (") <> 0 Then
       TMt.Append (" DROP PRIMARY KEY,")
      End If
     Else
      If InStrB(sct, "KEY " & Str(0, i, j).Value) <> 0 Then
       TMt.AppVar Array(" DROP KEY ", Str(0, i, j), ",")
      End If
     End If
     TMt.AppVar Array(" add ", Str(1, i, j), ",")
    End If
   Next j
   If TMt.length <> 0 Then
    TMt.Cut (TMt.length - 1)
    Call doEx("ALTER TABLE `" & Tbl & "` " & TMt.Value, -1)
   End If
  End If ' InStrB(Str(1, i, 0), "CREATE TABLE") <> 0 THEN
 Next i
 For i = 0 To 8
  If InStrB(Str(1, i, 0), "CREATE TABLE") <> 0 Then
   Tbl = Str(0, i, 0)
   ZZ = ArtZ(0, i) + ArtZ(1, i)
   ' SET rsc = nothing
   myFrag rsc, "SHOW CREATE TABLE `" & Tbl & "`", adOpenStatic, cnz, adLockReadOnly
   sct = rsc.Fields(1)
   ZZ = ZZ + ArtZ(2, i) + 1
   For j = ArtZ(0, i) + ArtZ(1, i) + 1 To ZZ - 1 'Constraints
    If InStrB(sct, Str(1, i, j)) = 0 Then
     If InStrB(sct, "CONSTRAINT " & Str(0, i, j)) <> 0 Then
      Call doEx("ALTER TABLE `" & Tbl & "` DROP FOREIGN KEY " & Str(0, i, j), 0)
     End If
     Call doEx("ALTER TABLE `" & Tbl & "` ADD" & Str(1, i, j), 0)
    End If
   Next j
  End If ' InStrB(Str(1, i, 0), "CREATE TABLE") <> 0 THEN
 Next i
 Dim runde%
 For runde = 0 To 4
  For i = 0 To 8
   If InStrB(Str(1, i, 0), "DEFINER VIEW") <> 0 Then
    Dim obCr%
    obCr = 0
    ' SET rsc = Nothing
    myFrag rsc, "SHOW TABLES FROM `" & DBn & "` WHERE `tables_in_" & DBn & "` = """ & Str(0, i, 0) & """", adOpenStatic, cnz, adLockReadOnly
    If rsc.BOF Then
     obCr = True
    Else
     Set rsc = Nothing
     myFrag rsc, "SHOW CREATE TABLE `" & Str(0, i, 0) & "`", adOpenStatic, cnz, adLockReadOnly
     If rsc.Fields(1) <> Str(1, i, 0) Then
      Call doEx("DROP TABLE IF EXISTS `" & Str(0, i, 0) & "`", 0)
      Call doEx("DROP VIEW IF EXISTS `" & Str(0, i, 0) & "`", 0)
      obCr = True
     End If
    End If
    If obCr Then
     Call doEx(Str(1, i, 0).Value, True)
    End If
   End If
   Dim fsql$
   fsql = _
   "CREATE DEFINER=`praxis`@`%` FUNCTION `obft`(dt DATE) RETURNS float" & _
   "    NO SQL" & _
   "    DETERMINISTIC" & _
   "BEGIN  DECLARE a,b,c,d,e,tag,monat,jahr INT;" & _
   " DECLARE erg DATE;" & _
   " SET jahr=YEAR(dt);" & _
   " SET a=jahr MOD 19; SET b=jahr MOD 4; SET c=jahr MOD 7;" & _
   " SET d=(19*a+24)MOD 30; SET e=(2*b+4*c+6*d+5)MOD 7; SET Tag=22+d+e; SET monat=3;" & _
   " IF Tag>31 THEN SET Tag=d+e-9; SET monat=4; ELSEIF Tag=26 && monat=4 THEN SET Tag=19;" & _
   " ELSEIF Tag=25&&monat=4&&d=28&&e=6&&a>10 THEN SET Tag=18; END IF;" & _
   " SET erg = STR_TO_DATE(CONCAT(jahr,'-',monat,'-',Tag),'%Y-%m-%d');" & _
   " SET tag=DAY(dt);" & _
   " SET monat=MONTH(dt);" & _
   " IF (tag=1&&monat IN(1,5,11))||(tag=6&&monat=1)||(tag=15&&monat=8)||(tag=3&&monat=10)||(tag IN(24,25,26)&&monat=12)||DATEDIFF(dt,erg)IN(-2,0,1,39,49,50,60)THEN RETURN 1; END IF;" & _
   " IF DATEDIFF(dt,erg)=-47 THEN RETURN 0.5; END IF;" & _
   " RETURN 0;"
   Call doEx(fsql, True)
  Next i
 Next runde
 Call doEx("SET FOREIGN_KEY_CHECKS = 1", 0)
 If obProt Then Close #302
 If Not obStumm Then
  MsgBox "Fertig mit doMach_dp(" & DBn & ", DBCn," & Server & ")!"
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = currentDB.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) & vbCrLf & "LastDLLError: " & CStr(Err.LastDllError) & vbCrLf & "Source: " & IIf(IsNull(Err.source), "", CStr(Err.source)) & vbCrLf & "Description: " & Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doMach_dp/" & AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function 'doMach_dp

Function GetServr$(DBCn As ADODB.Connection)
 Dim spos&, sp2&, DBCs$
 DBCs = DBCn.Properties("Extended Properties")
 spos = InStr(1, DBCs, "server=", vbTextCompare)
 If spos <> 0 Then
  sp2 = InStr(spos, DBCs, ";")
  If sp2 = 0 Then sp2 = Len(DBCs)
  GetServr = Mid$(DBCs, spos + 7, sp2 - spos - 7)
 End If ' spos <> 0 THEN
End Function ' GetServr

Function AIoZ(Ursp) As CString ' Ursp kann $ oder CString sein
 Const Such$ = "AUTO_INCREMENT="
 Set AIoZ = New CString
 AIoZ = Ursp
 Dim p0&, p1&
 p0 = AIoZ.Instr(Such)
 If p0 <> 0 Then
  p1 = AIoZ.Instr(" ", p0)
  AIoZ.Cut (p0 - 2)
  AIoZ.Append Mid$(Ursp, p1)
 End If
End Function ' AIoZ(Ursp$) AS CString
