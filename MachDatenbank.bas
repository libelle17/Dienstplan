Attribute VB_Name = "MachDatenbank"
Public mpwd$

' in aktualisiercon
Public Function setzmpwd$(Optional neu%)
 If neu Or mpwd = "" Then
  mpwd = InputBox("Datenbankpasswort für Benutzer `mysql`:", "Passworteingabe", mpwd)
 End If ' neu Or mpwd = "" Then
 setzmpwd = mpwd
End Function ' setzmpwd(Optional neu%)

