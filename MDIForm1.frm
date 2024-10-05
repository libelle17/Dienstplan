VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.MDIForm MDI 
   BackColor       =   &H8000000C&
   Caption         =   "Dienstplan"
   ClientHeight    =   9030
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13725
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":030A
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox ucMDIKeys 
      Align           =   1  'Oben ausrichten
      Height          =   9015
      Left            =   0
      ScaleHeight     =   8955
      ScaleWidth      =   13665
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13725
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Left            =   6840
         TabIndex        =   11
         Top             =   1440
         Width           =   2055
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSH2 
         Height          =   1215
         Left            =   9120
         TabIndex        =   10
         Top             =   1320
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2143
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton Rechts 
         Caption         =   "&>"
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Links 
         Caption         =   "&<"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   5280
         TabIndex        =   1
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Tb1 
         Height          =   285
         Left            =   3000
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1815
      End
      Begin VB.ComboBox Cb1 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1815
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MFG 
         Height          =   1095
         Left            =   0
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1931
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label cnLab 
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   0
         Width           =   11415
      End
      Begin VB.Label Jahr 
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   0
         Width           =   495
      End
      Begin VB.Label aCtl 
         Height          =   255
         Left            =   13080
         TabIndex        =   6
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Überschrift 
         Height          =   255
         Left            =   -120
         TabIndex        =   4
         Top             =   0
         Width           =   11415
      End
   End
   Begin VB.Menu Datei 
      Caption         =   "&Datei"
      Begin VB.Menu Pfeilu 
         Caption         =   "&Pfeilu"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Seitu 
         Caption         =   "&Seite nach unten"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Erneuern 
         Caption         =   "&Erneuern"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Farbauswahl 
         Caption         =   "&Farbauswahl"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu ProtokollAnzeigen 
         Caption         =   "&Protokoll anzeigen"
      End
      Begin VB.Menu Neuberechnen 
         Caption         =   "Neu &Berechnen"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu NeuZeichnen 
         Caption         =   "Neu &Zeichnen"
         Shortcut        =   {F5}
      End
      Begin VB.Menu zuHeuteGehen 
         Caption         =   "zu &heute gehen"
         Shortcut        =   ^H
      End
      Begin VB.Menu Beenden 
         Caption         =   "&Beenden"
      End
   End
   Begin VB.Menu ZeigDienstplan 
      Caption         =   "D&ienstplan"
   End
   Begin VB.Menu ZeigMitarbeiter 
      Caption         =   "&Mitarbeiter"
   End
   Begin VB.Menu ZeigArten 
      Caption         =   "&Arten"
   End
   Begin VB.Menu Optionen 
      Caption         =   "&Optionen"
      Begin VB.Menu Vordergrundfarbe1 
         Caption         =   "&Vordergrundfarbe 1"
      End
      Begin VB.Menu Vordergrundfarbe2 
         Caption         =   "V&ordergrundfarbe 2"
      End
      Begin VB.Menu DatenAusgeben 
         Caption         =   "Daten für Mitarbeiter &ausgeben"
         Shortcut        =   ^M
      End
      Begin VB.Menu Urlaubausgeben 
         Caption         =   "U&rlaub für Mitarbeiter anzeigen"
         Shortcut        =   ^U
      End
      Begin VB.Menu Neuberechnen_für_Mitarbeiter 
         Caption         =   "&Neuberechnen für Mitarbeiter"
         Shortcut        =   ^N
      End
      Begin VB.Menu KürzelAnzeigenWechseln 
         Caption         =   "&Kürzel anzeigen wechseln"
      End
      Begin VB.Menu DatenSpeichern 
         Caption         =   "Daten &speichern"
      End
      Begin VB.Menu Datenbank 
         Caption         =   "&Datenbank wählen"
      End
      Begin VB.Menu DatenbankErstellen 
         Caption         =   "Datenbank erstellen"
      End
      Begin VB.Menu userliste 
         Caption         =   "&Userliste"
      End
      Begin VB.Menu userlöschen 
         Caption         =   "user &löschen"
      End
      Begin VB.Menu userhinzufügen 
         Caption         =   "user &hinzufügen"
      End
   End
   Begin VB.Menu loggen 
      Caption         =   "einlo&ggen"
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Bedingungen: Für F4-Hilfe muss das Erklärungsfeld in der Kindtabelle gleich nach dem Auswahlfeld kommen
'              Wenn Untertabelle beabeitet wird ("einschr("), dann wird die Datensatzauswahl dort anhand eines singulären Referenzfeldes identifiziert, das in beiden Tabellen gleich heißen muß
Option Explicit
Const tzsbd As Date = #1/1/2020# ' tagesarbeitszeitspezifische Urlaubsberechnung ab diesem Datum
Public Wapp  As Object ' As Word.Application '
Dim WordWasNotRunning As Boolean ' Flag For final word unload
Const wdWindowStateMaximize% = 1
Const BezCol& = 0, NNCol& = 2, VNCol& = 3 ' Spaltenzwecke in der Mitarbeitersicht
Dim ZeiZa%
Dim DoNotChange%
Dim sql$
Public Tabl$
Public MfGTyp As azgtyp, altMfgTyp As azgtyp
Dim verwerfen% ' für Esc in Tb1 oder Cb1
'Dim nichtweiter% ' z.B. Befehlt "Mitarbeiter" soll nicht zu cellweiter führen
Dim iCol%
Dim obtb%
Dim merkCol&(Az0 To azgende - 1), merkRow&(Az0 To azgende - 1), merkTop&(Az0 To azgende - 1), merkLeft&(Az0 To azgende - 1)
Dim vormerkCol&(Az0 To azgende - 1), vormerkRow&(Az0 To azgende - 1), vormerkTop&(Az0 To azgende - 1), vormerkLeft&(Az0 To azgende - 1)
Const MaxPrim% = 2 ' Primärindex, bis zu 3 Felder
Dim PrimI$(MTBeg To tbende - 1, MaxPrim)
Dim obAuto%(MTBeg To tbende - 1) ' auto_increment
Dim obSp1Fest%(MTBeg To tbende - 1) ' Tabellen-Spalte 1 fest
Dim SpZ%(MTBeg To tbende - 1) ' Tabellen-Spaltenzahl
Dim SpvDBSp%() ' Spaltennummer eines Tabellen-Feldes
Dim DBSpvSp%() ' Feldnummer einer Tabellen-Spalte
Dim SpNm$() ' Feldname einer Tabellen- Spalte
Dim SpCm$() ' Feldbeschriftung einer Tabellen-Spalte
Dim FdTyp() As DataTypeEnum ' Feldtyp einer Tabellen-Spalte
Dim RefTab$()
Dim RefSp$()
Dim Einschr$(Az0 To azgende - 1)
Dim EinsNm$(Az0 To azgende - 1)
Dim EinsFd$(Az0 To azgende - 1)
Dim EinsWt(Az0 To azgende - 1)
Dim altFarbe&(Az0 To azgende - 1)
Dim rAf&
Dim NoLostFocus%
Dim BegD As Date, EndD As Date, pn&(), kue$() ' Beginn, Ende, Personalnummern für die Reihenfolge in der Dienstplanseite
Dim cRow&(Az0 To azgende - 1), cCol&(Az0 To azgende - 1) ' für Tb1 und Cb1 und spätere Änderungen
Dim noenter%
Dim MFG_click_abbrech%
Dim nouc% ' <> 0 => Die Steuerung wird aktuell nicht vom Benutzersteuerelement übernommen
Dim fgespei%(Az0 To azgende - 1)  ' Farbe gespeichert
Dim rsf As ADODB.Recordset ' Farben
Dim AusFeldNr%
Const SZZ% = 6 ' Sonderzeilenzahl
Dim VGF1&, VGF2& ' Vordergrundfarbe1, Vordergrundfarbe2
Const mitfett% = True
Const mitVGF% = False
Private m_iSortCol As Integer
Private m_iSortType As Integer
Public WithEvents dbv As DBVerb
Attribute dbv.VB_VarHelpID = -1
Public frmL As New frmLogin
Public User$, Usergeprüft As Date
Public Clip$
Dim wiederholt%
Dim obMSH2%
Const obMySQL% = True

' benötigte API-Deklarationen
Private Declare Sub keybd_event Lib "user32" ( _
  ByVal bVk As Byte, _
  ByVal bScan As Byte, _
  ByVal dwFlags As Long, _
  ByVal dwExtraInfo As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds&)
 
Private Declare Function VkKeyScan% Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte)
Private Declare Function MapVirtualKey& Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode&, ByVal wMapType&)
 
' Hier die benötigten API-Deklarationen
Private Declare Function ShellExecute Lib "shell32.dll" _
  Alias "ShellExecuteA" ( _
  ByVal hwnd As Long, _
  ByVal lpOperation As String, _
  ByVal lpFile As String, _
  ByVal lpParameters As String, _
  ByVal lpDirectory As String, _
  ByVal nShowCmd As Long) As Long
 
Private Const KEYEVENTF_KEYUP = &H2
Private Const KEYEVENTF_EXTENDEDKEY = &H1
 
' Virtual KeyCodes
Private Enum eVirtualKeyCode
  VK_BAK = &H8
  VK_TAB = &H9
  VK_CLEAR = &HC
  VK_RETURN = &HD
  VK_SHIFT = &H10
  VK_CONTROL = &H11
  VK_MENU = &H12
  VK_PAUSE = &H13
  VK_CAPITAL = &H14
  VK_ESCAPE = &H1B
  VK_PRIOR = &H21
  VK_NEXT = &H22
  VK_END = &H23
  VK_HOME = &H24
  VK_LEFT = &H25
  VK_UP = &H26
  VK_RIGHT = &H27
  VK_DOWN = &H28
  VK_SELECT = &H29
  VK_SNAPSHOT = &H2C  ' NEU! Windows-Taste
  VK_INSERT = &H2D
  VK_DELETE = &H2E
  VK_HELP = &H2F
  VK_F1 = &H70
  VK_F2 = &H71
  VK_F3 = &H72
  VK_F4 = &H73
  VK_F5 = &H74
  VK_F6 = &H75
  VK_F7 = &H76
  VK_F8 = &H77
  VK_F9 = &H78
  VK_F10 = &H79
  VK_F11 = &H7A
  VK_F12 = &H7B
  VK_F13 = &H7C
  VK_F14 = &H7D
  VK_F15 = &H7E
  VK_F16 = &H7F
  VK_NUMLOCK = &H90
  VK_SCROLL = &H91
  VK_WIN = &H5B     ' NEU! Windows-Taste
  VK_APPS = &H5D    ' NEU! Taste für Kontextmenü
End Enum

Enum Richtung
 Rec = 1
 Lin
 obe
 unt
 gre ' ganz rechts
 gli ' ganz links
 gob ' ganz oben
 gun ' ganz unten
 stno ' Seite nach oben
 stnu ' Seite nach unten
End Enum ' Richtung

' Die nachfolgende Prozedur aktiviert den im
' System registrierten Standard-Browser und lädt
' die durch URL angegebene Internetadresse
Private Sub URLGoTo(ByVal hwnd As Long, ByVal URL As String)
  ' hWnd: Das Fensterhandle des
  ' aufrufenden Formulars
  Screen.MousePointer = 11
  Call ShellExecute(hwnd, "Open", URL, "", "", 1)
  Screen.MousePointer = 0
End Sub


' Text durch Simulieren von Tastenanschlägen
' an das aktive Control senden
Public Sub SendKeysEx(ByVal sText As String)
  Dim vk As eVirtualKeyCode
  Dim sChar As String
  Dim i As Integer
  Dim bShift As Boolean
  Dim bAlt As Boolean
  Dim bCtrl As Boolean
  Dim nScan As Long
  Dim nExtended As Long
 
  ' Jedes Zeichen einzeln senden
  For i = 1 To Len(sText)
    ' aktuelles Zeichen extrahieren
    sChar = Mid$(sText, i, 1)
 
    ' Sonderzeichen?
    bShift = False: bAlt = False: bCtrl = False
    If sChar = "{" Then
      If UCase$(Mid$(sText, i + 1, 9)) = "BACKSPACE" Then
        vk = VK_BAK
        i = i + 9
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "BS" Then
        vk = VK_BAK
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "BKSP" Then
        vk = VK_BAK
        i = i + 5
      ElseIf UCase$(Mid$(sText, i + 1, 5)) = "BREAK" Then
        vk = VK_PAUSE
        i = i + 6
      ElseIf UCase$(Mid$(sText, i + 1, 8)) = "CAPSLOCK" Then
        vk = VK_CAPITAL
        i = i + 9
      ElseIf UCase$(Mid$(sText, i + 1, 6)) = "DELETE" Then
        vk = VK_DELETE
        i = i + 7
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "DEL" Then
        vk = VK_DELETE
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "DOWN" Then
        vk = VK_DOWN
        i = i + 5
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "UP" Then
        vk = VK_UP
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "LEFT" Then
        vk = VK_LEFT
        i = i + 5
      ElseIf UCase$(Mid$(sText, i + 1, 5)) = "RIGHT" Then
        vk = VK_RIGHT
        i = i + 6
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "END" Then
        vk = VK_END
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 5)) = "ENTER" Then
        vk = VK_RETURN
        i = i + 6
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "HOME" Then
        vk = VK_HOME
        i = i + 5
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "ESC" Then
        vk = VK_ESCAPE
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "HELP" Then
        vk = VK_HELP
        i = i + 5
      ElseIf UCase$(Mid$(sText, i + 1, 6)) = "INSERT" Then
        vk = VK_INSERT
        i = i + 7
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "INS" Then
        vk = VK_INSERT
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 7)) = "NUMLOCK" Then
        vk = VK_NUMLOCK
        i = i + 8
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "PGUP" Then
        vk = VK_PRIOR
        i = i + 5
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "PGDN" Then
        vk = VK_NEXT
        i = i + 5
      ElseIf UCase$(Mid$(sText, i + 1, 10)) = "SCROLLLOCK" Then
        vk = VK_SCROLL
        i = i + 11
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "TAB" Then
        vk = VK_TAB
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F1" Then
        vk = VK_F1
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F2" Then
        vk = VK_F2
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F3" Then
        vk = VK_F3
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F4" Then
        vk = VK_F4
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F5" Then
        vk = VK_F5
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F6" Then
        vk = VK_F6
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F7" Then
        vk = VK_F7
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F8" Then
        vk = VK_F8
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F9" Then
        vk = VK_F9
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "F10" Then
        vk = VK_F10
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "F11" Then
        vk = VK_F11
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "F12" Then
        vk = VK_F12
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "F13" Then
        vk = VK_F13
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "F14" Then
        vk = VK_F14
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "F15" Then
        vk = VK_F15
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "F16" Then
        vk = VK_F16
        i = i + 4
 
      ' NEU! Windows-Taste
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "WIN" Then
        vk = VK_WIN
        i = i + 4
 
      ' NEU! Kontextmenü
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "APPS" Then
        vk = VK_APPS
        i = i + 5
 
      ' NEU! PrintScreen-Taste (DRUCK)
      ElseIf UCase$(Mid$(sText, i + 1, 5)) = "PRINT" Then
        vk = VK_SNAPSHOT
        i = i + 6
      End If
 
    ElseIf sChar = "+" Then
      ' Umschalttaste
      vk = VK_SHIFT
 
    ElseIf sChar = "%" Then
      ' ALT
      vk = VK_MENU
 
    ElseIf sChar = "^" Then
      ' STRG
      vk = VK_CONTROL
 
    Else
      ' Virtual KeyCode ermitteln...
      vk = VkKeyScan(Asc(sChar))
    End If
 
    nScan = MapVirtualKey(vk, 2)
    nExtended = 0
    If nScan = 0 Then nExtended = KEYEVENTF_EXTENDEDKEY
    nScan = MapVirtualKey(vk, 0)
 
    If vk <> VK_SHIFT Then
      ' Großbuchstabe...?
      bShift = (vk And &H100)
      bCtrl = (vk And &H200)
      bAlt = (vk And &H400)
      vk = (vk And &HFF)
    End If
 
    ' niederdrücken und wieder loslassen
    If bShift Then keybd_event VK_SHIFT, 0, 0, 0
    If bCtrl Then keybd_event VK_CONTROL, 0, 0, 0
    If bAlt Then keybd_event VK_MENU, 0, 0, 0
 
    keybd_event vk, nScan, nExtended, 0
    keybd_event vk, nScan, KEYEVENTF_KEYUP Or nExtended, 0
 
    ' Shift (Umsch)-Taste wieder loslassen
    If bShift Then keybd_event VK_SHIFT, 0, KEYEVENTF_KEYUP, 0
    If bCtrl Then keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
    If bAlt Then keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
  Next i
End Sub ' SendKeysEx(ByVal sText As String)


Private Sub Befehl(Bef$, Optional Bef2$)
 Dim Ctrl As Control, altCol&
 On Error GoTo fehler
 If obdebug Then Debug.Print "Befehl(", Bef
 With Me.MFG
  altCol = .Col
  .Col = .Cols - 1
  Set Ctrl = Me.Command1
  Ctrl.Visible = True
  On Error Resume Next
  Ctrl.Height = .CellHeight  'minus 10 so that grid lines
  On Error GoTo fehler
'  Ctrl.Width = .CellWidth  '  will not be overwritten
  Ctrl.Left = .CellLeft + .CellWidth + .Left
  Ctrl.Top = .CellTop + .Top - 10
  DoNotChange = True
  Ctrl.Caption = Bef
  DoNotChange = False
  If LenB(Bef2) <> 0 Then
   Set Ctrl = Me.Command2
   Ctrl.Visible = True
   On Error Resume Next
   Ctrl.Height = .CellHeight  'minus 10 so that grid lines
   On Error GoTo fehler
'  Ctrl.Width = .CellWidth  '  will not be overwritten
   Ctrl.Left = .CellLeft + .CellWidth + .Left + Me.Command1.Width
   Ctrl.Top = .CellTop + .Top - 10
   DoNotChange = True
   Ctrl.Caption = Bef2
   DoNotChange = False
  End If
  .Col = altCol
 End With
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Befehl/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde (Me)
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Befehl

Private Sub aCtl_Click()
 If obdebug Then Debug.Print "aCtl_Click("
End Sub ' aCtl_Click()

Private Sub Beenden_Click()
 ProgEnde (Me)
End Sub ' Beenden_Click()

Private Sub Cb1_Click()
 On Error GoTo fehler
 If obdebug Then Debug.Print "cb1_click("
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Cb1_Click/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde (Me)
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' CB1_Click

Private Sub Cb1_GotFocus()
 aCtl = Cb1.name
End Sub ' Cb1_GotFocus()

Private Sub Cb1_KeyDown(KeyCode As Integer, Shift As Integer)
  Static obucMDIKeys%
  On Error GoTo fehler
 If obdebug Then Debug.Print "Cb1_KeyDown(", KeyCode, Shift
 If KeyCode = 18 Or KeyCode = 17 Then
 Else
  Call Key(KeyCode, Shift, Me, ActiveControl)
  If Not obucMDIKeys Then
   If KeyCode = 9 Then
    Call StDirekt("{RIGHT}", Shift)
   End If
   obucMDIKeys = True
  End If
 End If
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Cb1_KeyDown/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde (Me)
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub 'Cb1_KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub Cb1_KeyPress(KeyAscii As Integer)
'  If KeyAscii = 9 Then Stop
End Sub ' Cb1_KeyPress

Private Sub Command2_GotFocus()
 aCtl = Command2.name
End Sub ' Command2_GotFocus

Private Sub Command1_GotFocus()
 aCtl = Command1.name
End Sub ' Command1_GotFocus

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
 If obdebug Then Debug.Print "Command2_KeyDown(", KeyCode, Shift
 If KeyCode = 18 Or KeyCode = 17 Or KeyCode = 16 Then
 Else
  Call Key(KeyCode, Shift, Me, ActiveControl)
 End If
End Sub ' Command2_KeyDown

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
 If obdebug Then Debug.Print "Command1_KeyDown(", KeyCode, Shift
 If KeyCode = 18 Or KeyCode = 17 Or KeyCode = 16 Then
 Else
  Call Key(KeyCode, Shift, Me, ActiveControl)
 End If
End Sub ' Command1_KeyDown

Private Sub Datei_Click()
 If obdebug Then Debug.Print "Datei_Click("
End Sub ' Datei_Click

Private Sub Datenbank_Click()
 Call wähleBenutzer
 Call dbv.Auswahl("dp", tbm(tbwp), tbm(tbdp)) '(vns, "--multi", vns)
 Call MFGRefresh(Me.MfGTyp)
' Unload Me
' Me.Show
End Sub ' Datenbank_Click

Private Sub KürzelAnzeigenWechseln_Click()
 obMSH2 = Not obMSH2
' Call schließen
 Call MFGRefresh(azgdp)
End Sub ' KürzelAnzeigenWechseln_Click()

Private Sub Neuberechnen_für_Mitarbeiter_Click()
 Call tuNeuBerechnen
End Sub ' Neuberechnen_für_Mitarbeiter_Click()

'Urlaub für Mitarbeiter anzeigen
Private Sub Urlaubausgeben_Click()
 Call tuAusgeben(nurU:=True)
End Sub ' Urlaubausgeben_Click()

' Daten für Mitarbeiter ausgeben
Private Sub DatenAusgeben_Click()
 Call tuAusgeben
 Me.MFG.SetFocus
End Sub ' Datenausgeben_Click

' in Neuberechnen_für_Mitarbeiter_Click
Sub tuNeuBerechnen()
 Dim Persnr&, Kuerzel$
 If maAusWahl(Persnr, Kuerzel, "Neuberechnen") Then
   Screen.MousePointer = vbHourglass
   Call gesBilanz(Persnr, Kuerzel, 0, 0, obschreib:=False)
   Screen.MousePointer = vbNormal
 End If ' maAusWahl(Persnr, Kuerzel, "Neuberechnen", aus, Nachname, Vorname) Then
End Sub ' tuNeuBerechnen

' in tuNeuBerechnen und tuAusgeben
Function maAusWahl%(Persnr&, Kuerzel$, Titel$, Optional aus As Date, Optional Nachname$, Optional Vorname$)
 Dim Marb$, Vgb$
 Dim ma As ADODB.Recordset
 On Error Resume Next
 Vgb = kue(MFG.Col - 1)
 On Error GoTo fehler
 Marb = InputBox(Titel & " für Mitarbeiter: ", "Rückfrage", Vgb)
 If Marb = "" Then Exit Function
 If Not IsNumeric(Marb) Then
  Dim mrs As New ADODB.Recordset
'  dbv.wCn.Close
'  dbv.wCn.Open
'  mrs.Open "SELECT persnr FROM `" & tbm(tbma) & "` WHERE kuerzel='" & marb & "' OR nachname='" & marb & "' ORDER BY persnr DESC", dbv.wCn, adOpenStatic, adLockReadOnly
  myFrag mrs, "SELECT persnr FROM `" & tbm(tbma) & "` WHERE kuerzel='" & Marb & "' OR nachname='" & Marb & "' ORDER BY persnr DESC", adOpenStatic, dbv.wCn, adLockReadOnly
  If mrs.BOF Then
   MsgBox Marb & " nicht gefunden"
   Marb = ""
  Else
   Marb = mrs!Persnr
  End If ' mrs.BOF
 End If ' not IsNumeric(Marb)
 If LenB(Marb) <> 0 Then
'  Me.MFGRefresh (azgdp)
'  Me.MFG.SetFocus
  DoEvents
'  Me.Show
'  Call ma.Open("SELECT * FROM `" & tbm(tbma) & "` WHERE `nachname` = '" & marb & "' ORDER BY persnr DESC", dbv.wCn, adOpenStatic, adLockReadOnly)
  myFrag ma, "SELECT * FROM `" & tbm(tbma) & "` WHERE `nachname` = '" & Marb & "' ORDER BY persnr DESC", adOpenStatic, dbv.wCn, adLockReadOnly
  If ma.EOF Then
   Set ma = Nothing
'   Call ma.Open("SELECT * FROM `" & tbm(tbma) & "` WHERE `kuerzel` = '" & marb & "' ORDER BY persnr DESC", dbv.wCn, adOpenStatic, adLockReadOnly)
   myFrag ma, "SELECT * FROM `" & tbm(tbma) & "` WHERE `kuerzel` = '" & Marb & "' ORDER BY persnr DESC", adOpenStatic, dbv.wCn, adLockReadOnly
   If ma.EOF Then
    Set ma = Nothing
'    Call ma.Open("SELECT * FROM `" & tbm(tbma) & "` WHERE `persnr` = '" & marb & "' ORDER BY persnr DESC", dbv.wCn, adOpenStatic, adLockReadOnly)
    myFrag ma, "SELECT * FROM `" & tbm(tbma) & "` WHERE `persnr` = '" & Marb & "' ORDER BY persnr DESC", adOpenStatic, dbv.wCn, adLockReadOnly
   End If
  End If ' ma.EOF
  If Not ma.EOF Then
   Persnr = ma!Persnr
   Kuerzel = ma!Kuerzel
   Nachname = ma!Nachname
   Vorname = ma!Vorname
   If IsNull(ma!aus) Then aus = 0 Else aus = ma!aus
   maAusWahl = True
  Else ' Not ma.EOF Then
   MsgBox "Mitarbeiter '" & Marb & "' nicht in der Datenbank gefunden!"
  End If ' Not ma.EOF Then else
 End If ' LenB(marb) <> 0 Then
 Exit Function
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in maAusWahl/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde (Me)
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' maAusWahl%(Persnr&, Kuerzel$, Titel$, Optional aus As Date, Optional Nachname$, Optional Vorname$)

Sub tuAusgeben(Optional nurU% = False) ' nur Urlaub
 Dim aus As Date, Persnr&, Nachname$, Vorname$, Kuerzel$
 Dim wp As ADODB.Recordset
 Dim ar As ADODB.Recordset
 Dim abdat As Date, bisdat As Date
 Dim Dt$
 If maAusWahl(Persnr, Kuerzel, IIf(nurU, "Urlaub", "Daten") & " ausgeben", aus, Nachname, Vorname) Then
   Screen.MousePointer = vbHourglass
   If aus = 0 Or aus = #12:00:00 AM# Then bisdat = CDate("1.1." & Year(Now()) + 1) - 1 Else bisdat = aus
   On Error Resume Next
   Dt = Environ("userprofile") & "\eigene dateien\" & IIf(nurU, "Urlaubs", "Dienstplan") & "ausgabe_" & Nachname & "_" & Date & "_" & REPLACE$(Time, ":", ".") & ".html"
   Open Dt For Output As #323
   If Err.Number <> 0 Then
    Dt = Environ("userprofile") & "\documents\" & IIf(nurU, "Urlaubs", "Dienstplan") & "ausgabe_" & Nachname & "_" & Date & "_" & REPLACE$(Time, ":", ".") & ".html"
    Open Dt For Output As #323
   End If
   On Error GoTo fehler
   Print #323, "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">"
   Print #323, "<HTML>"
   Print #323, "<HEAD>"
   Print #323, "    <META HTTP-EQUIV=""CONTENT-TYPE"" CONTENT=""text/html; charset=windows-1252"">"
   Print #323, "    <TITLE></TITLE>"
   Print #323, "    <META NAME=""GENERATOR"" CONTENT=""Dienstplanprogramm"">"
   Print #323, "    <META NAME=""CREATED"" CONTENT=""0;0"">"
   Print #323, "    <META NAME=""CHANGEDBY"" CONTENT=""Gerald Schade"">"
   Print #323, "    <META NAME=""CHANGED"" CONTENT=""" & Format(Now(), "yyyymmDd;1284937") & """>"
   Print #323, "    <STYLE TYPE=""text/css"">"
   Print #323, "    <!--"
   Print #323, "        @page { size: 21cm 29.7cm; margin: 2cm }"
   Print #323, "        P { margin-bottom: 0.21cm }"
   Print #323, "    -->"
   Print #323, "    </STYLE>"
   Print #323, "</HEAD>"
   Print #323, "<BODY LANG=""de-DE"" DIR=""LTR"">"
   Print #323, "<PRE><span style='position:fixed;top:0px;background:#FFCC99'><B>" & IIf(nurU, "Urlaubs", "Dienstplan") & "einträge</B> von <B>" & Nachname & ", " & Vorname & "</B> (<B>" & Kuerzel & "</B>), Persnr. " & Persnr & IIf(aus = 0, "", ", Austritt: " & aus) & ", vom " & Now() & ", Dienstplanprogramm-Version: " & App.Major & "." & App.Minor & "." & App.Revision
'   Print #323, "PersNr:  " & Chr(9) & PersNr
'   Print #323, "Kürzel:  " & Chr(9) & Kuerzel
'   Print #323, "Nachname:" & Chr(9) & "<B>" & NachName & "</B>"
'   Print #323, "Vorname: " & Chr(9) & "<B>" & VorName & "</B>"
'   Print #323, "Austritt:" & Chr(9) & aus
   Print #323, "<B>" & IIf(nurU, "     ", "") & "Tag      " & Chr(9) & IIf(nurU, "         ", "") & "Ursp" & Chr(9) & "Art" & Chr(9) & "Std." & Chr(9) & "Url." & Chr(9) & "Urlstd." & Chr(9) & "Urlh/d" & Chr(9) & "Üst" & Chr(9) & "Fbdg." & Chr(9) & "Kkeintr." & Chr(9) & "geändert" & "</span>"
   Print #323, "<b><b>"
   Call gesBilanz(Persnr, Kuerzel, abdat, bisdat, obschreib:=True, nurU:=nurU)
'   Call wp.Open("SELECT * FROM `" & tbm(tbwp) & "` WHERE persnr = " & PersNr & " ORDER BY ab", dbv.wCn, adOpenStatic, adLockOptimistic)
   myFrag wp, "SELECT * FROM `" & tbm(tbwp) & "` WHERE persnr = " & Persnr & " ORDER BY ab", adOpenStatic, dbv.wCn, adLockOptimistic
   If Not wp.BOF Then
    Print #323, vbCrLf & "<B>Wochenarbeitszeit</B>:" & vbCrLf & "gültig ab" & Chr(9) & "Mo" & Chr(9) & "Di" & Chr(9) & "Mi" & Chr(9) & "Do" & Chr(9) & "Fr" & Chr(9) & "Sa" & Chr(9) & "So" & Chr(9) & "Wo'St." & Chr(9) & "Urlaubstage/Jahr"
    Do While Not wp.EOF
     If abdat = 0 Then abdat = wp!ab
     Print #323, Round(wp!ab, 1) & Chr(9) & wp!mo & Chr(9) & wp!di & Chr(9) & wp!mi & Chr(9) & wp!Do & Chr(9) & wp!fr & Chr(9) & wp!sa & Chr(9) & wp!so & Chr(9) & wp!WAZ & "  " & Chr(9) & wp!urlaub
     wp.Move 1
    Loop
   End If ' Not wp.BOF Then
   If Not nurU Then
'    Call ar.Open("SELECT * FROM `" & tbm(tbar) & "` WHERE zusatz = 0", dbv.wCn, adOpenStatic, adLockOptimistic)
    myFrag ar, "SELECT * FROM `" & tbm(tbar) & "` WHERE zusatz = 0", adOpenStatic, dbv.wCn, adLockOptimistic
    If Not ar.BOF Then
     Print #323, vbCrLf & "<B>Dienstplanarten</B>:" & vbCrLf & "ArtNr" & Chr(9) & Left("Erklärung" & Space(22), 22) & Chr(9) & "Stunden"
     Do While Not ar.EOF
      Print #323, ar!artnr & Chr(9) & Left(ar!erkl & Space(22), 22) & Chr(9) & ar!Stdn
      ar.Move 1
     Loop
    End If ' Not ar.BOF Then
   End If ' Not nurU Then
   Print #323, "   </PRE>"
   Print #323, "</BODY>"
   Print #323, "</HTML>"
   Close #323
'   Call Shell("notepad " & Dt, vbMaximizedFocus)
   If 1 = 1 Then
    ' Leeres Browser-Fenster öffnen
'     URLGoTo Me.hWnd, "about:blank"
    ' Jetzt gewünschte Webseite laden
     URLGoTo Me.hwnd, Dt
   ElseIf 1 = 0 Then
    GetWord
    Wapp.Documents.Open Dt
    Wapp.Activate
   Else
    Dim oDoc, arg()
    GetOpenOffice
    Set oDoc = Wapp.loadComponentFromURL("file:///" & REPLACE(Dt, "\", "/"), "_blank", 0, arg())
    Wapp.Activate
   End If ' 1 = 1 else
  Screen.MousePointer = vbNormal
 End If ' maAusWahl(Persnr, Kuerzel, Titel, aus, Nachname, Vorname) Then
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in DatenAusgeben_Click/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde (Me)
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Datenausgeben_Click

' aufgerufen in Neuberechnen_Click, tuNeuBerechnen und tuAusgeben
Function gesBilanz(Persnr&, Kuerzel$, abdat As Date, bisdat As Date, Optional obschreib%, Optional nurU%)
   Dim UBilh!, UBil!, ÜBil!, FBil!, altUBilh!, altUBil!, altÜBil!, altFBil!, ArtVgb$, alttag As Date, PSt! ' ... Planstunden
   Dim akttag As Date, ArtDp$, aktjahr%
   Dim WAZ!, WAZv! ' Wochenarbeitszeit des Vortags
   Dim dp As ADODB.Recordset, abd As ADODB.Recordset, zuloe As ADODB.Recordset
   Dim aktstd!, iststd!, i%, UrlTagNr%, ausbez!, urlhaus!
   On Error GoTo fehler:
   syscmd 4, "Berechne neu: " & Kuerzel
   WAZ = 0
   alttag = 0
   If abdat = 0 Then
'    abdat = dbv.wCn.Execute("SELECT IF(MIN(ab)=0 OR ISNULL(MIN(ab)),DATE(20040701),MIN(ab)) FROM `" & tbm(tbwp) & "` WHERE persnr=" & CStr(PersNr)).Fields(0)
'    Set abd = Nothing
    myFrag abd, "SELECT IF(MIN(ab)=0 OR ISNULL(MIN(ab)),DATE(20040701),MIN(ab)) FROM `" & tbm(tbwp) & "` WHERE persnr=" & CStr(Persnr), , dbv.wCn
    abdat = abd.Fields(0)
   End If
   If bisdat = 0 Then bisdat = dbv.wCn.Execute("SELECT IF(aus>0,aus,SUBDATE(CAST(CONCAT(YEAR(NOW())+4,'0101') as date),1)) FROM `" & tbm(tbma) & "` WHERE persnr=" & CStr(Persnr)).Fields(0)
'   dbv.wCn.Execute ("DELETE FROM `" & tbm(tbbi) & "` WHERE persnr = " & CStr(PersNr))
fangnochmalan:
   Set zuloe = Nothing
   myFrag zuloe, "DELETE FROM `" & tbm(tbbi) & "` WHERE persnr = " & CStr(Persnr), , dbv.wCn
'   If persnr = 83 Then Stop
   akttag = abdat
   sql = "SELECT d.tag,d.artnr,a.Stdn,COALESCE(z.ausbez,0) ausbez, COALESCE(z.urlhaus,0) urlhaus " & vbCrLf & _
         "FROM `" & tbm(tbdp) & "` d " & vbCrLf & _
         "LEFT JOIN ausbez z USING (persnr,tag) " & vbCrLf & _
         "LEFT JOIN `" & tbm(tbar) & "` a ON d.artnr = a.artnr " & vbCrLf & _
         "WHERE persnr = " & Persnr & " GROUP BY d.tag ORDER BY d.tag"
'   sql = "SELECT * FROM `" & tbm(tbdp) & "` d LEFT JOIN `" & tbm(tbar) & "` a on d.artnr = a.artnr WHERE persnr = " & CStr(persnr) & " AND tag = " & Format(akttag, "yyyymmdd")
   Set dp = Nothing
'    If akttag = #1/1/2012# Then Stop
'    dbv.wCn.CursorLocation = adUseServer
'   Call dp.Open(sql, dbv.wCn, adOpenStatic, adLockReadOnly)
   If dbv.wCn.DefaultDatabase = "" Then
    dbv.wCn.Close
    dbv.wCn.Open
   End If
   myFrag dp, sql, adOpenStatic, dbv.wCn, adLockReadOnly
'    dp.Index = "tag"
   Do While akttag <= bisdat
   syscmd 4, "Berechne neu: " & Kuerzel & ", " & akttag
'    If akttag = tzsbd Then Stop
    If Year(akttag) <> aktjahr Then
     aktjahr = Year(akttag)
     UrlTagNr = 0
     Call FTbeleg(aktjahr)
    End If
'    dp.Seek Format(akttag, "yyyymmdd")
'    If akttag = #6/11/2019# Then Stop
    ausbez = 0
    urlhaus = 0
    If dp.State = 0 Then
     GoTo fangnochmalan
    End If
    If Not dp.EOF Then
     If akttag > dp!Tag Then
      dp.MoveNext
      If Not dp.EOF Then
       ausbez = dp!ausbez
'       If ausbez Then Stop
       urlhaus = dp!urlhaus
      End If
     End If
    End If ' Not dp.EOF Then
'    If dp.EOF AND (Weekday(akttag) = 7 OR Weekday(akttag) = 1) Then GoTo weiter 'Sa oder So, wenn keine Überstunden da
'    For i = 0 To UBound(ftag) ' Feiertage
'     If ftag(i).Datum = akttag Then
''      If Not ftag(i).obhalb Then GoTo weiter
'     End If
'    Next i
'   If Not dp.BOF Then
'    Do While Not dp.EOF
'     akttag=dp!tag
     If Not dp.EOF Then
      If akttag = dp!Tag Then
       ArtDp = dp!artnr
       aktstd = dp!Stdn
       If ArtDp = "g" Or ArtDp = "u" Or ArtDp = "uw" Then UrlTagNr = UrlTagNr + 1
       GoTo w0:
      End If
     End If
     ArtDp = ""
     aktstd = 0
w0:
     iststd = aktstd
     If obschreib Then
      altUBilh = UBilh
      altUBil = UBil ' diese drei zum Fettformatieren von Änderungen
      altÜBil = ÜBil
      altFBil = FBil
     End If
     ArtVgb = ""
     
'     If Not dp.EOF Then ' wegen der Prüfung in der nächsten Zeile nötig
'      If akttag = dp!Tag Then
''      If marb = 45 AND YEAR(akttag) = 2006 Then
''       Open "v:\P1.txt" For Append As #196
''       Print #196, ArtVgb, ArtDp, UBil, ÜBil, FBil, PSt, WAZ, akttag
''       Close #196
''      End If
'       If PersNr = 82 Then Stop
'     If ausbez Then Stop
     Call EinzelBilanz(Persnr, ArtVgb, ArtDp, ausbez, urlhaus, UBilh, UBil, ÜBil, FBil, PSt, akttag, WAZv, mitdruck:=obschreib)
'       GoTo w1:
'      End If
'     End If
     If LenB(ArtDp) = 0 Then If IsNumeric(ArtVgb) Then iststd = CDbl(ArtVgb)
w1:
'    If Month(akttag) = 12 And Day(akttag) = 31 Then Stop
    If obschreib Then
     Dim reint As ADODB.Recordset, rgeae As ADODB.Recordset
     'If Int(dp!Tag) = #4/14/2020# Then Stop
     Dim sqlteil$, reintct&, rgeaegeae
'     sqlteil = " FROM quelle.eintraege WHERE zeitpunkt BETWEEN " & Format(akttag, "yyyymmdd") & " AND " & Format(akttag + 1, "yyyymmdd") & " AND (art = '" & kuerzel & "' OR ((inhalt COLLATE latin1_bin LIKE '%Mitarbeiter: " & kuerzel & "%' OR inhalt COLLATE latin1_bin LIKE '% " & kuerzel & "' OR inhalt COLLATE latin1_bin LIKE '% (" & kuerzel & ")') AND NOT inhalt COLLATE latin1_bin LIKE '%mit " & kuerzel & "%'))"
     sqlteil = " FROM quelle.eintraege WHERE zeitpunkt BETWEEN " & Format(akttag, "yyyymmdd") & " AND " & Format(akttag + 1, "yyyymmdd") & " AND (art = '" & Kuerzel & "' OR ((inhalt LIKE '%Mitarbeiter: " & Kuerzel & "%' OR inhalt LIKE '% " & Kuerzel & "' OR inhalt LIKE '% (" & Kuerzel & ")') AND NOT inhalt LIKE '%mit " & Kuerzel & "%'))"
'     reint.Open "SELECT COUNT(0) ct" & sqlteil, dbv.wCn, adOpenStatic, adLockReadOnly
     myFrag reint, "SELECT COUNT(0) ct" & sqlteil, adOpenStatic, dbv.wCn, adLockReadOnly
     reintct = reint!ct
     Dim obfalsch%
     obfalsch = 0
     ' Fehler können auch entstehen durch:
     ' Nutzen einer vorbestehenden Leerzeile für aktuellen Eintrag,
     ' Vertippen bei der Mitarbeiterauswahl,
     ' ungenau datierte Schulungen
     Const biszahlweg% = 2
     Const abzahlda% = 1
     Dim pid$
     pid = ""
'     If Not dp.EOF Then ' wegen der Prüfung in der nächsten Zeile nötig
'      If akttag = dp!Tag Then
       If akttag < Int(Now()) Then ' And (ArtDp <> "k") Then
        If (reint!ct <= biszahlweg And ArtDp <> "b" And iststd <> 0) Then
         obfalsch = 1
        ElseIf (reint!ct >= abzahlda And (ArtDp = "b" Or iststd = 0)) Then
'        rpid.Open "SELECT GROUP_CONCAT(DISTINCT pat_id ORDER BY pat_id) pid " & sqlteil, dbv.wCn, adOpenStatic, adLockReadOnly
'        pid = "; Patid: " & rpid!pid
         pid = dbv.wCn.Execute("SELECT GROUP_CONCAT(CONCAT('\n         ',lpad(pat_id,5,' '),' ',DATE_FORMAT(zeitpunkt,'%d.%m.%y %H:%i '),art,' ',inhalt)) pid " & sqlteil)!pid
         obfalsch = 2
        End If
       End If
'      End If
'     End If
'     rgeae.Open "SELECT GROUP_CONCAT(CONCAT(aenddat,': ''',artnr,''' (',user,' ',aendpc,')') ORDER BY aenddat desc separator '; ') geae FROM dp.`" & tbm(tbpr) & "` p WHERE persnr=" & PersNr & " AND tag= " & Format(akttag, "yyyymmdd"), dbv.wCn, adOpenStatic, adLockReadOnly
     myFrag rgeae, "SELECT COALESCE(GROUP_CONCAT(CONCAT(aenddat,': ''',artnr,''' (',user,' ',aendpc,')') ORDER BY aenddat DESC SEPARATOR '; '),'') geae FROM dp.`" & tbm(tbpr) & "` p WHERE persnr=" & Persnr & " AND tag= " & Format(akttag, "yyyymmdd"), adOpenStatic, dbv.wCn, adLockReadOnly
     rgeaegeae = rgeae!geae
     If Not nurU Or ArtDp Like "g *" Or ArtDp Like "u *" Or ArtDp Like "uw *" Then
      Dim Stil$
      If obfalsch Or ArtDp <> "" Then
       If obfalsch = 1 Then
        Stil = "style='background:#FF99CC'" ' rot
       ElseIf obfalsch = 2 Then
        Stil = "style='background:#CC99FF'" ' lila
       Else
        Stil = "style='background:#E6E6E6'" ' hellgrau
       End If
      ElseIf akttag = bisdat Then ' Ende
       Stil = "style='color:#FFFAF0;background:#000000'"
      ElseIf UCase$(ArtVgb) = "WF" Then
       Stil = "style='color:#A9A9A9'" ' hellgrau
      Else
       Stil = ""
      End If
      Dim divisor$
      If UBil = 0 Then
       divisor = "entf."
      Else
       divisor = Round(UBilh / UBil, 1)
      End If
      Print #323, IIf(nurU, Format(UrlTagNr, "000) "), "") & IIf(Stil <> "", "<span " & Stil & ">", "") & Format(akttag, "ddd, dd.mm.yy") & Chr(9) & ArtVgb & _
       Chr(9) & ArtDp & Chr(9) & iststd & Chr(9) & IIf(altUBil = UBil, vNS, "<B style=""color:red;"">") & Round(-UBil, 1) & IIf(altUBil = UBil, vNS, "</B>") & Chr(9) & IIf(altUBilh = UBilh, vNS, "<B style=""color:red;"">") & Round(-UBilh, 1) & IIf(altUBilh = UBilh, vNS, "</B>") & Chr(9) & divisor & Chr(9) & IIf(altÜBil = ÜBil, vNS, "<B style=""color:Tomato;"">") & ÜBil & IIf(altÜBil = ÜBil, vNS, "</B>") & Chr(9) & IIf(altFBil = FBil, vNS, "<B style=""color:Tomato;"">") & -FBil & IIf(altFBil = FBil, vNS, "</B>") & Chr(9) & reintct & Chr(9) & rgeaegeae & pid & IIf(Stil <> "", "</span>", "")
     End If ' Not nurU or ...
     Set reint = Nothing
     Set rgeae = Nothing
     If (Not nurU And ausbez) Or urlhaus Then
      Print #323, Format(akttag, "ddd, dd.mm.yy") & Chr(9) & "<span style='background:#FFFF99'>ausbezahlt: " & ausbez & " Stunden, " & urlhaus & " Urlaubsstunden</span>"
     End If
    End If ' obschreib
    If (Month(akttag) = 12 And Day(akttag) = 31) Or akttag = bisdat Then ' obschreib
     sql = "INSERT INTO `" & tbm(tbbi) & "`(urlstd,urlaub,überstunden,fortbildung,persnr,jahr,planstunden) VALUES ('" & Str$(UBilh) & "','" & Str$(UBil) & "','" & Str$(ÜBil) & "','" & Str$(FBil) & "'," & Persnr & "," & Year(akttag) & ",'" & Str$(PSt) & "')" ' Apostrophe wegen krummen Zahlen nötig
     dbv.wCn.Execute sql, rAf
    End If
    alttag = akttag
'     dp.Move 1
'    Loop
'   End If
'weiter:
    akttag = akttag + 1
   Loop ' While akttag <= bisdat
   syscmd 5
 Exit Function
fehler:
 If Err.Number = -2147467259 And InStr(Err.Description, "Lost connection") <> 0 Then
  dbv.wCn.Close
  dbv.wCn.Open
  Resume
 End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GesBilanz/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde Me
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' GesBilanz

' aufgerufen in tuAufgeben
Public Sub GetOpenOffice()
 On Error GoTo fehler
  Set Wapp = getAppl("SALFRAME", "OpenOffice")
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in GetWord/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde Me
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' GetOpenOffice

Public Sub GetWord()
 On Error GoTo fehler
  Set Wapp = getAppl("OpusApp", "Word.Application")
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in GetWord/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde Me
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' GetWord

' aufgerufen in GetOpenOffice und GetWord
Public Function getAppl(className, ObjName) As Object 'Word.Application
' Test to see if there is a copy of Micr
' osoft Word already running.
'on error resume next' Defer error trapping.
' Getobject function called without the
' first argument returns a
' reference to an instance of the applic
' ation. If the application isn't
' running, an error occurs.
Dim FZahl%
On Error Resume Next
vonvorne:
Set getAppl = GetObject(, ObjName)

If Err.Number <> 0 Then
WordWasNotRunning = True
Else
WordWasNotRunning = False
End If
Err.Clear ' Clear Err object In Case Error occurred.
' Check for Microsoft Word. If Microsoft
' Word is running,
' enter it into the Running Object table
' .
On Error GoTo fehler

If WordWasNotRunning = True Then
'Set the object variable to start a new
' instance of Word.
neu:
Select Case ObjName
 Case "Word.Application"
  Set getAppl = CreateObject("Word.Application") 'wobj ' New Word.Application
 Case "OpenOffice"
  Dim zwi
  Set zwi = CreateObject("com.sun.star.ServiceManager")
  Set getAppl = zwi.createInstance("com.sun.star.frame.Desktop")
 Case Else
End Select
End If
' Show Microsoft Word through its Applic
' ation property. Then
' show the actual window containing the
' file using the Windows
' collection of the MyWord object refere
' nce.
On Error Resume Next
getAppl.Visible = True
If Err.Number <> 0 And FZahl < 10 Then
 FZahl = FZahl + 1
 GoTo vonvorne
End If
getAppl.Application.WindowState = wdWindowStateMaximize
Select Case Err.Number
 Case 0
 Case 5825: GoTo neu
End Select
Exit Function
On Error GoTo fehler
Screen.MousePointer = 0 ' vbDefault
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in getAppl/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde Me
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' getAppl
'Demo of how to call the above sub

Private Sub DatenSpeichern_Click()
 Dim erg$
 erg = InputBox("Daten speichern in Verzeichnis: ")
 If LenB(erg) <> 0 Then
  erg = machCSVs(dbv.wCn.ConnectionString, erg)
  MsgBox "Daten von " & Chr(34) & dbv.wCn.Properties("Server Name") & Chr(34) & "/" & dbv.wCn.DefaultDatabase & vbCrLf & " in " & erg & " gespeichert"
 End If
End Sub ' DatenSpeichern_Click()

Private Sub dbv_wCnAendern(CnStr As String)
  On Error Resume Next
  Me.cnLab = "Programm: '" & App.ProductName & "', Version: " & App.Major & "." & App.Minor & "." & App.Revision & ", Datenbank: '" & dbv.wCn.DefaultDatabase & "' auf '" & dbv.wCn.Properties("server name") & "', akt. Steuerelement: "
  Me.cnLab.ToolTipText = App.Path & "\" & App.EXEName & ".exe"
  Me.cnLab.Alignment = vbRightJustify
End Sub ' dbv_wCnAendern

Private Sub Erneuern_Click()
 If obdebug Then Debug.Print "Erneuern_Click("
 Call Grundausstatt
End Sub ' Erneuern_Click()

Private Sub wähleBenutzer()
' dbv.Ü2 = "Benutzer"
 dbv.rücksetzBedTbl
 Call dbv.setzBedTbl(tbm(tbar))
 Call dbv.setzBedTbl(tbm(tbma))
 Call dbv.setzBedTbl(tbm(tbdp))
End Sub ' wähleBenutzer()

Private Sub Farbauswahl1_Click()
 If obdebug Then Debug.Print "Farbauswahl_Click("
'Dim pot_filename$
'Dim tZeile$
FmCD.CmDlg.CancelError = True
'fmcd.cmdlg.Filter = "Potentiale(*.pot)|*.pot" 'Voreinstellung der Dateiart
'fmcd.cmdlg.Flags = &H1000& 'vgl. obige Flag-Tabelle
FmCD.CmDlg.Action = 3 'Aufruf des Dialogfensters. Es ist nur noch dieses ansprechbar, bis es mit Abbruch oder OK geschlossen wird.
On Error Resume Next
FmCD.CmDlg.ShowColor
On Error GoTo 0
'pot_filename = fmcd.cmdlg.FileName 'Das Fenster ist wieder geschlossen, die ausgewählte Datei wird an eine Variable übergeben
'Open pot_filename For Input As 100 'Öffnen der ausgewählten Datei, es folgt die selbst zu gestaltende Leseroutine
'Input #100, tZeile 'hier nur eine Textzeile, die in die Variable des Namens tZeile eingelesen wird.
'Close #100 'Datei schließen
End Sub ' Farbauswahl1_Click

Private Sub Jahr_Click()
 If obdebug Then Debug.Print "Jahr_Click("
End Sub ' Jahr_Click

Private Sub loggen_Click()
 If LenB(User) = 0 Then
  Call prüfeUser
 Else
  User = vNS
 End If
 If LenB(User) <> 0 Then
  MDI.loggen.Caption = User & " auslo&ggen"
 Else
  MDI.loggen.Caption = "einlo&ggen"
 End If
End Sub ' loggen_Click

Private Sub MDIForm_Unload(Cancel As Integer)
 Call merken(MfGTyp)
 MfGTyp = azgnix
 ProgEnde Me
End Sub ' MDIForm_Unload

Private Sub MFG_DblClick()
'-------------------------------------------------------------------------------------------
' Code in DblClick-Ereignis der Tabelle aktiviert Spaltensortierung
'-------------------------------------------------------------------------------------------

    Dim i As Integer

    ' nur dann sortieren, wenn eine feste Zeile angeklickt wurde
    If MFG.MouseRow >= MFG.FixedRows Then Exit Sub

    i = m_iSortCol                  ' alte Spalte speichern
    m_iSortCol = MFG.Col   ' neue Spalte festlegen

    ' Sortiertyp inkrementieren
    If i <> m_iSortCol Then
        ' wenn eine neue Spalte geklickt wird, mit aufsteigender Sortierung beginnen
        m_iSortType = 1
    Else
        ' wenn dieselbe Spalte geklickt wird, zwischen aufsteigender und absteigender Sortierung umschalten
        m_iSortType = m_iSortType + 1
    If m_iSortType = 3 Then m_iSortType = 1
    End If

    DoColumnSort

End Sub ' MFG_DblClick

Sub DoColumnSort()
'-------------------------------------------------------------------------------------------
' Führt Exchange-Sortierung von column m_iSortCol durch
'-------------------------------------------------------------------------------------------

    With MFG
        .Redraw = False
        .Row = 1
        .RowSel = .Rows - 1
        .Col = m_iSortCol
        .Sort = m_iSortType
        .Redraw = True
    End With

End Sub 'DpoColumnSort

Private Sub MFG_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
 Dim c%, summe&, aktCol&
 For c = 0 To MFG.Cols - 1
  summe = summe + MFG.ColWidth(c)
 Next
 If x < summe Then
'  Call MFG_Click
 Else
  MFG_click_abbrech = True
 End If
 Dim rh2#
 On Error Resume Next
 rh2 = 0
 rh2 = MFG.RowHeight(2)
 On Error GoTo 0
 If Y > MFG.RowHeight(0) + MFG.RowHeight(1) And Y < MFG.RowHeight(0) + MFG.RowHeight(1) + rh2 Then
  ' -> Überstunden aktuell auflisten
 ElseIf Y < MFG.RowHeight(0) Then
  summe = 0
  For c = 0 To MFG.Cols - 1
   summe = summe + MFG.ColWidth(c)
   If x < summe Then
    aktCol = c
    Exit For
   End If
  Next c
  If MfGTyp = azgdp Then
   merkCol(MfGTyp) = MFG.Col
   merkRow(MfGTyp) = MFG.Row
   Einschr(azgwp) = "WHERE `PersNr` = " & pn(aktCol - 1)
   EinsFd(azgwp) = "KalSp4"
   EinsWt(azgwp) = "'" & pn(aktCol - 1) & "'"
   EinsNm(azgwp) = "Wochenplan für " & MFG.TextMatrix(0, aktCol) '& ", " & MFG.TextMatrix(MFG.Row, VNCol)
   Call merken2(MfGTyp)
   altMfgTyp = MfGTyp
   MfGTyp = azgnix
   Call Me.machWochenplan(Einschr(azgwp), EinsNm(azgwp))
  Else
  End If ' MfgTyp = azgdp else
 End If ' Y  > MFG.RowHeight(0) ... else
End Sub ' MouseDown

Private Sub MSH2_Click()
 If obdebug Then Debug.Print "MSH2_Click("
End Sub ' MSH2_Click()

' neu berechnen
Private Sub NeuBerechnen_Click()
' Dim ArtNamen$(), UBilh!(), UBil!(), ÜBil!(), FBil!(), PSt!(), ArtZ%, maZ%, aktJ%, rAf& ' jahrZ%,
 Dim maZ%
 Dim Statistik%(), obAus%
 Dim wpln$(1 To 7), j%, maxJahr%
 Dim aktDat As Date
 ' rsVarben
 Dim rsv As ADODB.Recordset, ma As ADODB.Recordset ', wp As New ADODB.Recordset, dp As New ADODB.Recordset
 Const BerDebug% = 0
 
 DoEvents
 Screen.MousePointer = vbHourglass
 DoEvents
' ReDim ArtNamen(0)
 
' Set rsv = Nothing
'' rsv.Open "SELECT YEAR(max(tag)) as maxJahr FROM `" & tbm(tbdp) & "`", dbv.wCn, adOpenDynamic, adLockReadOnly '  desc
' myFrag rsv, "SELECT YEAR(max(tag)) as maxJahr FROM `" & tbm(tbdp) & "`", adOpenStatic, dbv.wCn, adLockReadOnly
' maxJahr = rsv!maxJahr
' Set rsv = Nothing
' rsv.Open "SELECT COUNT(artnr) FROM `" & tbm(tbar) & "` WHERE zusatz = 0", dbv.wCn, adOpenDynamic, adLockReadOnly '  desc
' myFrag rsv, "SELECT COUNT(artnr) FROM `" & tbm(tbar) & "` WHERE zusatz = 0", adOpenStatic, dbv.wCn, adLockReadOnly
' If Not rsv.EOF Then ArtZ = rsv.Fields(0)
' Set rsv = Nothing
'' rsv.Open "SELECT COUNT(*) FROM `" & tbm(tbma) & "`", dbv.wCn, adOpenStatic, adLockOptimistic
 myFrag rsv, "SELECT COUNT(0) FROM `" & tbm(tbma) & "`", adOpenStatic, dbv.wCn, adLockReadOnly
 If Not rsv.EOF Then maZ = rsv.Fields(0)
' Set rsv = Nothing
'' rsv.Open "SELECT YEAR(min(ab)) FROM `" & tbm(tbwp) & "`", dbv.wCn, adOpenStatic, adLockReadOnly
' myFrag rsv, "SELECT YEAR(MIN(ab)) FROM `" & tbm(tbwp) & "`", adOpenStatic, dbv.wCn, adLockReadOnly

' If Not rsv.EOF Then jahrZ = YEAR(Now) - rsv.Fields(0) + 1
 
' ReDim ArtNamen(ArtZ)
' ReDim Statistik(maZ, rsv.Fields(0) To maxJahr, ArtZ)
' ReDim UBilh(maZ, rsv.Fields(0) To maxJahr)
' ReDim UBil(maZ, rsv.Fields(0) To maxJahr)
' ReDim ÜBil(maZ, rsv.Fields(0) To maxJahr)
' ReDim PSt(maZ, rsv.Fields(0) To maxJahr)
' ReDim FBil(maZ, rsv.Fields(0) To maxJahr)
' If BerDebug Then Open "testausg07.txt" For Output As #7
' Set rsv = Nothing
'' rsv.Open "SELECT artnr FROM `" & tbm(tbar) & "` WHERE zusatz = 0", dbv.wCn, adOpenDynamic, adLockReadOnly '  desc
' myFrag rsv, "SELECT artnr FROM `" & tbm(tbar) & "` WHERE zusatz = 0", adOpenStatic, dbv.wCn, adLockReadOnly
' For j = 0 To ArtZ - 1
'  ArtNamen(j) = rsv!artnr
'  rsv.Move 1
' Next j
 
 Set ma = Nothing
' ma.Open "SELECT * FROM `" & tbm(tbma) & "`", dbv.wCn, adOpenStatic, adLockOptimistic
 myFrag ma, "SELECT * FROM `" & tbm(tbma) & "`", adOpenStatic, dbv.wCn, adLockReadOnly
 For j = 0 To maZ - 1
  If True Or ma!Persnr = 70 Then
'  If True Then
   
   Call gesBilanz(ma!Persnr, ma!Kuerzel, 0, 0, obschreib:=False)
   GoTo w1
  
'  If ma!persnr = 83 Then Open "t:\testausg" & ma!persnr & ".txt" For Output As #323
'  Dim WAZ!, WAZv! ' Wochenarbeitszeit, ' ~ des Vortags
'  WAZ = 0
'  Set wp = Nothing
'  wp.Open "SELECT * FROM `" & tbm(tbwp) & "` WHERE `persnr` = " & ma!persnr & " AND `ab` <> 0 ORDER BY `ab`", dbv.wCn, adOpenStatic, adLockOptimistic
'  If Not wp.BOF() Then
'   aktDat = wp!ab
'   aktJ = YEAR(aktDat)
'   obAus = 0
'   Do
'    wpln(2) = IIf(IsNull(wp!mo), "0", wp!mo)
'    wpln(3) = IIf(IsNull(wp!di), "0", wp!di)
'    wpln(4) = IIf(IsNull(wp!mi), "0", wp!mi)
'    wpln(5) = IIf(IsNull(wp!Do), "0", wp!Do)
'    wpln(6) = IIf(IsNull(wp!fr), "0", wp!fr)
'    wpln(7) = IIf(IsNull(wp!sa), "0", wp!sa)
'    wpln(1) = IIf(IsNull(wp!so), "0", wp!so)
'    If wp.EOF Then WAZ = 0 Else WAZ = wp!WAZ
'    wp.Move 1
'    Do
'     If aktDat = ma!aus Or aktDat = CDate("31.12." & maxJahr) Then
'      obAus = True
'      Exit Do
'     End If
'     If Not wp.EOF Then
'      If aktDat = wp!ab Then Exit Do
'     End If
'''     Print #7, ma!persnr, aktDat, wpln(Weekday(aktDat))
'     Set dp = Nothing
'     Call dp.Open("SELECT artnr FROM `" & tbm(tbdp) & "` WHERE persnr = " & ma!persnr & " AND tag = " & datform(aktDat), dbv.wCn, adOpenStatic, adLockReadOnly)
'     If Not dp.EOF Then
'      Dim Index%
'      Index = Weekday(aktDat)
''      If ma!persnr = 45 AND YEAR(aktDat) = 2006 Then
''       Open "v:\P2.txt" For Append As #196
''       Print #196, wpln(Index), dp!artnr, UBil(j, aktJ), ÜBil(j, aktJ), FBil(j, aktJ), PSt(j, aktJ), WAZ, aktDat
''       Close #196
''      End If
'      Call EinzelBilanz(wpln(Index), dp!artnr, UBil(j, aktJ), ÜBil(j, aktJ), FBil(j, aktJ), PSt(j, aktJ), WAZ, aktDat, WAZv, IIf(ma!persnr = 83, True, False))
''      If BerDebug Then If ma!persnr = 43 Then Print #7, aktDat, ÜBil(j, aktJ), UBil(j, aktJ), FBil(j, aktJ), PSt(j, aktJ)
'      If ma!persnr = 83 Then Print #323, aktDat, ÜBil(j, aktJ), UBil(j, aktJ), FBil(j, aktJ), PSt(j, aktJ)
'     End If
'     aktDat = aktDat + 1
'     aktJ = YEAR(aktDat)
'    Loop
'    If obAus Then Exit Do
'   Loop
'   For aktJ = LBound(UBil, 2) To UBound(UBil, 2)
'    rAF = 0
'    If Not (UBil(j, aktJ) = 0 And ÜBil(j, aktJ) = 0 And FBil(j, aktJ) = 0 And PSt(j, aktJ) = 0) Then
'     sql = "INSERT INTO `" & tbm(tbbi) & "`(urlaub,überstunden,fortbildung,persnr,jahr,planstunden) values ('" & Str$(UBil(j, aktJ)) & "','" & Str$(ÜBil(j, aktJ)) & "','" & Str$(FBil(j, aktJ)) & "'," & ma!persnr & "," & aktJ & ",'" & Str$(PSt(j, aktJ)) & "')" ' Apostrophe wegen krummen Zahlen nötig
'     dbv.wCn.Execute sql, rAF
'    End If
''    If BerDebug Then Print #7, ma!persnr, aktJ, UBil(j, aktJ), ÜBil(j, aktJ), FBil(j, aktJ), PSt(j, aktJ), rAF
'   Next aktJ
''   End If
'   If ma!persnr = 83 Then Close #323
   End If ' persnr=83 then
w1:
   ma.Move 1
  Next j
'  Close #7
  Screen.MousePointer = vbNormal
  Call MFGRefresh(azgdp)
  If BerDebug Then
   Call Shell(Environ("windir") & "\notepad.exe " & "testausg07.txt", vbMaximizedFocus)
   Kill "testausg07.txt"
  End If
  MsgBox "Fertig mit Neuberechnen"
  Exit Sub
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in NeuBerechnen_Click/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Neuberechnen

Private Sub NeuZeichnen_Click()
 Call MFGRefresh(Me.MfGTyp)
End Sub ' NeuZeichnen_Click

Private Sub ProtokollAnzeigen_Click()
 Call MFGRefresh(azgpr)
End Sub ' ProtokollAnzeigen_Click

' Rechtspfeil beim Menü
Private Sub Rechts_Click()
 If obdebug Then Debug.Print "Rechts_Click("
 Me.Jahr = Me.Jahr + 1
 Me.MFG.TopRow = Me.MFG.FixedRows
 Me.MFG.Row = Me.MFG.FixedRows
 Call MFGRefresh(azgdp)
 Me.ucMDIKeys.SetFocus
End Sub ' Rechts_Click

' Linkspfeil beim Menü
Private Sub Links_Click()
 If obdebug Then Debug.Print "Links_Click("
 Me.Jahr = Me.Jahr - 1
 Me.MFG.TopRow = min(Me.MFG.Rows - 1, 322)
 Me.MFG.Row = min(Me.MFG.Rows - 1, 370)
 Call MFGRefresh(azgdp)
 Me.ucMDIKeys.SetFocus
End Sub ' Links_Click
Private Sub zuHeuteGehen_Click()
 Me.Jahr = Year(Now)
 Me.MFG.Row = DatePart("y", Now) + SZZ
 Me.MFG.TopRow = Me.MFG.Row - 10
 Call MFGRefresh(azgdp)
 Me.ucMDIKeys.SetFocus
End Sub ' zuHeuteGehen_Click()

Private Sub MDIForm_Activate()
 Me.loggen.Caption = "einlo&ggen                          - Dienstplan Ver." & App.Major & "." & App.Minor & "." & App.Revision & "; für alle Programmfehler verantwortlich: Gerald Schade"
 aCtl = Me.name
End Sub ' MDIForm_Activate

Private Sub MDIForm_Click()
 If obdebug Then Debug.Print "MDIForm_Click(", "Row:", Me.MFG.Row, "Col:", Me.MFG.Col, "noenter:", noenter, "nouc:", nouc
End Sub ' MDIForm_Click

Private Sub MFG_Entercell()
 aCtl = MFG.name
 If obdebug Then Debug.Print "mfg_entercell(", "Row:", Me.MFG.Row, "Col:", Me.MFG.Col, "noenter:", noenter, "nouc:", nouc
 If noenter = 0 Then
  If fgespei(MfGTyp) = 0 Then
   altFarbe(MfGTyp) = Me.MFG.CellBackColor
   Me.MFG.CellBackColor = vbYellow
   fgespei(MfGTyp) = -1
  End If
  If MfGTyp = azgma And Me.MFG.Rows > 2 Then
   Call Befehl("&Wochenplan", "A&usbezahlungen")
  ElseIf MfGTyp = azgar And Me.MFG.Rows > 2 Then
   Call Befehl("&Farbe", "Farbe f&ür alle gleichen")
  End If
  If nouc = 0 Then Me.ucMDIKeys.SetFocus
 End If
End Sub ' mfg_entercell()

Private Sub MFG_KeyDown(KeyCode As Integer, Shift As Integer)
 If obdebug Then Debug.Print "MFG_KeyDown(", KeyCode, Shift
 If KeyCode = 18 Or KeyCode = 17 Then
 Else
  If nouc = 0 Then Call Key(KeyCode, Shift, Me, ActiveControl)
 End If
End Sub ' MFG_KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub MFG_KeyPress(KeyAscii As Integer)
 Call Key(KeyAscii, 0, Me, MFG)
End Sub ' MFG_KeyPress(KeyAscii As Integer)

Private Sub mfg_leavecell()
 If obdebug Then Debug.Print "mfg_leavecell("
 If noenter = 0 Then
  Me.MFG.CellBackColor = IIf(Me.MFG.Row = 0, vbActiveBorder, altFarbe(MfGTyp))
'  If MFG.CellBackColor = 0 Then Stop
  fgespei(MfGTyp) = 0
 End If
End Sub ' mfg_leavecell()

Private Sub mfg_gotfocus()
 aCtl = MFG.name
 If obdebug Then Debug.Print "mfg_gotfocus("
 Call MFG_Entercell
 Me.Cb1.Visible = False
 Me.Tb1.Visible = False
End Sub ' mfg_gotfocus()

Public Sub machAusbez(Optional Einschr$, Optional EinsNm$)
 If obdebug Then Debug.Print "machAusbez("
 Me.Überschrift = EinsNm
 Call MFGRefresh(azgab)
End Sub ' machAusbez(Optional Einschr$, Optional EinsNm$)

Public Sub machWochenplan(Optional Einschr$, Optional EinsNm$)
 If obdebug Then Debug.Print "machWochenplan("
 Me.Überschrift = EinsNm
' With Me.MFG
  Call MFGRefresh(azgwp)
'  nichtweiter = True
'  Me.ucMDIKeys.SetFocus
' End With
End Sub ' machWochenplan(

Private Sub userlöschen_Click()
 Dim altcapt$, altUser$, altPwd$
 If prüfeUser Then
  altUser = frmL.txtUserName
  altPwd = frmL.txtPassword
  frmL.txtUserName = vNS
  frmL.txtPassword = vNS
  frmL.obdefinier = True
  altcapt = frmL.Caption
  frmL.Caption = "Zu löschender User"
  frmL.txtPassword.Visible = False
  frmL.Show 1
  If LenB(frmL.txtUserName) <> 0 Then
   Call MDI.dbv.wCn.Execute("DELETE FROM `" & tbm(tbul) & "` WHERE user = '" & frmL.txtUserName & "'", rAf)
  End If
  frmL.txtPassword.Visible = True
  frmL.txtPassword = altPwd
  frmL.txtUserName = altUser
  frmL.Caption = altcapt
  frmL.obdefinier = False
 End If ' prüfeUser
End Sub ' userlöschen_Click()

Private Sub userhinzufügen_Click()
 Dim rszahl As New ADODB.Recordset, obok%
 Dim altcapt$, altUser$, altPwd$
 rszahl.Open "SELECT COUNT(0) as ct FROM `" & tbm(tbul) & "` WHERE user <> ''", MDI.dbv.wCn, adOpenStatic, adLockReadOnly
 If rszahl!ct = 0 Then
  obok = True
 Else
  If prüfeUser Then obok = True
 End If
 If obok Then
  altUser = frmL.txtUserName
  altPwd = frmL.txtPassword
  frmL.txtUserName = vNS
  frmL.txtPassword = vNS
  frmL.obdefinier = True
  altcapt = frmL.Caption
  frmL.Caption = "Hinzuzufügender User"
  frmL.Show 1
  If LenB(frmL.txtUserName) <> 0 And LenB(frmL.txtPassword) <> 0 Then
   Call MDI.dbv.wCn.Execute("INSERT INTO `" & tbm(tbul) & "`(user,Passwort,hinzugefügt,geändert) VALUES('" & frmL.txtUserName & "',aes_encrypt('" & frmL.txtPassword & "','0&F54'),NOW(),0)", rAf)
  End If
  frmL.Caption = altcapt
  frmL.txtPassword = altPwd
  frmL.txtUserName = altUser
  frmL.obdefinier = False
 End If
End Sub ' userhinzufügen_Click

Private Sub userliste_Click()
 Call MFGRefresh(azgul)
End Sub ' userliste_Click()

Private Sub zeigarten_Click()
 If obdebug Then Debug.Print "zeigarten_Click("
 Call schließen
 Call MFGRefresh(azgar)
' nichtweiter = True
' Me.ucMDIKeys.SetFocus
End Sub ' zeigarten_Click()

Sub cellweiter() ' zur Zeit nicht verwendet
 If obdebug Then Debug.Print "cellweiter("
 With Me.MFG
  .CellBackColor = altFarbe(MfGTyp)
  If .Col = .Cols - 1 Then
   If .Row = .Rows - 1 Then
    .Row = 1
   Else
    .Row = .Row + 1
   End If
   .Col = 1
  Else
   .Col = .Col + 1
  End If
 End With
 Call MFG_Entercell
End Sub ' cellweiter()

Sub einsweiter(ri As Richtung)
 Dim altTop&
 With Me.MFG
  .CellBackColor = altFarbe(MfGTyp)
  fgespei(MfGTyp) = 0
  If ri = Rec Then
   If .Col < .Cols - 1 Then .Col = .Col + 1
  ElseIf ri = Lin Then
   If .Col Then .Col = .Col - 1
  ElseIf ri = obe Then
   If .Row = .TopRow Then
    If .TopRow > .FixedRows Then
     .TopRow = .TopRow - 1
     .Row = .Row - 1
    Else
     .Row = .FixedRows - 1
    End If
   Else
    If .Row Then .Row = .Row - 1
   End If
  ElseIf ri = unt Then
   If .Row = .FixedRows - 1 Then
    .Row = .Row + (.TopRow - .FixedRows + 1)
   Else
    If .Row < .Rows - 1 Then .Row = .Row + 1
    If .Row > ZeiZa + .TopRow - .FixedRows Then
     .TopRow = .TopRow + 1
    End If
   End If
  ElseIf ri = gre Then
   .Col = .Cols - 1
  ElseIf ri = gli Then
   .Col = 1
  ElseIf ri = gob Then
   .Row = 1
  ElseIf ri = gun Then
   .Row = .Rows - 1
  ElseIf ri = stno Then
   altTop = .TopRow
   .TopRow = MAX(.TopRow - 30, .FixedRows)
   If .Row >= .FixedRows Then
    If .TopRow <> altTop Then
     .Row = .Row + .TopRow - altTop
    Else
     .Row = .FixedRows
    End If
   End If
  ElseIf ri = stnu Then
   altTop = .TopRow
   .TopRow = .TopRow + 30
   If .Row >= .FixedRows Then
    If .TopRow <> altTop Then
     .Row = .Row + .TopRow - altTop
    Else
     .Row = .Rows - 1
    End If
   End If
  End If
 End With
 Call MFG_Entercell
End Sub ' einsweiter

Private Sub MFG_LostFocus()
 If obdebug Then Debug.Print "mfg_lostfocus(", "Row:", Me.MFG.Row, "Col:", Me.MFG.Col
' Stop
 If NoLostFocus = 0 Then
  Select Case Me.ActiveControl.name
   Case "Tb1", "Cb1", "Command1", "Command2"
   Case Else
    Me.ucMDIKeys.SetFocus
'    If nichtweiter Then
'     nichtweiter = 0
'    Else
''     Call cellweiter
'    End If
  End Select
 End If
End Sub ' MFG_LostFocus

Private Sub zeigmitarbeiter_Click()
 If obdebug Then Debug.Print "zeigmitarbeiter_Click("
 Call schließen
 Call MFGRefresh(azgma)
' nichtweiter = True
' Me.ucMDIKeys.SetFocus
End Sub 'zeigmitarbeiter_Click()

Private Sub zeigdienstplan_Click()
 If obdebug Then Debug.Print "zeigdienstplan_Click("
 Call schließen
 Call MFGRefresh(azgdp)
' nichtweiter = True
' Me.ucMDIKeys.SetFocus
End Sub ' zeigdienstplan_Click

Sub schließen()
 Select Case ActiveControl.name
  Case Cb1.name, Tb1.name
   ActiveControl.Visible = False
   verwerfen = True
 End Select
End Sub ' schließen()

Private Sub merken(MfGTyp As azgtyp)
 If obdebug Then Debug.Print "merk(", "Row:", Me.MFG.Row, "Col:", Me.MFG.Col
 merkRow(MfGTyp) = Me.MFG.Row
 merkCol(MfGTyp) = Me.MFG.Col
 merkTop(MfGTyp) = Me.MFG.TopRow
 merkLeft(MfGTyp) = Me.MFG.LeftCol
End Sub ' merk()

Private Sub merken2(MfGTyp As azgtyp)
 If obdebug Then Debug.Print "merk(", "Row:", Me.MFG.Row, "Col:", Me.MFG.Col
 merkRow(MfGTyp) = vormerkRow(MfGTyp)
 merkCol(MfGTyp) = vormerkCol(MfGTyp)
 merkTop(MfGTyp) = vormerkTop(MfGTyp)
 merkLeft(MfGTyp) = vormerkLeft(MfGTyp)
End Sub ' merk()

Private Sub vormerken(MfGTyp As azgtyp)
' If obdebug Then Debug.Print "vormerk(", "Row:", Me.MFG.Row, "Col:", Me.MFG.Col
 vormerkRow(MfGTyp) = Me.MFG.Row
 vormerkCol(MfGTyp) = Me.MFG.Col
 vormerkTop(MfGTyp) = Me.MFG.TopRow
 vormerkLeft(MfGTyp) = Me.MFG.LeftCol
End Sub ' vormerken()

' aufgerufen in: key, MFGRefresh
Private Sub SpuckAus()
 Dim i&
 Static obSpuck%
 On Error GoTo fehler
 If obdebug Then Debug.Print "SpuckAus("
 On Error Resume Next
 If Not obSpuck Then
  obSpuck = True
  If merkCol(MfGTyp) <> 0 Then Me.MFG.Col = merkCol(MfGTyp) Else Me.MFG.Col = obSp1Fest(azt(MfGTyp)) ' Else Me.MFG.Col = 1
  Me.MFG.ColSel = Me.MFG.Col
  If merkRow(MfGTyp) <> 0 Then
'  Me.MFG.Row = merkRow(MfgTyp)
#If True Then
'  MFG.Col = merkCol(MfgTyp)
   MFG.Row = merkRow(MfGTyp)
   MFG.TopRow = merkTop(MfGTyp) 'max(merkRow(MfgTyp) - 10, 0)
   MFG.LeftCol = merkLeft(MfGTyp)
   MFG.RowSel = MFG.Row
  Else
   Me.MFG.TopRow = Me.MFG.Rows - 1
   Me.MFG.LeftCol = 0
   Me.MFG.Row = Me.MFG.Rows - 1
#Else
   Me.MFG.Visible = True
   Dim ZeilenpSeite%, vorRow&
   ZeilenpSeite = 54
   noenter = True
   If merkRow(MfGTyp) > MFG.Rows / 2 Then
    Call StDirekt("{PGDN}", vbCtrlMask, obweiter:=True)
   Else
    Call StDirekt("{PGUP}", vbCtrlMask, obweiter:=True)
   End If
   If merkRow(MfGTyp) > MFG.Rows - 1 Then
    merkRow(MfGTyp) = MFG.Rows - 1
   End If
   Do While MFG.Row <> merkRow(MfGTyp)
    If MFG.Row > merkRow(MfGTyp) Then
     If MFG.Row - merkRow(MfGTyp) > ZeilenpSeite / 2 Then
      vorRow = MFG.Row
      Call StDirekt("{PGUP}", obvorher:=True, obweiter:=True)
      ZeilenpSeite = vorRow - MFG.Row
     Else
      Call StDirekt("{UP}", obvorher:=True, obweiter:=True)
     End If
    Else
     If merkRow(MfGTyp) - MFG.Row > ZeilenpSeite / 2 Then
      vorRow = MFG.Row
      Call StDirekt("{PGDN}", obvorher:=True, obweiter:=True)
      ZeilenpSeite = MFG.Row - vorRow
     Else
      Call StDirekt("{DOWN}", obvorher:=True, obweiter:=True)
     End If
    End If
   Loop
   For i = MFG.FixedCols + 2 To merkCol(MfGTyp)
    Call StDirekt("{RIGHT}", obvorher:=True, obweiter:=True)
   Next i
   noenter = False
   Me.ucMDIKeys.SetFocus
  Else
   Me.MFG.Visible = True
   Call StDirekt("{PGDN}", vbCtrlMask)
#End If
'  Me.MFG.Row = 1 ' Else If Me.MFG.Rows > 2 Then Me.MFG.Row = Me.MFG.Rows - 2 Else Me.MFG.Row = Me.MFG.Rows - 1
  End If
  On Error GoTo fehler
  Call MFG_Entercell
  obSpuck = False
 End If ' obSpuck
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in SpuckAus/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' SpuckAus()

' aufgerufen in MDIForm_Load
Sub TabAnalyse()
 Dim CtTbl$, sql$
 Dim rsh As New ADODB.Recordset
 Dim DBSp%, j%, Spalte&, res%, k As TbTyp
 On Error GoTo fehler
anfang:
 ReDim SpvDBSp(MTBeg To tbende - 1, 0)
 ReDim RefTab(MTBeg To tbende - 1, 0)
 For k = MTBeg To tbende - 1
  Set rsc = Nothing
  On Error Resume Next
  Err.Clear
  sql = "SHOW CREATE TABLE `" & tbm(k) & "`"
  Call rsc.Open(sql, dbv.wCn, adOpenStatic, adLockReadOnly)
  If Err.Number <> 0 Then
   MsgBox "Fehler: " & Err.Description & vbCrLf & "bei: " & sql & vbCrLf & "über: " & dbv.Constr & vbCrLf & "Bitte Verbindung verbessern!"
   Call wähleBenutzer
   Call dbv.Auswahl("dp", tbm(tbwp), tbm(tbdp)) '("", "--multi", "")
   GoTo anfang
  End If
  On Error GoTo fehler
  CtTbl = rsc.Fields(1)
  Set rsc = Nothing
  rsc.Open "SHOW INDEX FROM `" & tbm(k) & "` WHERE key_name = 'PRIMARY'", dbv.wCn, adOpenDynamic, adLockReadOnly
  j = 0
  Do While Not rsc.EOF
   PrimI(k, j) = rsc!column_name
   j = j + 1
   rsc.Move 1
  Loop
  Set rsc = Nothing
  Call rsc.Open("SHOW FULL COLUMNS FROM `" & tbm(k) & "`", dbv.wCn, adOpenStatic, adLockReadOnly)
  Set rs = Nothing
  rs.Open "SELECT * FROM `" & tbm(k) & "` LIMIT 1", dbv.wCn, adOpenDynamic, adLockReadOnly
' catx.ActiveConnection = Cn
  Dim obakt%
  If k = MTBeg Then
   obakt = -1
  ElseIf rs.Fields.Count - 1 > UBound(SpvDBSp, 2) Then
   obakt = -1
  Else
   obakt = 0
  End If
  If obakt Then
   ReDim Preserve SpvDBSp(MTBeg To tbende - 1, rs.Fields.Count)
   ReDim Preserve DBSpvSp(MTBeg To tbende - 1, rs.Fields.Count)
   ReDim Preserve SpNm(MTBeg To tbende - 1, rs.Fields.Count)
   ReDim Preserve SpCm(MTBeg To tbende - 1, rs.Fields.Count)
   ReDim Preserve FdTyp(MTBeg To tbende - 1, rs.Fields.Count)
   Dim obRefAkt%
   obRefAkt = True
   If k > tbdp Then If rs.Fields.Count - 1 <= UBound(RefTab, 2) Then obRefAkt = False
   If obRefAkt Then
    ReDim Preserve RefTab(MTBeg To tbende - 1, rs.Fields.Count)
    ReDim Preserve RefSp(MTBeg To tbende - 1, rs.Fields.Count)
    ReDim Preserve SpvDBSp(MTBeg To tbende - 1, rs.Fields.Count)
    ReDim Preserve SpNm(MTBeg To tbende - 1, rs.Fields.Count)
    ReDim Preserve SpCm(MTBeg To tbende - 1, rs.Fields.Count)
    ReDim Preserve DBSpvSp(MTBeg To tbende - 1, rs.Fields.Count)
    ReDim Preserve FdTyp(MTBeg To tbende - 1, rs.Fields.Count)
   End If
  End If
  If k = tbdp Then
   Call rsh.Open("SELECT COUNT(*) ct FROM `" & tbm(tbma) & "`", dbv.wCn, adOpenStatic, adLockReadOnly)
   If rsh!ct + 1 > UBound(RefTab, 2) Then
    ReDim Preserve RefTab(MTBeg To tbende - 1, rsh!ct + 1)
    ReDim Preserve RefSp(MTBeg To tbende - 1, rsh!ct + 1)
    ReDim Preserve SpvDBSp(MTBeg To tbende - 1, rsh!ct + 1)
    ReDim Preserve DBSpvSp(MTBeg To tbende - 1, rsh!ct + 1)
    ReDim Preserve FdTyp(MTBeg To tbende - 1, rsh!ct + 1)
   End If
   For DBSp = 1 To UBound(RefTab, 2)
    RefTab(tbdp, DBSp) = tbm(tbar)
    RefSp(tbdp, DBSp) = "ArtNr"
    SpvDBSp(tbdp, DBSp) = DBSp
    DBSpvSp(tbdp, DBSp) = DBSp
    FdTyp(k, DBSp) = adChar
   Next DBSp
   obSp1Fest(k) = True
  Else ' k = tbdp Then
'  With Me.MFG
    Spalte = 0
    SpZ(k) = 1
    Do While Not rsc.EOF
     SpZ(k) = SpZ(k) + 1
     If rsc!extra = "auto_increment" Then
      obAuto(k) = True
      DBSp = 0
      SpZ(k) = SpZ(k) - 1
      res = 1
     Else
      DBSp = Spalte + 1 - res
     End If
     SpvDBSp(k, Spalte) = DBSp
     DBSpvSp(k, DBSp) = Spalte
     SpNm(k, Spalte) = rsc!Field
     Dim ConstrPos%, RefPos%, KlaPos%, TeilS$
     ConstrPos = InStr(CtTbl, "FOREIGN KEY (`" & rsc!Field & "`)")
     If ConstrPos <> 0 Then
      RefPos = InStr(ConstrPos, CtTbl, "REFERENCES")
      TeilS = Mid(CtTbl, RefPos + 12)
      RefTab(k, DBSp) = Left(TeilS, InStr(TeilS, "`") - 1)
      RefPos = InStr(RefPos, CtTbl, "(")
      TeilS = Mid(CtTbl, RefPos + 2)
      RefSp(k, DBSp) = Left(TeilS, InStr(TeilS, "`") - 1)
     End If
     SpCm(k, DBSp) = IIf(LenB(rsc!comment) = 0, rsc.Fields(0), rsc!comment)
     rsc.Move 1
     FdTyp(k, Spalte) = TypCast(rs.Fields(DBSp).Type)
     Spalte = Spalte + 1
    Loop
    If rs.Fields(0).Type = adDBDate Or rs.Fields(0).Type = adDBTime Or obAuto(k) Then
     obSp1Fest(k) = True
    End If
'   End With
  End If
 Next k
 On Error GoTo fehler
 dbv.wCn.Execute (doppelteWeg)
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in TabAnalyse/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' TabAnalyse

Function TypCast(divTyp As DataTypeEnum) As DataTypeEnum
     Select Case divTyp
       Case adDBTime: TypCast = adDBTime
       Case adDate
       Case adDBDate, adDBTime, adDBTimeStamp
        TypCast = adDBDate
       Case adChar, adLongVarChar, adLongVarWChar, adVarChar, adVariant, adVarWChar, adWChar
        TypCast = adChar
       Case adBigInt, adCurrency, adDecimal, adDouble, adInteger, adNumeric, adSingle, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
        TypCast = adNumeric
       Case adBoolean
        TypCast = adBoolean
       Case adArray, adBinary, adBSTR, adEmpty, adError, adGUID, adIDispatch, adIUnknown, adLongVarBinary, adUserDefined, adVarBinary ' , adVector + adByRef gibts wohl nicht
        TypCast = adIUnknown
     End Select
End Function ' TypCast(divTyp As DataTypeEnum) As DataTypeEnum

Sub MFGRefresh(neuMFGTyp As azgtyp, Optional aktMSH As MSHFlexGrid, Optional nichtAusSpucken%)
 Dim i%, obFarbe%
 On Error GoTo fehler
 If obdebug Then Debug.Print "MFGRefresh(", azm(neuMFGTyp), neuMFGTyp
 If aktMSH Is Nothing Then Set aktMSH = Me.MFG
 If MfGTyp <> azgnix And aktMSH.name = Me.MFG.name Then
  If wiederholt Then
   Call merken(MfGTyp)
  Else
   wiederholt = True
  End If
 End If
 fgespei(neuMFGTyp) = 0
 Me.Tb1.Visible = False
 Me.Cb1.Visible = False
 Me.Command1.Visible = False
 Me.Command2.Visible = False
 aktMSH.Visible = False
 Me.MSH2.Visible = False
 If neuMFGTyp = azgdp Then
  If obMSH2 Then
   Call MFGRefresh(azgar, Me.MSH2, nichtAusSpucken:=True)
   Me.MSH2.Height = (Me.MSH2.CellHeight + 20) * (Me.MSH2.Rows)
   Me.MSH2.Width = 0
   For i = 0 To Me.MSH2.Cols - 1
    Me.MSH2.Col = i
    Me.MSH2.Width = (Me.MSH2.Width + 30) + Me.MSH2.CellWidth
   Next i
   Me.MSH2.Col = 1
   Me.MSH2.Left = Me.Left + Me.Width - Me.MSH2.Width - 400
   Me.MSH2.Enabled = False
   Me.MSH2.Visible = True
  End If
 End If
 If neuMFGTyp = azgdp Or neuMFGTyp = azgwp Or neuMFGTyp = azgab Or nichtAusSpucken Then
  aktMSH.Top = 240
  Me.MSH2.Top = aktMSH.Top
  Me.Rechts.Visible = True
  Me.Links.Visible = True
  Select Case neuMFGTyp
   Case azgdp
    Me.Überschrift.Visible = False
    Me.Jahr.Visible = True
   Case azgwp, azgab
    Me.Überschrift.Visible = True
    Me.Links.Visible = False
    Me.Rechts.Visible = False
    Me.Jahr.Visible = False
  End Select ' case neuMFGTyp
 Else
  aktMSH.Top = 0
  Me.Rechts.Visible = False
  Me.Links.Visible = False
 End If ' neuMFGTyp = azgdp Or neuMFGTyp = azgwp Or nichtAusSpucken  else
 Me.Tabl = tbm(azt(neuMFGTyp))
 Screen.MousePointer = vbHourglass
' NoLostFocus = True
' DoEvents
' NoLostFocus = False
 With aktMSH
snochmal:
  .Clear
  .Rows = 2
  i = 0
  Set rs = Nothing
  If neuMFGTyp = azgpr Then
   sql = "SELECT id, tag, kuerzel, artnrv vorher, artnr nachher, DATE_FORMAT(aenddat,'%d.%m.%Y %T') 'Änderung', aendpc 'PC', aenduser '-benutzer',user FROM `" & tbm(tbpr) & "` LEFT JOIN `" & tbm(tbma) & "` on `" & tbm(tbpr) & "`.`persnr` = `" & tbm(tbma) & "`.`persnr` ORDER BY `aenddat` DESC"
  ElseIf neuMFGTyp = azgul Then
   sql = "SELECT user, DATE_FORMAT(hinzugefügt,'%d.%m.%Y %T') as 'hinzugefügt', id FROM `" & tbm(tbul) & "` ORDER BY `user`"
  Else
   sql = "SELECT * FROM `" & tbm(azt(neuMFGTyp)) & "`" & IIf(LenB(Einschr(neuMFGTyp)) = 0, vNS, Einschr(neuMFGTyp)) & IIf(neuMFGTyp = azgar, " WHERE zusatz = 0", vNS)
  End If
'  rs.Open sql, dbv.wCn, adOpenDynamic, adLockReadOnly
  myFrag rs, sql, adOpenStatic, dbv.wCn, adLockReadOnly
  If neuMFGTyp = azgpr Or neuMFGTyp = azgul Then
   If rs.Fields.Count > UBound(SpCm, 2) Then
    ReDim Preserve SpCm(azt(neuMFGTyp), rs.Fields.Count)
    ReDim Preserve FdTyp(azt(neuMFGTyp), rs.Fields.Count)
   End If
   For i = 0 To rs.Fields.Count - 1
    SpCm(azt(neuMFGTyp), i) = rs.Fields(i).name
    FdTyp(azt(neuMFGTyp), i) = TypCast(rs.Fields(i).Type)
   Next i
  End If
  Dim rsfcn As New ADODB.Connection
  rsfcn.Open (DBCnS)
  Select Case neuMFGTyp
   Case azgwp, azgdp, azgpr, azgul
    Set rsf = Nothing
'    rsf.Open "SELECT artnr, farbe FROM `" & tbm(tbar) & "` WHERE zusatz = 0", dbv.wCn, adOpenDynamic, adLockReadOnly '  desc
    myFrag rsf, "SELECT artnr, farbe FROM `" & tbm(tbar) & "` WHERE zusatz = 0", adOpenDynamic, rsfcn, adLockReadOnly
  End Select
  Select Case neuMFGTyp
   Case azgar, azgma, azgwp, azgpr, azgul, azgab
    .Cols = SpZ(azt(neuMFGTyp))
    If obSp1Fest(azt(neuMFGTyp)) = 0 Then .Cols = .Cols - 1
    For i = 0 To .Cols - 1
     .TextMatrix(0, i) = SpCm(azt(neuMFGTyp), i)
    Next i
    Do While Not rs.EOF
     .Row = .Rows - 1
     For i = 0 To rs.Fields.Count - 1
      .Col = SpvDBSp(azt(neuMFGTyp), i)
      If mitVGF Then .CellForeColor = VGF1
      Select Case FdTyp(azt(neuMFGTyp), i) ' rs.Fields(i).Type '
       Case adDate
       Case adDBTime
        .Text = Format(rs.Fields(i), "dd.mm.yyyy hh:mm:ss")
        .CellAlignment = flexAlignCenterCenter
       Case adDBDate
        .Text = Format(rs.Fields(i), "dd.mm.yyyy")
        .CellAlignment = flexAlignCenterCenter
       Case adChar
        .Text = IIf(IsNull(rs.Fields(i)), vNS, rs.Fields(i))
        .CellAlignment = flexAlignLeftCenter
       Case adNumeric
         .Text = IIf(IsNull(rs.Fields(i)), vNS, rs.Fields(i))
        .CellAlignment = flexAlignRightCenter
       Case adBoolean
        .Text = IIf(IsNull(rs.Fields(i)), IIf(rs.Fields(i), "x", "-"), rs.Fields(i))
        .CellAlignment = flexAlignCenterBottom
       Case adIUnknown
        .Text = IIf(IsNull(rs.Fields(i)), vNS, rs.Fields(i))
        .CellAlignment = flexAlignRightBottom
      End Select ' Case FdTyp(azt(neuMFGTyp), i) ' rs.Fields(i).Type '
      Select Case neuMFGTyp
       Case azgwp, azgpr
        Call Einfärben(mitaltFar:=False)
       Case azgar
        If aktMSH.name = MFG.name Then
         .CellBackColor = rs!Farbe
         If mitVGF Then .CellForeColor = VGF1
        Else
         .CellBackColor = Heller(rs!Farbe, 2)
         If mitVGF Then
          .CellForeColor = Heller(.CellForeColor, 2)
         Else
          .CellForeColor = 6250335 ' Heller(.CellForeColor, 2)
         End If
        End If
'        If MFG.CellBackColor = 0 Then Stop
      End Select ' Case neuMFGTyp
     Next i
     rs.Move 1
     .Rows = .Rows + 1
    Loop
   Case azgdp
    Dim aktD As Date, ab() As Date, aus() As Date, nText$, j%, k%, sql0$
    Dim ftp% ' Feiertagszeiger
    ReDim pn(0)
    ReDim kue(0)
    Dim rsh As ADODB.Recordset, rs0 As ADODB.Recordset, rs1 As ADODB.Recordset, rs2 As ADODB.Recordset
    Do
     BegD = CDate("1.1." & Me.Jahr)
     EndD = CDate(Day(BegD) & "." & Month(BegD) & "." & Year(BegD) + 1)
'     Call FTbeleg(Me.Jahr)
'     sql0 = "SELECT * FROM `" & tbm(tbwp) & "` LEFT JOIN `" & tbm(tbma) & "` on `" & tbm(tbwp) & "`.PersNr = `" & tbm(tbma) & "`.PersNr WHERE not `" & tbm(tbma) & "`.Ein < " & datform(BegD) & " AND ab <= " & datform(EndD) & " AND NOT EXISTS (SELECT * FROM `" & tbm(tbwp) & "` as azn WHERE persnr = `" & tbm(tbwp) & "`.persnr AND ab > `" & tbm(tbwp) & "`.ab AND ab < " & datform(BegD) & ") ORDER BY `" & tbm(tbwp) & "`.`persnr`,`ab`"
     .Cols = 1
     Set rsh = Nothing
'    rsh.Open "SELECT * FROM `" & tbm(tbma) & "` WHERE (isnull(aus) OR (not isnull(aus) AND aus >= " & datform(BegD) & ")) AND (isnull(ein) OR (not isnull(ein) AND ein <= " & datform(EndD) & ")) ORDER BY `persnr`", Cn, adOpenDynamic, adLockReadOnly
     sql0 = "SELECT `" & tbm(tbwp) & "`.*,`" & tbm(tbma) & "`.Kuerzel,`" & tbm(tbma) & "`.Nachname,`" & tbm(tbma) & "`.Vorname,`" & tbm(tbma) & "`.Aus " & _
            "FROM `" & tbm(tbwp) & "` LEFT JOIN `" & tbm(tbma) & "` ON `" & tbm(tbwp) & "`.PersNr = `" & tbm(tbma) & "`.PersNr " & _
            "WHERE (`" & tbm(tbma) & "`.Aus > " & datform(BegD) & " OR isnull(`" & tbm(tbma) & "`.Aus) OR (`" & tbm(tbma) & "`.Aus = " & datform(CDate(0)) & " OR `" & tbm(tbma) & "`.Aus = " & "'0000-00-00'" & ")) AND ab <= " & datform(EndD) & " " & _
            "AND NOT EXISTS (SELECT * FROM `" & tbm(tbwp) & "` azn WHERE persnr = `" & tbm(tbwp) & "`.persnr AND ab > `" & tbm(tbwp) & "`.ab AND ab <= " & datform(BegD) & ") " & _
            "ORDER BY `" & tbm(tbwp) & "`.`persnr`,`ab`"
     sql = "SELECT persnr,kuerzel,nachname,vorname,min(ab) ab,aus FROM (" & sql0 & ") i GROUP BY kuerzel ORDER BY persnr"
'     rsh.Open sql, dbv.wCn, adOpenDynamic, adLockOptimistic
     On Error GoTo EigFehler
     Set rsh = Nothing
     myFrag rsh, sql, adOpenStatic, dbv.wCn, adLockReadOnly
     On Error GoTo fehler
     If rsh.BOF Then
      If Me.Jahr < Year(Now) Then
       Me.Jahr = Me.Jahr + 1
      Else
       If Me.Jahr > Year(Now) Then
        Me.Jahr = Me.Jahr - 1
       Else
        MsgBox "Bitte erst Mitarbeiterzeiträume eingeben!"
        Call zeigmitarbeiter_Click
        Exit Sub
       End If
      End If
     Else
      Exit Do
     End If
    Loop
    Call FTbeleg(Me.Jahr)
    Do While Not rsh.EOF
     .Cols = .Cols + 1
     .Col = .Cols - 1
     ReDim Preserve pn(.Col - 1)
     ReDim Preserve ab(.Col - 1)
     ReDim Preserve aus(.Col - 1)
     ReDim Preserve kue(.Col - 1)
     pn(.Col - 1) = rsh!Persnr
     kue(.Col - 1) = rsh!Kuerzel
     If IsNull(rsh!ab) Then
      ab(.Col - 1) = 0
     Else
      ab(.Col - 1) = rsh!ab
     End If
     aus(.Col - 1) = IIf(IsNull(rsh!aus), CDate("31.12.9999"), rsh!aus)
     .Row = 0
     .Text = IIf(IsNull(rsh!Nachname), """""", rsh!Nachname)
     rsh.Move 1
    Loop ' While Not rsh.EOF
    Set rsh = Nothing
    .Rows = SZZ + 2
    For i = 0 To .Cols - 2
     If i <= UBound(pn) Then
     .Col = i + 1
     Dim UAB!, UAA!, FBA! ' Urlaubsanspruch bisher, Urlaubsanspruch aktuell, Fortbildung aktuell
'     Call UrlAnsp(pn(i), CDate("1.1." & Me.Jahr), CDate("1.1." & Me.Jahr + 1), dbv.wCn, UAA!, UAB!)
     Call UrlAnspr(pn(i), Me.Jahr, dbv.wCn, UAA!, UAB!)
     Set rs0 = Nothing
'     rs0.Open "SELECT urlaub, Überstunden uest, fortbildung FROM `" & tbm(tbbi) & "` WHERE persnr = " & pn(i) & " AND jahr = " & Me.Jahr, dbv.wCn, adOpenStatic, adLockReadOnly
     myFrag rs0, "SELECT urlaub, Überstunden uest, fortbildung FROM `" & tbm(tbbi) & "` WHERE persnr = " & pn(i) & " AND jahr = " & Me.Jahr, adOpenStatic, dbv.wCn, adLockReadOnly
     Set rs1 = Nothing
'     rs1.Open "SELECT urlaub, Überstunden uest, fortbildung FROM `" & tbm(tbbi) & "` WHERE persnr = " & pn(i) & " AND jahr = " & Me.Jahr - 1, dbv.wCn, adOpenStatic, adLockReadOnly
     myFrag rs1, "SELECT urlaub, Überstunden uest, fortbildung FROM `" & tbm(tbbi) & "` WHERE persnr = " & pn(i) & " AND jahr = " & Me.Jahr - 1, adOpenStatic, dbv.wCn, adLockReadOnly
     .Row = 1 ' Überstunden Vorjahr
     If rs1.BOF Then .Text = 0 Else .Text = rs1!uest
     .Text = Round(.Text, 1)
     .Row = 2 ' Überstunden gesamt
     .Text = 0
     If rs0.State <> 0 Then If Not rs0.BOF Then .Text = rs0!uest
'     If Not rs1.BOF Then .Text = .Text - rs1!uest
     .Text = Round(.Text, 1)
     .Row = 3 ' Urlaub Vorjahr
     If rs1.BOF Then .Text = 0 Else .Text = -rs1!urlaub
     .Text = Round(.Text, 1)
     Dim uvj!
     uvj = .Text
     .Row = 5 ' Urlaub gesamt
     .Text = 0
     If rs0.State <> 0 Then If Not rs0.BOF Then .Text = rs0!urlaub
     .Text = -Round(.Text, 1) ' 6.3.23
     Dim uges!
     uges = -.Text ' 6.3.23
     .Row = 4 ' Url'eintr. heuer
'     .Text = uges - uvj
     Set rs2 = Nothing
     sql = "SELECT COUNT(0) Zahl FROM `" & tbm(tbdp) & "` WHERE persnr = " & pn(i) & " AND YEAR(tag) = " & Me.Jahr & " AND ArtNr IN ('u','uw')"
     myFrag rs2, sql, adOpenStatic, dbv.wCn, adLockReadOnly
     If rs2.State <> 0 Then If Not rs2.BOF Then .Text = rs2!zahl
'     If Not rs1.BOF Then .Text = .Text + rs1!urlaub
'     .Row = 2
'     If Not rs0.BOF Then If Not IsNull(rs0!uest) Then .Text = rs0!uest
'     .Row = 5
'     If Not rs0.BOF Then If Not IsNull(rs0!Fortbildung) Then FBA = rs0!Fortbildung
'     .Text = Round(FBA, 1)
'     .Row = 4
'     If Not rs0.BOF Then If Not IsNull(rs0!urlaub) Then UAA = UAA - rs0!urlaub
'     .Text = Round(UAA, 1)
'     Set rs0 = Nothing
'     rs0.Open "SELECT sum(urlaub) as urlaub, sum(überstunden) uestsum FROM `" & tbm(tbbi) & "` WHERE persnr = " & pn(i) & " AND jahr < " & Me.Jahr, dbv.wCn, adOpenStatic, adLockReadOnly
'     .Row = 1
'     If Not IsNull(rs0!uestsum) Then .Text = rs0!uestsum
'     .Row = 3
'     If Not IsNull(rs0!urlaub) Then UAB = UAB - rs0!urlaub
'     .Text = Round(UAB, 1)
     End If ' i <= ubound(pn)
    Next i
    Set rs0 = Nothing
    Dim wp()
'    Einschr(tbdp) = " LEFT JOIN `" & tbm(tbar) & "` ON `" & tbm(tbwp) & "`.artnr = `" & tbm(tbar) & "`.artnr"
'    sql = "SELECT * FROM `" & tbm(tbwp) & "`" & " ORDER BY `persnr`, `ab`"
    Set rs0 = New ADODB.Recordset
    Dim raff&
    On Error GoTo EigFehler
    rs0.Open sql0, dbv.wCn, adOpenDynamic, adLockReadOnly '  desc
    On Error GoTo fehler
'    myFrag rs0, sql0, adOpenStatic, dbv.wCn, adLockReadOnly
    ReDim wp(rs0.Fields.Count - 1, 1, 0)
    Dim Zeiger&
    Zeiger = 0
    Do While Not rs0.EOF
     For i = 0 To rs0.Fields.Count - 1
      If AusFeldNr = 0 Then If rs0.Fields(i).name = "Aus" Then AusFeldNr = i
      wp(i, 0, Zeiger) = rs0.Fields(i)
'      If rsf.State = 0 Then
'       Set rsf = Nothing
'       myFrag rsf, "SELECT artnr, farbe FROM `" & tbm(tbar) & "` WHERE zusatz = 0", adOpenDynamic, dbv.wCn, adLockOptimistic
'      End If
      If rsf.State = 0 Then
         GoTo snochmal
      End If
      Call rsf.Find("artnr = '" & rs0.Fields(i) & "'", , adSearchForward, 1)
      If Not rsf.EOF Then
       wp(i, 1, Zeiger) = rsf!Farbe
      End If
     Next i
     rs0.Move 1
'     If Not rs0.EOF Then
'      If rs0!ab < BegD Then ' alten Datensatz überschrieben
'      Else
        Zeiger = Zeiger + 1
        ReDim Preserve wp(rs0.Fields.Count - 1, 1, Zeiger)
'      End If
'     End If
    Loop ' While Not rs0.EOF
    Dim FTFarbe&
    FTFarbe = 6250335 ' Dunkelgrau
    Call rsf.Find("artnr = 'WF'", , adSearchForward, 1)
    If Not rsf.EOF Then
     FTFarbe = rsf!Farbe
    End If
    ReDim Preserve wp(rs0.Fields.Count - 1, 1, UBound(wp, 3) - 1)
    .Rows = EndD - BegD + 2 + SZZ
    .Col = 0
    .Row = 1
    .Text = "Überstd.Vorjahr"
    .Row = 2
    .Text = "Überstd.gesamt"
    .Row = 3
    .Text = "Resturl.Vorjahr"
    .Row = 4
    .Text = "Url'eintr. heuer"
    .Row = 5
    .Text = "Resturl.gesamt"
    .Row = 6
    .Text = "Fortbildung"
    .Row = SZZ + 1
    .FixedRows = SZZ + 1
    rs0.MoveFirst
    ftp = 0
    For aktD = BegD To EndD - 1
     nText = Format(aktD, "ddd,d.m.yy")
     If aktD = ftag(ftp).Datum Then
       .Text = nText & " " & ftag(ftp).KuNa
       .CellBackColor = FTFarbe
       For j = ftp + 1 To UBound(ftag) ' 2008: Maifeiertag = Christi Himmelfahrt
        If ftag(j).Datum > ftag(ftp).Datum Then
         ftp = j
         Exit For
        End If
       Next j
     Else
      .Text = nText
     End If
     If aktD Mod 7 < 2 Then .CellBackColor = FTFarbe
     .Row = .Row + 1
    Next
    .Row = SZZ + 1
    Dim BegMD As Date, EndMD As Date, AltPN&, Wd%
    ' Wochenplan eintragen
    AltPN = -1
 '   aktMSH.Visible = True ' Zum Fehlerfinden
    For i = 0 To UBound(wp, 3) ' Mitarbeiterdefinitionszeiträume
'     If wp(0, 0, i) = 70 Then Stop
     If AltPN = -1 Then
      .Col = 1
     Else
      If wp(0, 0, i) <> AltPN Then
       If .Col < .Cols - 1 Then
        .Col = .Col + 1 ' alte Personalnummer, Mitarbeiterspalte im Grid
       End If
      End If
     End If
     If IsNull(wp(1, 0, i)) Then
      BegMD = 0
     Else
      BegMD = wp(1, 0, i)
     End If
     If BegMD < BegD Then BegMD = BegD
     ftp = 0
     Do While ftag(ftp).Datum < BegMD
      ftp = ftp + 1
      If ftp > UBound(ftag) Then Exit Do
     Loop
     EndMD = IIf(IsNull(wp(AusFeldNr, 0, i)), #12/31/9999#, wp(AusFeldNr, 0, i)) ' Austrittsdatum
     If i < UBound(wp, 3) Then If wp(0, 0, i) = wp(0, 0, i + 1) Then EndMD = wp(1, 0, i + 1) ' -1
     If EndMD > EndD Then EndMD = EndD
     If EndMD > aus(.Col - 1) Then EndMD = aus(.Col - 1)
     .Row = BegMD - BegD + SZZ + 1
     Wd = Weekday(BegMD)
     If Wd = 1 Then Wd = 8
     For aktD = BegMD To EndMD - 1
      If mitVGF Then .CellForeColor = VGF1
      If aktD = ftag(ftp).Datum Then
       If ftag(ftp).obhalb Then
        If Left(wp(Wd, 0, i), 1) = "a" Then
         .Text = "hFT"
         Call rsf.Find("artnr = '" & .Text & "'", , adSearchForward, 1)
         If Not rsf.EOF Then
          .CellBackColor = rsf!Farbe
         End If
        Else
        Dim test$
        If IsNull(wp(Wd, 0, i)) Then
         .Text = vNS
        Else
         .Text = wp(Wd, 0, i)
        End If
        .CellBackColor = wp(Wd, 1, i)
        End If
       Else
        If Not IsNull(wp(8, 0, i)) Then .Text = wp(8, 0, i) ' Feiertag -> Sonntag
        .CellBackColor = wp(8, 1, i)
       End If
       For j = ftp + 1 To UBound(ftag) ' 2008: Maifeiertag = Christi Himmelfahrt
        If ftag(j).Datum > ftag(ftp).Datum Then
         ftp = j
         Exit For
        End If
       Next j
      Else
       If IsNull(wp(Wd, 0, i)) Then
        .Text = "Null"
       Else
        .Text = wp(Wd, 0, i)
       End If
       .CellBackColor = wp(Wd, 1, i)
      End If
'      If MFG.CellBackColor = 0 Then Stop
      Wd = Wd + 1
      If Wd = 9 Then Wd = 2
      .Row = .Row + 1
     Next aktD
     AltPN = wp(0, 0, i)
    Next i
    For i = 0 To UBound(pn)
     .Col = i + 1
     Einschr(azgdp) = " LEFT JOIN `" & tbm(tbar) & "` ON `" & tbm(tbdp) & "`.artnr = `" & tbm(tbar) & "`.artnr"
     sql = "SELECT * FROM `" & tbm(tbdp) & "` " & Einschr(azgdp) & " WHERE `PersNr` = " & pn(i) & " AND `tag` >= " & datform(BegD) & " AND `tag` <= " & datform(EndD)
'     rsh.Open sql, dbv.wCn, adOpenDynamic, adLockReadOnly
     Set rsh = Nothing
     myFrag rsh, sql, adOpenStatic, dbv.wCn, adLockReadOnly
     Do While Not rsh.EOF
      .Row = rsh!Tag - BegD + SZZ + 1
      If Not IsNull(rsh!artnr) Then .Text = rsh!artnr
'      If rsh!Farbe = 33023 Then Stop
      If Not IsNull(rsh!Farbe) Then
       If mitVGF Then .CellForeColor = VGF2 ' 255 '13565951 ' Vordergrundfarbe2
       If mitfett Then
         .CellFontBold = True
         If .Text = "-" Then .CellFontSize = .CellFontSize + .CellFontSize '.Text = "---"
       End If
       .CellBackColor = rsh!Farbe ' .CellBackColor = Heller(rsh!Farbe, 1)
      End If
'      If MFG.CellBackColor = 0 Then Stop
      rsh.Move 1
     Loop ' while not rsh.eof
     Set rsh = Nothing
    Next ' i = 0 To UBound(pn)
  End Select ' Case neuMFGTyp
 End With ' aktMSH
 If obSp1Fest(azt(neuMFGTyp)) Then
  aktMSH.FixedCols = 1
 Else
  aktMSH.FixedCols = 0
 End If
 Call SizeColumns(aktMSH)
 If Me.MfGTyp <> neuMFGTyp Then
  fgespei(neuMFGTyp) = 0
  Me.MfGTyp = neuMFGTyp
 End If
' If merkCol(MfgTyp) = 0 Then
' End If
 aktMSH.Visible = True
 If nichtAusSpucken = 0 Then
  Call SpuckAus
 End If
' altFarbe(azgdp) = MFG.CellBackColor
 Screen.MousePointer = vbNormal
 DoEvents
 Exit Sub
EigFehler:
 Dim errnum&
 errnum = Err.Number
 If errnum = -2147217887 Then ' Der ODBC-Treiber unterstützt die angeforderten Eigenschaften nicht.
  dbv.wCn.Close
  dbv.wCn.Open
  Resume
 End If
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in MFGRefresh/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' MFGRefresh

' Make the FlexGrid's columns big enough to hold all values.
Private Sub SizeColumns(ByVal flx As MSHFlexGrid)
Const ZusatzWeite% = 100
Dim max_wid As Single
Dim wid As Single
Dim max_row As Integer
Dim r As Integer
Dim c As Integer
On Error GoTo fehler
If obdebug Then Debug.Print "SizeColumns("
    max_row = flx.Rows - 1
    For c = 0 To flx.Cols - 1
        max_wid = 0
        For r = 0 To max_row
            wid = Me.ucMDIKeys.TextWidth(flx.TextMatrix(r, c))
            If max_wid < wid Then max_wid = wid
        Next r
        flx.ColWidth(c) = max_wid + ZusatzWeite
    Next c
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in SizeColumns/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' SizeColumns

' Befüllung der Usertabelle: INSERT INTO dp.user(user,Passwort,hinzugefügt) VALUES(<user>,aes_encrypt(<passwort>,'0&F54'),NOW());
' oder: UPDATE dp.user set passwort = aes_encrypt(<pw>,'0&F54'), geändert=NOW() WHERE user='<user>';
Private Function prüfeUser%()
 Dim zp As Date
 On Error GoTo fehler
 zp = Now
 If zp - Usergeprüft > 1 / 24 / 60 Then
  User = vNS
 End If
 If LenB(User) = 0 Then
  frmL.txtPassword = vNS
  frmL.Show 1
  prüfeUser = frmL.LoginSucceeded
  If prüfeUser Then
   User = frmL.txtUserName
   MDI.loggen.Caption = User & " auslo&ggen"
  Else
   User = vNS
   MDI.loggen.Caption = "einlo&ggen"
  End If
'  Stop
 Else
  prüfeUser = True
 End If
 Usergeprüft = Now
 Exit Function
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in prüfeUser/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' prüfeUser

Private Sub MFG_Click()
 Dim Ctrl As Control, Fd0$, Fd1$, i%
 Dim rs As ADODB.Recordset
 On Error GoTo fehler
 If obdebug Then Debug.Print "MFG_Click("
 If MFG_click_abbrech Then
  MFG_click_abbrech = False
  Exit Sub
 End If
 If MfGTyp = azgpr Or MfGTyp = azgul Then Exit Sub
 cRow(MfGTyp) = Me.MFG.Row
 cCol(MfGTyp) = Me.MFG.Col
 If obdebug Then Debug.Print "MFG_Click(", "Row:", cRow(MfGTyp), "Col:", cCol(MfGTyp), "MFGTyp:", MfGTyp
'algo to position textbox inside flexgrid cells
'column 0 is used as a marker
 If Not prüfeUser Then Exit Sub
 If LenB(RefSp(azt(MfGTyp), DBSpvSp(azt(MfGTyp), cCol(MfGTyp)))) = 0 Then obtb = True Else obtb = False
 If obtb Then
  Set Ctrl = Me.Tb1
 Else
  Set Ctrl = Me.Cb1
  Cb1.Clear
'  Call rs.Open("SELECT `" & RefSp(DBSpvSp(ccol)) & "` FROM `" & RefTab(DBSpvSp(ccol)) & "`", Cn, adOpenDynamic, adLockReadOnly)
  sql = "SELECT `" & RefSp(azt(MfGTyp), DBSpvSp(azt(MfGTyp), cCol(MfGTyp))) & "`,`" & RefTab(azt(MfGTyp), DBSpvSp(azt(MfGTyp), cCol(MfGTyp))) & "`.* FROM `" & RefTab(azt(MfGTyp), DBSpvSp(azt(MfGTyp), cCol(MfGTyp))) & "`"
  If obdebug Then Debug.Print "SQL in MFG_Click:", sql
'  Call rs.Open(sql, dbv.wCn, adOpenDynamic, adLockReadOnly)
  myFrag rs, sql, adOpenStatic, dbv.wCn, adLockReadOnly
  For i = 1 To rs.Fields.Count - 2
   If LCase(rs.Fields(i).name) = LCase(RefSp(azt(MfGTyp), DBSpvSp(azt(MfGTyp), cCol(MfGTyp)))) Then
    Fd1 = rs.Fields(i + 1).name
    Exit For
   End If
  Next
  Do While Not rs.EOF
   Call Cb1.AddItem(rs.Fields(0) & IIf(Fd1 <> vNS, Chr(9) & rs.Fields(Fd1), vNS))
   rs.Move 1
  Loop
 End If
' AltInhalt = Me.MFG
 Ctrl.Visible = True
 On Error Resume Next
 Ctrl.Height = Me.MFG.CellHeight - 10 'minus 10 so that grid lines
 On Error GoTo fehler
 If Ctrl.name = "Tb1" Then
  Ctrl.Width = Me.MFG.CellWidth - 10 '  will not be overwritten
 Else
  Dim max_wid!, r%, wid!
        max_wid = 0
        For r = 0 To Cb1.ListCount - 1
            wid = Me.ucMDIKeys.TextWidth(Cb1.List(r))
            If max_wid < wid Then max_wid = wid
        Next r
  Ctrl.Width = max_wid
 End If
 Ctrl.Left = Me.MFG.CellLeft + Me.MFG.Left
 Ctrl.Top = Me.MFG.CellTop + Me.MFG.Top
 DoNotChange = True
 Ctrl.Text = Me.MFG.Text
' Cb1.Top = 300
' Cb1.Left = 0
 DoNotChange = False
' If ActiveControl <> MFG Then MFG.SetFocus
 If obdebug Then Debug.Print "MFG_Click: ", Ctrl.name, "Activecontrol:", ActiveControl.name
 Ctrl.SetFocus
 Ctrl.SelStart = Len(Ctrl.Text)
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in MFG_Click/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' MFG_Click()

Private Sub P1_KeyDown(KeyCode As Integer, Shift As Integer)
 If obdebug Then Debug.Print "P1_KeyDown(", KeyCode, Shift
 If KeyCode = 18 Or KeyCode = 17 Then
 Else
  Call Key(KeyCode, Shift, Me, ActiveControl)
 End If
End Sub ' P1_KeyDown

Private Sub MDIForm_KeyDown(KeyCode As Integer, Shift As Integer)
 If obdebug Then Debug.Print "MDIForm_KeyDown(", KeyCode, Shift
 If KeyCode = 18 Or KeyCode = 17 Then
 Else
  Call Key(KeyCode, Shift, Me, ActiveControl)
 End If
End Sub ' MDIForm_KeyDown

Public Function Key(KeyCode%, Shift%, frm As Form, Optional Ctrl As Control, Optional noF%)
 Dim aktSp%
 On Error GoTo fehler
 If obdebug Then Debug.Print "key(", KeyCode, Shift, frm.name, Ctrl.name
 If KeyCode = 16 Or KeyCode = 18 Or KeyCode = 17 Then ' Alt
' zuerst die Befehle, die oft hintereinander kommen können müssen
 ElseIf KeyCode = 46 And Ctrl.name = MFG.name Then ' Entf
   Call doChange(obLoe:=True, Qu:="key0")
'    Call doChange(obLoe:=0, Qu:="key7") ' überträgt cb1 auf die aktuelle Zelle
 ElseIf KeyCode = 67 And Shift = 2 And Ctrl.name = MFG.name Then ' Streuerung + C
  If prüfeUser Then
   Clip = Ctrl
  End If
 ElseIf KeyCode = 88 And Shift = 2 And Ctrl.name = MFG.name Then ' Steuerung + x
  If prüfeUser Then
   Clip = Ctrl
   Call doChange(obLoe:=True, Qu:="Ctrl+x")
  End If
 ElseIf KeyCode = 86 And Shift = 2 And Clip <> vNS And Ctrl.name = MFG.name Then
  If prüfeUser Then
'   obtb = True
   Cb1.Text = Clip
   Call doChange(obLoe:=False, Qu:="Ctrl+v")
  End If
 ElseIf (KeyCode >= 32 And KeyCode <= 255 And (KeyCode < 33 Or KeyCode > 40) And Shift < 2) And Ctrl.name = MFG.name Then
  Call MFG_Click ' auch KeyCode 113, F2
  If frm.ActiveControl.name <> MFG.name And frm.ActiveControl.name <> Me.ucMDIKeys.name Then  ' bei Schreibschutz, Protokoll
'  Debug.Print KeyCode, Chr(KeyCode)
  If KeyCode < 112 Or KeyCode > 123 Then
   If Shift = 0 Then
    frm.ActiveControl.Text = Lc(KeyCode)
   Else
    frm.ActiveControl.Text = uc(Lc(KeyCode))
   End If
  End If
  frm.ActiveControl.SelStart = Len(frm.ActiveControl.Text)
 End If
 ElseIf KeyCode = 33 And Ctrl.name <> Cb1.name And Ctrl.name <> Tb1.name Then
  Call einsweiter(stno)
'  Call StDirekt("{PGUP}", Shift)
 ElseIf KeyCode = 34 And Ctrl.name <> Cb1.name And Ctrl.name <> Tb1.name Then
  Call einsweiter(stnu)
'  Call StDirekt("{PGDN}", Shift)
 ElseIf KeyCode = 35 And Ctrl.name <> Cb1.name And Ctrl.name <> Tb1.name Then
  Call einsweiter(gre)
'  Call StDirekt("{END}", Shift)
 ElseIf KeyCode = 36 And Ctrl.name <> Cb1.name And Ctrl.name <> Tb1.name Then
  Call einsweiter(gli)
'  Call StDirekt("{HOME}", Shift)
 ElseIf (KeyCode = 37 And Ctrl.name <> Cb1.name And Ctrl.name <> Tb1.name) Or (KeyCode = 9 And (Shift And vbShiftMask) > 0) Then
  Call einsweiter(Lin)
'  Call StDirekt("{LEFT}", Shift)
 ElseIf KeyCode = 38 And Ctrl.name <> Cb1.name And Ctrl.name <> Tb1.name Then
  Call einsweiter(obe)
'  Call StDirekt("{UP}", Shift)
 ' rechts
 ElseIf (KeyCode = 39 And Ctrl.name <> Cb1.name And Ctrl.name <> Tb1.name) Or (KeyCode = 9 And (Shift And vbShiftMask) = 0) Then
  Call einsweiter(Rec)
'  Call StDirekt("{RIGHT}", Shift)
 ElseIf KeyCode = 40 And Ctrl.name <> Cb1.name And Ctrl.name <> Tb1.name Then
  Call einsweiter(unt)
'  Call StDirekt("{DOWN}", Shift)
 ElseIf KeyCode = 113 And Ctrl.name = MFG.name Then ' F2
   Call MFG_Click
 ElseIf KeyCode = 113 Then ' F2
'   Stop
 ElseIf KeyCode = 76 And Shift = 2 Then ' Ctrl + L
   Call doChange(obLoe:=True, Qu:="key1")
 ElseIf KeyCode = 78 And Shift = 2 Then ' Ctrl + N
'   If frm.ActiveControl = MFG Then
'   Else
'    Call doChange(obLoe:=0, Qu:="key1,5")
'   End If
'   Me.MFG.CellBackColor = vbWhite
'   Do While Me.MFG.Col <> IIf(obSp1Fest(azt(MfgTyp)), 1, 0)
    Call StDirekt("{HOME}", obvorher:=True)
'   Loop
'   Do While Me.MFG.Row <> MFG.Rows - 1
    Call StDirekt("{PGDN}", Shift:=vbCtrlMask, obweiter:=True)
'   Loop
'   Me.MFG.Row = Me.MFG.Rows - 1
'   Me.MFG.Col = IIf(obSp1Fest(azt(MfgTyp)), 1, 0)
   Call MFG_Click
 ElseIf KeyCode = 13 And (Ctrl.name = Tb1.name Or Ctrl.name = Cb1.name) Then ' Return
'   Call doChange(0, "key2")
   Call Me.ucMDIKeys.SetFocus
 ElseIf KeyCode = 27 And (Ctrl.name = Tb1.name Or Ctrl.name = Cb1.name) Then
'  Ctrl = AltInhalt
  Ctrl.Visible = False
'  Call MFGRefresh
  verwerfen = True
  Me.ucMDIKeys.SetFocus
 ElseIf KeyCode = 27 And MfGTyp = azgwp Then
  Select Case altMfgTyp
   Case azgdp
    Call MFGRefresh(azgdp)
   Case Else ' mitarbeiter
    Call zeigmitarbeiter_Click
  End Select
 ElseIf KeyCode = 27 Then
' Änderungen verwerfen, indem Ursprungswerte aus der Registry geholt werden,
' bevor die aktuellen Werte über form_unload wieder zurückgespeichert werden
'    frm.Hide
    If MfGTyp = azgdp Then
     ProgEnde Me
    Else
     Call MFGRefresh(azgdp)
    End If
 ElseIf KeyCode = 13 And frm.name = "PatAuswahl" And (frm.ActiveControl.name = "PatAuswahl" Or frm.ActiveControl.name = "Pat_id") Then
'  frm.Pat_id = frm.getPat_id(frm.PatName)
 ElseIf KeyCode = 13 And (Ctrl.name = Tb1.name Or Ctrl.name = Cb1.name) Then
  Call frm.doChange(0, "key3")
  frm.ActiveControl.Visible = False
  Me.ucMDIKeys.SetFocus
 ElseIf (KeyCode = 70 Or KeyCode = 220 Or KeyCode = 252) And ((Shift And vbAltMask) > 0) Then ' Farbauswahl zu Arten ' F, Ü, ü
  If MfGTyp = azgar Then
'   Call Farbauswahl_Click
   Dim obErr%
   FmCD.Move Me.Command1.Left + Me.Command1.Width + 35, Me.Command1.Top + 600
   FmCD.CmDlg.CancelError = True
   On Error Resume Next
   FmCD.CmDlg.Action = 3 'Aufruf des Dialogfensters. Es ist nur noch dieses ansprechbar, bis es mit Abbruch oder OK geschlossen wird.
'   fmcd.cmdlg.ShowColor
   If obdebug Then Debug.Print Err.Number
   obErr = Err.Number
   On Error GoTo fehler
   If obErr = 0 Then
    Dim altCol&
    altCol = MFG.Col
    If KeyCode = 70 Then
     MFG.Col = 0
     dbv.wCn.Execute "UPDATE `" & tbm(tbar) & "` SET farbe = " & FmCD.CmDlg.Color & " WHERE artnr = '" & MFG.Text & "'"
    Else
     For aktSp = 0 To UBound(SpNm, 2)
      If SpNm(azt(azgar), aktSp) = "Farbe" Then
       MFG.Col = SpvDBSp(azt(azgar), aktSp)
       Exit For
      End If
     Next aktSp
     dbv.wCn.Execute "UPDATE `" & tbm(tbar) & "` SET farbe = " & FmCD.CmDlg.Color & " WHERE farbe = " & MFG.Text
    End If
    Call MFGRefresh(azgar)
    MFG.Col = altCol
    Call SpuckAus
   End If
  End If
 ElseIf KeyCode = 85 And Shift = 4 Then ' Ctrl + U
  'Ausbezahlungen aufrufen
'  KeyCode = 0
  If MfGTyp = azgma Then
   merkCol(MfGTyp) = MFG.Col
   MFG.Col = 0
   If LenB(Me.MFG.Text) <> 0 Then
    Einschr(azgab) = "WHERE `" & SpNm(azt(MfGTyp), MFG.Col) & "` = " & IIf(FdTyp(azt(MfGTyp), MFG.Col) = adChar, "'", vNS) & MFG.Text & IIf(FdTyp(azt(MfGTyp), MFG.Col) = adChar, "'", vNS)
    EinsFd(azgab) = SpNm(azt(MfGTyp), MFG.Col)
    EinsWt(azgab) = "'" & MFG.Text & "'"
    MFG.Col = 2
    EinsNm(azgab) = "   Auszahlungen für " & MFG.Text & ", "
    MFG.Col = 3
    EinsNm(azgab) = EinsNm(azgab) & MFG.Text & " (Nr." & EinsWt(azgab) & ")"
    MFG.Col = merkCol(MfGTyp)
    Call frm.machAusbez(Einschr(azgab), EinsNm(azgab))
   End If 'LenB(Me.MFG.Text) <> 0 Then
  Else
'   Call frm.machWochenplan
  End If
 ElseIf KeyCode = 87 And Shift = 4 Then ' Ctrl + W
  'Wochenplan aufrufen
'  KeyCode = 0
  If MfGTyp = azgma Then
   merkCol(MfGTyp) = MFG.Col
   MFG.Col = 0
   If LenB(Me.MFG.Text) <> 0 Then
    Einschr(azgwp) = "WHERE `" & SpNm(azt(MfGTyp), MFG.Col) & "` = " & IIf(FdTyp(azt(MfGTyp), MFG.Col) = adChar, "'", vNS) & MFG.Text & IIf(FdTyp(azt(MfGTyp), MFG.Col) = adChar, "'", vNS)
    EinsFd(azgwp) = SpNm(azt(MfGTyp), MFG.Col)
    EinsWt(azgwp) = "'" & MFG.Text & "'"
    MFG.Col = 2
    EinsNm(azgwp) = "   Wochenplan für " & MFG.Text & ", "
    MFG.Col = 3
    EinsNm(azgwp) = EinsNm(azgwp) & MFG.Text & " (Nr." & EinsWt(azgwp) & ")"
    MFG.Col = merkCol(MfGTyp)
    Call frm.machWochenplan(Einschr(azgwp), EinsNm(azgwp))
   End If 'LenB(Me.MFG.Text) <> 0 Then
  Else
'   Call frm.machWochenplan
  End If
 End If
 If noF = 0 And Ctrl.name <> Tb1.name And Ctrl.name <> Cb1.name And Ctrl.name <> Tb1.name Then ' noF = -1 kommt vor in: ucMDIKeys_KeyDown
  Me.ucMDIKeys.SetFocus
 Else
'  Stop
 End If
' If KeyCode = 33 Then Call doRückwärtsCmd(frm)
' If KeyCode = 34 Then Call doVorwärtsCmd(frm) <- stellt den aktuellen Feldinhalt falsch ein!
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in key/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' key

Function StDirekt(StStrg$, Optional Shift%, Optional obvorher%, Optional obweiter%)
  On Error GoTo fehler
Dim SZ$, UZ$, az$ ' Zusatzzeichen für Sendkeys: Umschaltzeichen, Streuerungszeichen, Altzeichen
 If (Shift And vbShiftMask) > 0 Then UZ = "+" Else UZ = vNS
 If (Shift And vbAltMask) > 0 Then az = "%" Else az = vNS
 If (Shift And vbCtrlMask) > 0 Then SZ = "^" Else SZ = vNS
  nouc = True
'  If Not obvorher Then If Not Me.MFG.Visible Then Me.MFG.Visible = True: Me.MFG.SetFocus
  If Not Me.MFG.Visible Then Me.MFG.Visible = True
  If Me.ActiveControl.name <> Me.MFG.name Then Me.MFG.SetFocus
'  On Error Resume Next ' 7.9.15
'  SendKeys UZ & AZ & SZ & StStrg, True
'  SendKeysEx UZ & AZ & SZ & StStrg ' , True
'  MySendKeys UZ & AZ & SZ & StStrg, True
   Sendschluessel UZ & az & SZ & StStrg, True
'  On Error GoTo fehler
  nouc = False
  If Not obweiter Then Me.ucMDIKeys.SetFocus
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in StDirekt/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' StDirekt
Function Lc$(cd%)
 Select Case cd
  Case 222: Lc = "ä"
  Case 192: Lc = "ö"
  Case 186: Lc = "ü"
  Case 186: Lc = "+"
  Case 188: Lc = ","
  Case 189: Lc = "-"
  Case 190: Lc = "."
  Case 191: Lc = "#"
  Case 219: Lc = "ß"
  Case Else
   Lc = LCase(Chr(cd))
 End Select
End Function ' Lc$(cd%)

Function uc$(cd$)
 Select Case cd
  Case "1": uc = "!"
  Case "2": uc = """"
  Case "3": uc = "§"
  Case "4": uc = "$"
  Case "5": uc = "%"
  Case "6": uc = "&"
  Case "7": uc = "7"
  Case "8": uc = "("
  Case "9": uc = ")"
  Case "0": uc = "="
  Case "ß": uc = "?"
  Case ",": uc = ";"
  Case ".": uc = ":"
  Case "-": uc = "_"
  Case "+": uc = "*"
  Case "#": uc = "'"
  Case Else: uc = UCase(cd)
 End Select
End Function ' uc$(cd$)

Private Function fmtText()
  On Error GoTo fehler
  If FdTyp(azt(MfGTyp), SpvDBSp(azt(MfGTyp), MFG.Col)) = adDBDate Or FdTyp(azt(MfGTyp), SpvDBSp(azt(MfGTyp), MFG.Col)) = adDBTime Then
   If Not IsDate(MFG.Text) Then
    If IsNumeric(MFG.Text) Then
     fmtText = datform(MFG.Text)
    Else
     fmtText = "0" '"''"
     Exit Function
    End If
   Else
    fmtText = datform(CDate(MFG.Text))
   End If
  ElseIf FdTyp(azt(MfGTyp), SpvDBSp(azt(MfGTyp), MFG.Col)) = adChar Then
   fmtText = "'" & MFG.Text & "'"
  Else ' adnumeric = 131
   If LenB(MFG.Text) = 0 Then
    fmtText = "0"
'   ElseIf IsDate(MFG.Text) Then
'    fmtText = CDate(MFG.Text)
   ElseIf IsNumeric(MFG.Text) Then
    fmtText = Str(CDbl(MFG.Text))
   Else
    If MFG.Text Like "??,*.*.*" Then
     fmtText = datform(CDate(Mid(MFG.Text, 4)))
    End If
   End If
  End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in fmtText/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' fmtText

' augerufen in EinzelBilanz, MFGRefresh
Function UrlAnspr(ByVal PNr%, ByVal Jahr%, ByVal Cn, ByRef UAAh!, ByRef UAA!, Optional ByRef UABh!, Optional ByRef UAB!, Optional ByVal mitdruck%)
 On Error GoTo fehler
 Dim rs0 As ADODB.Recordset
 Dim iru%, sql$
' If Jahr = 2017 Then Stop
' If Jahr = 2019 Then Stop
 For iru = 0 To 1
  ' muss zuerst mit iru=0 und dann 1 augerufen werden, gv=Grenze von, gb=Grenze bis, relwaz=relevante Wochenarbeitszeit (am 1.1. des Jahres)
  ' jtage=Jahrestage(365 oder 366, ab ab):
  ' ,DATEDIFF(DATE(CONCAT(YEAR(IF(abk<@gv AND NOT @obvor,@gv,abk))+1,'0101')),DATE(CONCAT(YEAR(IF(abk<@gv AND NOT @obvor,@gv,abk)),'0101'))) jtage
'   If Jahr = 2018 Then Stop
   sql = "SELECT " & vbCrLf & _
   "COALESCE(SUM(urlh),0) urlh,COALESCE(SUM(urlh)/rwaz*5,0,0) urld,COALESCE(CONCAT('Urlaubsanspruch ',IF(obvor,'Übertrag','gesamt'),': ','\n',group_CONCAT(erkl separator '\n'),'\nInsgesamt: ', CAST(ROUND(SUM(urlh),1) AS CHAR(10)),' h, /', CAST(ROUND(rwaz*0.2,1) AS CHAR(10)),' = ', COALESCE( CAST(ROUND(SUM(urlh)/rwaz*5,1) AS CHAR(10)),'0'),' d'),'') Erkl " & vbCrLf & _
   "FROM (SELECT obvor,urlh, CONCAT(DATE_FORMAT(abk,'%d.%m.%y'),' - ',DATE_FORMAT(ADDDATE(bisk,-1),'%d.%m.%y'),' = ',LPAD(tage,3,' '),' Tage, rWAZ: ',LPAD(rwaz,5,' '),' h/Wo, Url: ',LPAD(i.urlaub,2,' '),' d/a => ',LPAD(ROUND(urlh,1),5,' '),' h, ',LPAD(IF(rwaz,ROUND(urlh/rwaz*5,1),''),4,' '),' d') Erkl,rwaz,persnr " & vbCrLf & _
   "  FROM (SELECT i.*,DATEDIFF(bisk,abk) tage, COALESCE(DATEDIFF(bisk,abk)/IF(obvor,365.24,DATEDIFF(gb,gv))*rwaz*0.2*urlaub,0) Urlh " & vbCrLf & _
   "   FROM (SELECT IF(ab<gv AND NOT obvor,gv,ab) abk, IF(bis>gb,gb,bis) bisk,rwaz,urlaub,persnr,gv,gb,obvor " & vbCrLf & _
   "    FROM (SELECT ab " & vbCrLf & _
   "     ,COALESCE((SELECT MIN(ab) FROM `" & tbm(tbwp) & "` WHERE ab>wp.ab AND persnr=wp.persnr) " & vbCrLf & _
   "      ,(SELECT IF(aus>0,aus,DATE(99991231)) FROM `" & tbm(tbma) & "` WHERE persnr=wp.persnr)) bis " & vbCrLf & _
   "     ,IF(waz=0,(SELECT waz FROM " & tbm(tbwp) & " WHERE persnr=wp.persnr AND waz<>0 AND ab<wp.ab ORDER BY ab DESC LIMIT 1),waz) rwaz " & vbCrLf & _
   "     ,urlaub,persnr,obvor,gv " & vbCrLf & _
   "     ,IF(obvor,gv,adddate(gv,interval 1 year)) gb " & vbCrLf & _
   "     FROM (SELECT " & CStr(iru) & " obvor,DATE(CONCAT(" & CStr(Jahr) & ",'0101')) gv,w.* FROM `" & tbm(tbwp) & "` w) wp " & vbCrLf & _
   "    WHERE persnr=" & CStr(PNr) & ") i " & vbCrLf & _
   "   WHERE  bis>IF(obvor,19000101,gv) AND ab<gb " & vbCrLf & _
   "  ) i " & vbCrLf & _
   " ) i " & vbCrLf & _
   ") i"
'   rs0.Open sql, Cn, adOpenStatic, adLockReadOnly
   If Cn.DefaultDatabase <> "dp" Then
    Cn.Close
    Cn.Open
   End If
   myFrag rs0, sql, adOpenStatic, Cn, adLockReadOnly
'  sql = "SELECT " & vbCrLf & _
   "@relwaz:=(SELECT IF(ab>" & Format(tzsbd, "yyyymmdd") & " AND waz,waz,38.5) FROM `" & tbm(tbwp) & "` wpi WHERE persnr=i.persnr AND COALESCE((SELECT MIN(ab) FROM `" & tbm(tbwp) & "` WHERE ab>wpi.ab AND persnr=wpi.persnr),(SELECT IF(aus>0,aus,DATE(99991231)) FROM `" & tbm(tbma) & "` WHERE persnr=wpi.persnr))>=@gv ORDER BY ab LIMIT 1) " & vbCrLf & _
   ",SUM(urlh) urlh,COALESCE(SUM(urlh)/@relwaz*5,0,0) urld,COALESCE(CONCAT('Urlaubsanspruch ',IF(@obvor,'Übertrag','gesamt'),': ','\n',group_CONCAT(erkl separator '\n'),'\nInsgesamt: ', CAST(ROUND(SUM(urlh),1) AS CHAR(10)),' h, /', CAST(@relwaz/5 AS CHAR(10)),' = ', COALESCE( CAST(ROUND(SUM(urlh)/@relwaz*5,1) AS CHAR(10)),'0'),' d'),'') Erkl " & vbCrLf & _
   "FROM ( " & vbCrLf & _
   " SELECT urlh, CONCAT(DATE_FORMAT(abk,'%d.%m.%y'),' - ',DATE_FORMAT(ADDDATE(bis,-1),'%d.%m.%y'),' WAZ: ',i.waz,' h/Wo, Url: ',i.urlaub,' d/a => ',ROUND(urlh,1),' h, ',IF(aktwaz,ROUND(urlh/aktwaz*5,1),''),' d') Erkl,persnr FROM ( " & vbCrLf & _
   "  SELECT i.*,DATEDIFF(bis,abk)*urlaub/(DATEDIFF(DATE(CONCAT(YEAR(IF(abk<@gv AND NOT @obvor,@gv,abk))+1,'0101')),DATE(CONCAT(YEAR(IF(abk<@gv AND NOT @obvor,@gv,abk)),'0101'))))*aktwaz/5 urlh FROM ( " & vbCrLf & _
   "   SELECT IF(ab<@gv AND NOT @obvor,@gv,ab) abk, IF(bis>@gb,@gb,bis) bis,waz,aktwaz,urlaub,persnr " & vbCrLf & _
   "    FROM ( " & vbCrLf & _
   "     SELECT ab " & vbCrLf & _
   "     ,COALESCE((SELECT MIN(ab) FROM `" & tbm(tbwp) & "` WHERE ab>wp.ab AND persnr=wp.persnr) " & vbCrLf & _
   "     ,(SELECT IF(aus>0,aus,DATE(99991231)) FROM `" & tbm(tbma) & "` WHERE persnr=wp.persnr)) bis " & vbCrLf & _
   "     ,waz,IF(ab>" & Format(tzsbd, "yyyymmdd") & ",waz,38.5) aktwaz,urlaub,persnr " & vbCrLf & _
   "     FROM `" & tbm(tbwp) & "` wp " & vbCrLf & _
   "     WHERE persnr=" & CStr(PNr) & ") i " & vbCrLf & _
   "    WHERE bis>=IF(@obvor:=" & CStr(iru) & ",19000101,@gv:=DATE(CONCAT(" & CStr(Jahr) & ",'0101'))) AND ab<@gb:=IF(@obvor,@gv,ADDDATE(@gv,INTERVAL 1 YEAR)) " & vbCrLf & _
   "   ) i " & vbCrLf & _
   "  ) i " & vbCrLf & _
   " ) i "

'  rs0.Open "SELECT sum(urlh) urlh,coalesce(sum(urlh)/@relwaz*5,0) urld,coalesce(CONCAT('Urlaubsanspruch ',if(obvor,'bisher','aktuell'),': ','\n',group_CONCAT(erkl separator '\n'),'\nInsgesamt: ', CAST(round(sum(urlh),1) as char(10)),' h, /', CAST(@relwaz/5 as char(10)),' = ', coalesce( CAST(round(sum(urlh)/@relwaz*5,1) as char(10)),'0'),' d'),'') Erkl FROM ( " & vbCrLf & _
  " select ab,bis,waz,urlaub,urlh, CONCAT(DATE_FORMAT(ab,'%d.%m.%y'),' - ',DATE_FORMAT(bis,'%d.%m.%y'),' (= ', CAST(DATEDIFF(bis,ab) as char(5)),' Tage), WAZ: ',i.waz,' h/Wo, Url: ',i.urlaub,' d/a => ',round(urlh,1),' h, ',round(urlh/aktwaz*5,1),' d') Erkl,obvor FROM ( " & vbCrLf & _
  "  select i.*,DATEDIFF(bis,ab)*urlaub/jtage*aktwaz/5 urlh,if(ab=gv,@relwaz:=aktwaz,0) FROM ( " & vbCrLf & _
  "   select IF(ab<gv AND NOT obvor,gv,ab) ab,DATEDIFF(date(CONCAT(YEAR(if(ab<gv AND NOT obvor,gv,ab))+1,'0101')),date(CONCAT(YEAR(if(ab<gv AND NOT obvor,gv,ab)),'0101'))) jtage,if(bis>gb,gb,bis) bis,waz,if(if(ab<gv AND NOT obvor,gv,ab)>=" & Format(tzsbd, "yyyymmdd") & ",waz,38.5) aktwaz,urlaub,gv,obvor FROM ( " & vbCrLf & _
  "    select ab " & vbCrLf & _
  "    ,coalesce((SELECT MIN(ab) FROM `" & tbm(tbwp) & "` WHERE ab>wp.ab and persnr=wp.persnr) " & vbCrLf & _
  "     ,(select IF(aus>0,aus,date(99991231)) FROM `" & tbm(tbma) & "` WHERE persnr=wp.persnr)) bis " & vbCrLf & _
  "    ,waz,urlaub " & vbCrLf & _
  "    ,gv,if(obvor,gv,adddate(gv,interval 1 year)) gb,obvor " & vbCrLf & _
  "    FROM ( " & vbCrLf & _
  "     select date(CONCAT(" & CStr(Jahr) & ",'0101')) gv, " & CStr(iru) & " obvor, wp.* " & vbCrLf & _
  "     FROM `" & tbm(tbwp) & "` wp " & vbCrLf & _
  "    ) wp " & vbCrLf & _
  "    WHERE persnr=" & CStr(PNr) & ")i " & vbCrLf & _
  "   WHERE bis>=if(obvor,1900,gv) and ab<gb " & vbCrLf & _
  "  ) i " & vbCrLf & _
  " ) i " & vbCrLf & _
  ") i", cn, adOpenStatic, adLockReadOnly
'  If Jahr = 2019 Then Stop
  If iru Then
   UABh = rs0!urlh
   UAB = rs0!urld
  Else
   UAAh = rs0!urlh
   UAA = rs0!urld
  End If
  If mitdruck And Not iru Or mitdruck = 2 Then
   If rs0!erkl <> "" Then
    Print #323, "<span style='background:#FFFF99'>" & rs0!erkl & "</span>"
   End If
  End If
  On Error GoTo setfehler
  Cn.Execute ("SET @relwaz=NULL")
  On Error GoTo fehler
  Set rs0 = Nothing
 Next iru
 Exit Function
setfehler:
 Cn.Close
 Cn.Open
 Resume
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in UrlAnspr/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' UrlAnspr

' aufgerufen in: gesBilanz, doChange
Function EinzelBilanz(ByVal Persnr&, ByRef ArtVgb$, ByRef ArtDp$, ByVal ausbez!, ByVal urlhaus!, ByRef UBilh!, ByRef UBil!, ByRef ÜBil!, _
                      ByRef FBil!, ByRef PSt!, ByVal akttag As Date, ByRef WAZv!, Optional ByVal mitdruck%, _
                      Optional ByVal ohneumr%, Optional ByVal obLoe%)
 Dim UAAh!, UAA!, UABh!, UAB!, WAZ!, VgbStd!, neu$ ' richtige Stunden im Feld; Inhalt des richtigen Feldes im Wochenplan, in Stunden
 Dim urlb! ' Urlaubswert des aktuellen Tages nach tagesarbeitszeitspezifischer Berechnung
' If akttag = #5/28/2018# Then Stop
 Static altPersNr&, altjahr%
 On Error GoTo fehler
 If Persnr <> altPersNr Then
  altPersNr = Persnr
  altjahr = 0
 End If
 If obLoe Then
  ArtDp = ArtVgb
  ArtVgb = ""
 End If
 Dim ur As ADODB.Recordset
' ur.Open "SELECT `" & Left(Format(akttag, "ddd"), 2) & "` ArtVgb,WAZ FROM `" & tbm(tbwp) & "` WHERE ab = (SELECT max(ab) FROM `" & tbm(tbwp) & "` WHERE ab <= " & datform(akttag) & " AND persnr = " & PersNr & ") AND persnr = " & PersNr, dbv.wCn, adOpenStatic, adLockReadOnly
 myFrag ur, "SELECT `" & Left(Format(akttag, "ddd"), 2) & "` ArtVgb,WAZ FROM `" & tbm(tbwp) & "` WHERE ab = (SELECT max(ab) FROM `" & tbm(tbwp) & "` WHERE ab <= " & datform(akttag) & " AND ab<>0 AND persnr = " & Persnr & ") AND persnr = " & Persnr, adOpenStatic, dbv.wCn, adLockReadOnly
 If ur.BOF Then
  WAZ = 0
 Else
  WAZ = CDbl(ur!WAZ)
  If LenB(ArtVgb) = 0 Then ArtVgb = ur!ArtVgb
 End If
 
 Dim ii%
 For ii = 0 To UBound(ftag) ' Feiertage
  If ftag(ii).Datum = akttag Then
   If ftag(ii).obhalb Then
    urlb = urlb * 0.5
    ArtVgb = "hFT"
   Else
    ArtVgb = "WF"
   End If
'     If True Or akttag > tzsbd Then ' #4/1/2020#
   Exit For
  End If
 Next ii
 
 Set ur = Nothing
 If IsNumeric(ArtVgb) Then VgbStd = CDbl(ArtVgb) Else VgbStd = 0
 If akttag < tzsbd And ArtVgb = "-" And ArtDp = "g" Then ArtDp = "" ' 28.2.17 Faschingsdienstag persnr 70
 neu = ArtDp
 If LenB(neu) = 0 Then neu = VgbStd
'  If akttag = #2/28/2017# Then Stop ' Faschingsdienstag 2017
' If WAZ = 0 Then WAZ = 38.5
'  If ohneumr = 0 And WAZ <> WAZv And (akttag = tzsbd Or (UBil <> 0 And akttag > tzsbd And WAZv <> 0)) Then ' von Thomas entdeckter Korrekturbedarf
 If ohneumr = 0 And WAZ <> WAZv And UBil Then ' von Thomas entdeckter Korrekturbedarf
  Dim WAZd!, WAZvd!
  If WAZ = 0 Then WAZd = 38.5 Else WAZd = WAZ
  If WAZv = 0 Then WAZvd = 38.5 Else WAZvd = WAZv
  If mitdruck Then
   Print #323, "<span style='background:#FFCC99'>Urlaubsanspruchumrechnung: "
   Print #323, "Urlaubsbilanz nachher " & IIf(WAZ = 0, "[umger.auf 38,5h/Wo]", "") & " = Urlaubsbilanz vorher " & IIf(WAZv = 0, "[umger.auf 38,5h/Wo]", "") & " * Wochenarbeitszeit vorher " & IIf(WAZv = 0, "[umger.auf 38,5h/Wo]", "") & "/ Wochenarbeitszeit nachher" & IIf(WAZ = 0, " [umger.auf 38,5h/Wo]", "")
   Print #323, CStr(-Round(UBil * WAZvd / WAZd, 2)) & " = " & CStr(-Round(UBil, 2)) & " * " & CStr(Round(WAZvd, 2)) & " / " & CStr(WAZd) & "</span>"
  End If
  UBil = UBil * WAZvd / WAZd
 End If
 If akttag < tzsbd Then
  urlb = 1
 Else
'   If Day(akttag) = 1 And Month(akttag) = 1 Then Stop
  If IsNumeric(ArtVgb) And WAZ <> 0 Then urlb = 5 * (VgbStd / WAZ) Else urlb = 0 ' 1
 End If ' akttag < tzsbd
 If Year(akttag) <> altjahr Then
  Call FTbeleg(Year(akttag))
  If ohneumr = 0 Then
   Call UrlAnspr(Persnr, Year(akttag), dbv.wCn, UAAh, UAA, UABh, UAB, mitdruck:=IIf(mitdruck, IIf(altjahr, 1, 2), 0))
   UBilh = UBilh - UAAh
   UBil = UBil - UAA
   If altjahr = 0 Then
    UBilh = UBilh - UABh
    UBil = UBil - UAB
   End If
  End If ' ohneumr=0
  altjahr = Year(akttag)
 End If ' Year(akttag) <> altjahr Then
' hier wurde vor die ArtVgb durch Feiertage korrigiert
 WAZv = WAZ
'  If akttag = #1/24/2011# Then Stop
  ' vorgezogen, da ArtDp dann verändert wird
' 2. Teil: Bilanzen so ändern, als wenn statt einem "-" das Folgende eingetragen würde:
  Select Case ArtDp
   Case "b", "hFT", "WF", "k", "ki", "f", "fw", "g", "u", "uw", "su"
    If obLoe Then ÜBil = ÜBil - VgbStd Else ÜBil = ÜBil + VgbStd
  End Select
  Select Case neu
   Case vNS, "-", "b", "WF", "hFT", "k", "ki", "ü", "su"
   Case "f", "fw":      FBil = FBil + 1
   Case "g", "u", "uw": If obLoe Then UBil = UBil - urlb: UBilh = UBilh - urlb * WAZ * 0.2 Else UBil = UBil + urlb: UBilh = UBilh + urlb * WAZ * 0.2 ' 1 ' ArtVgb/WAZ
'  Case "ü":            ÜBil = ÜBil - VgbStd
    If Not obLoe Then ArtDp = ArtDp & " " & CStr(Round(urlb, 1))
   Case Else
    If IsNumeric(neu) Then
     If obLoe Then ÜBil = ÜBil - neu Else ÜBil = ÜBil + neu
    End If
  End Select
' 1. Teil: Bilanzen so ändern, als wenn statt dem Folgenden "-" eingetragen würde:
  Select Case ArtVgb
   Case "b", "hFT", "WF", "k", "ki", "f", "fw", "g", "u", "uw", "su"
    If obLoe Then ÜBil = ÜBil + VgbStd Else ÜBil = ÜBil - VgbStd
  End Select
  Select Case ArtVgb
   Case vNS, "-", "b", "WF", "hFT", "k", "ki", "ü", "su"
   Case "f", "fw":      FBil = FBil - 1
   Case "g", "u", "uw": If obLoe Then UBil = UBil + urlb: UBilh = UBilh + urlb * WAZ * 0.2 Else UBil = UBil - urlb: UBilh = UBilh - urlb * WAZ * 0.2 ' 1
    ArtDp = ArtDp & " " & CStr(Round(urlb, 1))
'  Case "ü": ÜBil = ÜBil + VgbStd
   Case Else
    If obLoe Then ÜBil = ÜBil + VgbStd Else ÜBil = ÜBil - VgbStd
  End Select
  If obLoe Then PSt = PSt - VgbStd Else PSt = PSt + VgbStd
  If ausbez Then
  ÜBil = ÜBil - ausbez
  End If
  If urlhaus Then UBilh = UBilh + urlhaus: UBil = UBil + urlhaus / WAZ * 5
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Einzelbilanz/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' EinzelBilanz

' aufgerufen in: Cb1_LostFocus, Tb1_LostFocus, Zeilenauswahl_Click, Key
Private Sub doChange(Optional obLoe%, Optional Qu$, Optional norefresh As Boolean)
 Dim erg&
 Dim nText$, fnText$, sqlwhere$, j%, jj%, aCol&, aRow&, eCol&, eRow&, jCol&, jRow&, i%, obIndSp%, k&, sql$, altText, imCol&
 Dim pos%, eingefügt%, obProt%
 Dim rs As ADODB.Recordset
 Dim ArtVgb$
 Dim akttag As Date
 Dim updrs As ADODB.Recordset
 On Error GoTo fehler
 If obdebug Then Debug.Print "doChange(", obLoe, Qu, Tb1, Cb1
 If obLoe Then If Not prüfeUser Then Exit Sub
 noenter = -1
 With Me.MFG
  aCol = .Col
  eCol = .ColSel
  aRow = .Row
  eRow = .RowSel
   If obLoe Then
    cRow(MfGTyp) = min(cRow(MfGTyp) - 1, 1)
   Else
    .Row = cRow(MfGTyp)
    .Col = cCol(MfGTyp)
   End If
'  Set rstu = Nothing
  If DoNotChange Then Exit Sub
  Select Case MfGTyp
   Case azgdp ' Dienstplan
    MFG.Redraw = False
    If aCol <> eCol Or aRow <> eRow Then
     erg = MsgBox("Wollen Sie alle markierten Zellen " & IIf(obLoe, "lösch", "befüll") & "en?", vbYesNo)
     If erg = vbNo Then GoTo Ende
' Zellenübergreifender Code
    End If
    If Not obLoe Then
     pos = InStr(Cb1.Text, Chr(9))
     If pos > 0 Then nText = Left(Cb1.Text, pos - 1) Else nText = Cb1.Text
    End If
'    dbv.wCn.Execute ("START TRANSACTION")
    For jCol = aCol To eCol
     Dim WAZv! ' WAZ!, Wochenarbeitszeit, ~ des Vortags
'     WAZ = 0
     For jRow = aRow To eRow
      akttag = BegD + jRow - 1 - SZZ
      If obLoe Then
       sql = "SELECT `artnr` FROM `" & tbm(tbdp) & "` WHERE `tag` = " & datform(akttag) & " AND `PersNr` = " & pn(jCol - 1)
'       rs.Open sql, dbv.wCn, adOpenStatic, adLockOptimistic
       Set rs = Nothing
       myFrag rs, sql, adOpenStatic, dbv.wCn, adLockReadOnly
       If rs.BOF Then
        ArtVgb = vNS
       Else
        ArtVgb = rs!artnr
        erg = MsgBox("Aus `" & tbm(tbdp) & "` soll folgendes gelöscht werden:" & vbCrLf & "Tag: " & datform(akttag) & vbCrLf & "Persnr: " & pn(jCol - 1) & vbCrLf & "Wollen Sie es nicht doch behalten?", vbYesNo, "Sicherheitsrückfrage")
        If erg = vbNo Then
   '     MsgBox "Datum zum Glück nicht löschbar!"
         sql = "DELETE FROM `" & tbm(tbdp) & "` WHERE `tag` = " & datform(akttag) & " AND `PersNr` = " & pn(jCol - 1)
   '     On Error Resume Next
'         dbv.wCn.Execute sql, rAF
         Set rs = Nothing
         myFrag rs, sql, , dbv.wCn, , , rAf
         If rAf <> 0 Then
          obProt = True
          nText = vNS
         End If
        End If
       End If ' rs.BOF
       Set rs = Nothing
      Else ' obLoe
       .Col = jCol
       .Row = jRow
       If nText <> .Text Then
        sql = "SELECT `artnr` FROM `" & tbm(tbdp) & "` WHERE `tag` = " & datform(akttag) & " AND `PersNr` = " & pn(jCol - 1)
'        rs.Open sql, dbv.wCn, adOpenStatic, adLockOptimistic
        myFrag rs, sql, adOpenStatic, dbv.wCn, adLockReadOnly
        If rs.BOF Then
         ArtVgb = vNS
        Else
         ArtVgb = rs!artnr
        End If
        Set rs = Nothing
        Dim maname$
        Dim mars As ADODB.Recordset
        Set mars = Nothing
'        mars.Open "SELECT nachname FROM `" & tbm(tbma) & "` WHERE persnr = " & pn(jCol - 1), dbv.wCn, adOpenStatic, adLockReadOnly
        myFrag mars, "SELECT nachname FROM `" & tbm(tbma) & "` WHERE persnr = " & pn(jCol - 1), adOpenStatic, dbv.wCn, adLockReadOnly
        If Not mars.EOF Then
         If LCase(mars!Nachname) Like "*notiz*" Then
          dbv.wCn.Execute ("INSERT IGNORE INTO `" & tbm(tbar) & "`(artnr,farbe,zusatz) VALUES('" & nText & "',14671839,1)")
         End If
        End If
        sql = "UPDATE `" & tbm(tbdp) & "` set `artnr` = '" & nText & "' WHERE `tag` = " & datform(akttag) & " AND `PersNr` = " & pn(jCol - 1)
'        Call dbv.wCn.Execute(sql, rAF)
        Set updrs = Nothing
        myFrag updrs, sql, , dbv.wCn, , , rAf
        
        If rAf = 0 Then
         Set updrs = Nothing
         sql = "INSERT INTO `" & tbm(tbdp) & "`(`ArtNr`,`tag`, `PersNr`) VALUES('" & nText & "'," & datform(akttag) & ",'" & pn(jCol - 1) & "')"
         On Error Resume Next
 '        Call dbv.wCn.Execute(sql, rAF)
         myFrag updrs, sql, , dbv.wCn, , , rAf
         If Err.Number <> 0 Then
          MsgBox "Dienstplanart " & nText & " hier im Moment nicht vorgesehen!" & vbCrLf & "Fehlermeldung: " & Err.Description
         End If
         On Error GoTo fehler
  '      Else
  '       MsgBox "Datenbankfehler! mehrere Datensätze auf einen Schlag geändert mit: " & "WHERE `tag` = " & datform(akttag) & " AND `PersNr` = " & pn(jcol - 1)
        End If
        If rAf <> 0 Then
         obProt = True
         .Text = nText
  '       Call Einfärben(mitaltFar:=True)
        End If
  '     Call MFGRefresh(Me.MfgTyp)
       End If ' nText <> .Text Then
      End If
      If obProt Then
       sql = "INSERT INTO `" & tbm(tbpr) & "` (`tag`,`PersNr`,`ArtNrV`,`ArtNr`,`AendDat`,`AendPC`,`AendUser`, `user`) values (" & datform(akttag) & "," & pn(jCol - 1) & ",'" & ArtVgb & "','" & nText & "'," & datform(Now) & ",'" & CptName & "','" & UserName & "','" & User & "')"
'       Call dbv.wCn.Execute(sql, rAF)
       Set updrs = Nothing
       myFrag updrs, sql, , dbv.wCn
       On Error GoTo fehler
       Dim UBilh!, UBil!, ÜBil!, FBil!, obBilanzNeu%, PSt! ' Urlaubsbilanz, Überstundenbilanz, Fortbildungsbilanz, , Planstunden
'       rs.Open "SELECT urlstd, urlaub, Überstunden uest, Fortbildung FROM `" & tbm(tbbi) & "` WHERE persnr = " & pn(jCol - 1) & " AND jahr = " & Me.Jahr, dbv.wCn, adOpenStatic, adLockReadOnly
       Set rs = Nothing
       myFrag rs, "SELECT urlstd, urlaub, Überstunden uest, Fortbildung FROM `" & tbm(tbbi) & "` WHERE persnr = " & pn(jCol - 1) & " AND jahr = " & Me.Jahr, adOpenStatic, dbv.wCn, adLockReadOnly
       If rs.EOF Then
        obBilanzNeu = True
       Else
        UBilh = rs!urlstd
        UBil = rs!urlaub
        ÜBil = rs!uest
        FBil = rs!Fortbildung
       End If
       Set rs = Nothing
'       Dim rsf! ' rsf = richtige Stunden im Feld; Inhalt des richtigen Feldes im Wochenplan, in Stunden
'       rsf = 0
'       If Not rs.BOF Then If IsNumeric(rs.Fields(0)) Then rsf = rs.Fields(0)
       If LenB(ArtVgb) <> 0 And LenB(nText) <> 0 Then ' wenn schon was drin stand, dann das zuerst löschen
        Call EinzelBilanz(pn(jCol - 1), ArtVgb, "", 0, 0, UBilh, UBil, ÜBil, FBil, PSt, akttag, WAZv, mitdruck:=False, ohneumr:=True, obLoe:=True)
        ArtVgb = ""
       End If
       Call EinzelBilanz(pn(jCol - 1), ArtVgb, nText, 0, 0, UBilh, UBil, ÜBil, FBil, PSt, akttag, WAZv, mitdruck:=False, ohneumr:=True, obLoe:=obLoe)
       Set updrs = Nothing
       If obBilanzNeu Then
'        Call dbv.wCn.Execute("INSERT INTO `" & tbm(tbbi) & "`(urlstd,urlaub,überstunden,fortbildung,persnr, jahr,planstunden) VALUES('" & Str$(UBilh) & "','" & Str$(UBil) & "','" & Str$(ÜBil) & "','" & Str$(FBil) & "'," & pn(jCol - 1) & "," & Me.Jahr & ",'" & Str$(PSt) & "')")
        myFrag updrs, "INSERT INTO `" & tbm(tbbi) & "`(urlstd,urlaub,überstunden,fortbildung,persnr, jahr,planstunden) VALUES('" & Str$(UBilh) & "','" & Str$(UBil) & "','" & Str$(ÜBil) & "','" & Str$(FBil) & "'," & pn(jCol - 1) & "," & Me.Jahr & ",'" & Str$(PSt) & "')", , dbv.wCn
       Else
'        Call dbv.wCn.Execute("UPDATE `" & tbm(tbbi) & "` SET urlstd = '" & Str$(UBilh) & "',urlaub = '" & Str$(UBil) & "',überstunden = '" & Str(ÜBil) & "',fortbildung = '" & Str(FBil) & "',planstunden = '" & Str$(PSt) & "' WHERE persnr = " & pn(jCol - 1) & " AND jahr = " & Me.Jahr)
        myFrag updrs, "UPDATE `" & tbm(tbbi) & "` SET urlstd = '" & Str$(UBilh) & "',urlaub = '" & Str$(UBil) & "',überstunden = '" & Str(ÜBil) & "',fortbildung = '" & Str(FBil) & "',planstunden = '" & Str$(PSt) & "' WHERE persnr = " & pn(jCol - 1) & " AND jahr = " & Me.Jahr, , dbv.wCn
       End If ' obBilanzNeu else
      End If ' obProt Then
     Next jRow
    Next jCol
'    dbv.wCn.Execute ("COMMIT")
    MFG.Redraw = True
    If Not norefresh Then Call MFGRefresh(azgdp)
   Case Else ' nicht dienstplan
     Call merken(MfGTyp)
     sqlwhere = vNS 'Einschr(MfgTyp)
     imCol = MFG.Col
     For j = 0 To MaxPrim
      If LenB(PrimI(azt(MfGTyp), j)) = 0 Then Exit For
      sqlwhere = sqlwhere & " AND `" & PrimI(azt(MfGTyp), j) & "` = "
      For i = 0 To UBound(SpNm, 2) ' UBound(SpvDBSp, 2)
       If LenB(PrimI(azt(MfGTyp), j)) = 0 Then Exit For
       If SpNm(azt(MfGTyp), i) = PrimI(azt(MfGTyp), j) Then
        MFG.Col = i 'SpvDBSp(MfgTyp, i)
        If .Col = imCol Then
         obIndSp = -1
         Exit For
        End If
       End If
      Next i
      sqlwhere = sqlwhere & fmtText()
     Next
     MFG.Col = imCol
     If Left(sqlwhere, 5) = " AND " Then sqlwhere = Mid(sqlwhere, 5)
    If obLoe Then
     On Error Resume Next
     If prüfeUser Then
     erg = MsgBox("Aus `" & Tabl & "` soll folgendes gelöscht werden:" & vbCrLf & sqlwhere & vbCrLf & "Wollen Sie es nicht doch behalten?", vbYesNo, "Sicherheitsrückfrage")
     If erg = vbNo Then
      dbv.wCn.Execute "DELETE FROM `" & Tabl & "` WHERE " & sqlwhere, rAf
      If Err.Number = 0 Then
       MsgBox "Aus `" & Tabl & "` wurde" & IIf(rAf = 1, " ", "n ") & rAf & " Zeile" & IIf(rAf = 1, "", "n") & " gelöscht:" & vbCrLf & sqlwhere, , "Rückmeldung von der Datenbank"
      ElseIf Err.Number = -2147467259 Then
       Dim ErrDesc$
       ErrDesc = Err.Description
       Dim p1%
       p1 = InStr(Err.Description, "constraint fails")
       If p1 > 0 Then
        Dim p2%, p3%, p4%, p5%, Erkl0$, Erkl1$
        Dim T1$, T2$, K1$
        p2 = InStr(p1, ErrDesc, "`")
        p3 = InStr(p2 + 1, ErrDesc, "`,")
        T1 = Mid(ErrDesc, p2 - 1, p3 - p2)
        T1 = Mid(ErrDesc, p2, p3 - p2 + 1)
        T1 = Mid(ErrDesc, p2, p3 - p2 + 1)
        p2 = InStr(p1, ErrDesc, "FOREIGN KEY (`")
        p3 = InStr(p2, ErrDesc, "`)")
        K1 = Mid(ErrDesc, p2 + 13, p3 - p2 - 12)
        Dim ranz As New ADODB.Recordset
        Set ranz = Nothing
        Err.Clear
        ranz.Open "SELECT COUNT(0) FROM " & T1 & " WHERE " & sqlwhere, dbv.wCn, adOpenStatic, adLockReadOnly
        Erkl0 = ranz.Fields(0)
        Set ranz = Nothing
        ranz.Open "SELECT * FROM " & T1 & " WHERE " & sqlwhere, dbv.wCn, adOpenStatic, adLockReadOnly
        Erkl1 = vNS
        For i = 0 To ranz.Fields.Count - 1
         Erkl1 = Erkl1 & ranz.Fields(i).name & ": " & ranz.Fields(i) & "; "
        Next i
        Dim src$
        src = ranz.source
        MsgBox "Löschvorgang von '" & .Text & "' nicht erfolgreich, da noch " & Erkl0 & " Datensätze:" & vbCrLf & Erkl1 & Erkl1 & "(ermittelt durch: '" & src & "')" & vbCrLf & " in folgender Abhängigkeit: " & vbCrLf & ErrDesc
       End If
      End If
      On Error GoTo fehler
      Call MFGRefresh(Me.MfGTyp)
'     .CellBackColor = vbWhite
      If Me.MFG.Row > Me.MFG.Rows - 2 Then Me.MFG.Row = Me.MFG.Rows - 2
     End If ' erg = vbNo
     End If ' prüfeuser
     Call MFG_Entercell
    Else ' Ändern oder einfügen
     If prüfeUser Then
     altText = .Text
     If obtb = True Then
      .Text = Me.Tb1.Text
     ElseIf obtb = False Then
      .Text = Me.Cb1.Text
      pos = InStr(.Text, Chr(9))
      If pos > 0 Then
       .Text = Left(.Text, pos - 1)
      End If
     End If
     If .Text <> altText Or LenB(.Text) = 0 Then
      nText = .Text
      fnText = fmtText()
 ' Einfügen
      If obIndSp Or (obAuto(azt(MfGTyp)) And .Row = MFG.Rows - 1) Then
       On Error Resume Next
       If SpNm(azt(MfGTyp), MFG.Col) <> EinsFd(MfGTyp) Then
        sql = "INSERT INTO `" & Tabl & "`(`" & SpNm(azt(MfGTyp), MFG.Col) & "`" & IIf(LenB(Einschr(MfGTyp)) = 0, vNS, ",`" & EinsFd(MfGTyp) & "`") & ") VALUES(" & fnText & IIf(LenB(Einschr(MfGTyp)) = 0, vNS, "," & EinsWt(MfGTyp)) & ")"
        Call dbv.wCn.Execute(sql)
        If Err.Number <> 0 Then
         If MfGTyp <> azgwp And MfGTyp <> azgar Then
          MsgBox "Fehler " & Err.Number & " beim Einfügen in die Datenbank mit dem Befehl:" & vbCrLf & sql & vbCrLf & "Fehlerbeschreibung: " & Err.Description & vbCrLf & "DLLError: " & Err.LastDllError
         End If
        Else
         eingefügt = True
         If MfGTyp = azgar Then
          sql = "UPDATE `" & Tabl & "` set farbe = 33023 WHERE `" & SpNm(azt(MfGTyp), MFG.Col) & "` = " & fnText ' orange setzen
          Call dbv.wCn.Execute(sql)
         End If
        End If
       End If
       On Error GoTo fehler
       Call MFGRefresh(Me.MfGTyp, nichtAusSpucken:=True)
'       .CellBackColor = vbWhite ' wird gebraucht beim Neueinfügen eines Datensatzes
       .Col = jCol
       For k = 1 To MFG.Rows - 1
        .Row = k
        If .Text = nText Then
         jRow = .Row
         Exit For
        End If
       Next
       Call MFG_Entercell
      Else
' Ändern
       On Error Resume Next
       sql = "UPDATE `" & Tabl & "` set `" & SpNm(azt(MfGTyp), MFG.Col) & "` = " & fnText & " WHERE " & sqlwhere
       Call dbv.wCn.Execute(sql)
       If Err.Number <> 0 Then
        MsgBox "Fehler " & Err.Number & " beim Aktualisieren der Datenbank mit dem Befehl:" & vbCrLf & sql & vbCrLf & "Fehlerbeschreibung: " & Err.Description & vbCrLf & "DLLError: " & Err.LastDllError
       End If
       On Error GoTo fehler
 '      Call cellweiter
      End If ' obIndSp OR (obAuto(azt(MFGTyp)) AND .Row = MFG.Rows - 1) Then
      If azt(MfGTyp) = tbwp And Not eingefügt Then
       Dim summe!, altRow&, altCol%
       summe = 0
       altRow = MFG.Row
       altCol = MFG.Col
       .Row = 0
       For i = 1 To MFG.Cols - 1
        .Col = i
        If .Text = "Montag" Then
         .Row = altRow
         For j = 0 To 6
          If IsNumeric(.Text) Then summe = summe + .Text
'          If .Text Like "a*" AND IsNumeric(Mid(.Text, 2)) Then Summe = Summe + Mid(.Text, 2)
          .Col = .Col + 1
         Next j
         i = i + 6
         .Row = 0
        ElseIf .Text = "WAZ" Then
         .Row = altRow
         If .Text <> Str(summe) Then
          .Text = summe
          sql = "UPDATE `" & Tabl & "` set `WAZ` = '" & Str(summe) & "' WHERE " & sqlwhere
          Call dbv.wCn.Execute(sql)
         End If
        End If
        .Col = altCol
       Next i
       Call Einfärben(mitaltFar:=True)
      End If
     End If ' .Text <> altText Then
    End If ' prüfeuser
   End If ' obLoe / else
  End Select ' case MfGTyp
  Call mfg_leavecell
  .Row = min(aRow, .Rows - 1)
  .Col = aCol
  noenter = 0
  Call MFG_Entercell
 End With
Ende:
 noenter = 0
 Tb1.Visible = False
 Cb1.Visible = False
 Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doChange/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' doChange

Private Function Einfärben(mitaltFar%)
 Call rsf.Find("artnr = '" & MFG.Text & "'", , adSearchForward, 1)
 If Not rsf.EOF Then
  MFG.CellBackColor = rsf!Farbe
'  If MFG.CellBackColor = 0 Then Stop
  If mitaltFar Then
   altFarbe(MfGTyp) = rsf!Farbe
  End If
 End If
End Function ' einfärben

Private Sub Command2_click() ' Der Hauptteil wird schon in key() verarbeitet
 If obdebug Then Debug.Print "Command2_click("
 Call merken(MfGTyp)
 If MfGTyp = azgma Then
  Call Key(85, vbAltMask, Me, ActiveControl)
 ElseIf MfGTyp = azgar Then
  Call Key(220, vbAltMask, Me, ActiveControl)
  Me.ucMDIKeys.SetFocus
 End If
' Me.MFG.SetFocus
End Sub ' Command2_click

Private Sub Command1_click() ' Der Hauptteil wird schon in key() verarbeitet
 If obdebug Then Debug.Print "Command1_click("
 Call merken(MfGTyp)
 If MfGTyp = azgma Then
  Call Key(87, vbAltMask, Me, ActiveControl)
 ElseIf MfGTyp = azgar Then
  Call Key(70, vbAltMask, Me, ActiveControl)
  Me.ucMDIKeys.SetFocus
 End If
' Me.MFG.SetFocus
End Sub ' Command1_click

Private Sub zeigdienstplan_KeyDown(KeyCode As Integer, Shift As Integer)
 If obdebug Then Debug.Print "zeigdienstplan_keydown(", KeyCode, Shift
 Call Key(KeyCode, Shift, Me, ActiveControl)
End Sub ' CancelButton_KeyDown

Private Sub zeigmitarbeiter_KeyDown(KeyCode As Integer, Shift As Integer)
 If obdebug Then Debug.Print "zeigmitarbeiter_Keydown(", KeyCode, Shift
 Call Key(KeyCode, Shift, Me, ActiveControl)
End Sub ' CancelButton_KeyDown

Private Sub Cb1_LostFocus()
 If obdebug Then Debug.Print "cb1_lostfocus(", "an ->", ActiveControl.name
 If verwerfen Then
  verwerfen = 0
 Else
  Call doChange(0, "cb1_lostfocus")
 End If
 Me.ucMDIKeys.SetFocus
End Sub ' Cb1_LostFocus

Private Sub Pfeilu_Click()
 If obdebug Then Debug.Print "Pfeilu_Click("
 MFG.SetFocus
 On Error Resume Next ' 7.9.15
 SendKeys "{DOWN}", 1
End Sub ' private sub Pfeilu_Click

' Menüeintrag, zur Zeit fehlend
Private Sub Seitu_Click()
 Me.MFG.Redraw = False
 Me.MFG.TopRow = min(Me.MFG.Row + 70, Me.MFG.Rows - 11)
 Me.MFG.Row = min(Me.MFG.Row + 80, Me.MFG.Rows - 1)
 Me.MFG.Redraw = True
End Sub ' Seitu_Click

Private Sub Tb1_Click()
 If obdebug Then Debug.Print "Tb1_Click("
End Sub ' Tb1_Click

Private Sub Tb1_GotFocus()
 If obdebug Then Debug.Print "Tb1_GotFocus("
 aCtl = Tb1.name
End Sub ' Tb1_GotFocus

Private Sub Tb1_KeyDown(KeyCode As Integer, Shift As Integer)
 If obdebug Then Debug.Print "Cb1_KeyDown(", KeyCode, Shift
 If KeyCode = 18 Or KeyCode = 17 Then
 Else
  Call Key(KeyCode, Shift, Me, ActiveControl)
 End If
End Sub ' Tb1_KeyDown

Private Sub Tb1_LostFocus()
 If obdebug Then Debug.Print "tb1_lostfocus(", "an ->", ActiveControl.name
 If verwerfen Then
  verwerfen = 0
 Else
  Call doChange(0, "tb1_lostfocus")
 End If
Me.ucMDIKeys.SetFocus
End Sub ' Tb1_LostFocus

Private Sub MDIForm_Load()
 On Error GoTo fehler
 If MfGTyp = azgnix Then MfGTyp = azgdp
 If obdebug Then Debug.Print "MDIForm_Load("
 pVerz = IIf(Dir("p:\") <> "", "p:\", "\\linux1\Daten\Patientendokumente\")
 Set dbv = New DBVerb
 With Me
  .Top = 0
  .Left = 0
  .Height = Screen.Height - 600
  .Width = Screen.Width
  With .ucMDIKeys
   .Top = Me.Top
   .Height = Me.Height
  End With
  Call wähleBenutzer
  Call dbv.cnVorb("dp", tbm(tbwp), tbm(tbdp)) '("", "--multi", vns)
  On Error GoTo fehler
  .Cb1.Visible = False
  .Tb1.Visible = False
  .Command1.Visible = False
  .Command2.Visible = False
  .Visible = True
  .MFG.Top = 400
  .MFG.Height = .Height - .MFG.Top - 1000
  .MFG.Left = 0
  .MFG.Width = .Width - .MFG.Left - 200
  .MFG.AllowUserResizing = flexResizeBoth
 End With
 Call TabAnalyse
 Dim Ctl As Control
 On Error Resume Next
 For Each Ctl In Me.Controls
  If Ctl.name <> ucMDIKeys.name Then
   Ctl.TabStop = False
  End If
 Next Ctl
 On Error GoTo fehler
 Me.Jahr = Year(Now)
' merkRow(azgdp) = Int(Now()) - CDate("1.1." & Year(Now)) + 1 + SZZ
 merkRow(azgdp) = DatePart("y", Now) + SZZ
 merkTop(azgdp) = merkRow(azgdp) - 10
 Me.Jahr = Year(Now)
 If mitVGF Then
  Set rs = Nothing
  Call rs.Open("SELECT einstellung e,wert w FROM `" & tbm(tbei) & "` WHERE einstellung LIKE 'Vordergrundfarbe%'", dbv.wCn, adOpenStatic, adLockReadOnly)
  Do While Not rs.EOF
   If rs!e = "Vordergrundfarbe 1" Then VGF1 = rs!w
   If rs!e = "Vordergrundfarbe 2" Then VGF2 = rs!w
   rs.Move 1
  Loop
  Set rs = Nothing
 Else
  Me.Vordergrundfarbe1.Visible = False
  Me.Vordergrundfarbe1.Enabled = False
  Me.Vordergrundfarbe2.Visible = False
  Me.Vordergrundfarbe2.Enabled = False
 End If
' Me.cnLab.Left = Me.aCtl.Left + Me.aCtl.Width + 50
 Me.aCtl.Left = Me.Width - Me.aCtl.Width
 Me.cnLab.Left = Me.aCtl.Left - Me.cnLab.Width
 Me.cnLab.Top = Me.aCtl.Top
 Me.cnLab.Height = Me.aCtl.Height
 Dim d1#, D2#
 d1 = Timer
 Call dbv.wCn.Execute("SET foreign_key_checks=0")
 Call dbv.wCn.Execute("DELETE FROM `" & tbm(tbar) & "` WHERE zusatz = 1  AND NOT EXISTS (SELECT * FROM `" & tbm(tbdp) & "` WHERE artnr = `" & tbm(tbar) & "`.artnr)")
 Call dbv.wCn.Execute("SET foreign_key_checks=1")
 D2 = Timer
 Debug.Print D2 - d1
 d1 = Now
 ZeiZa = Me.MFG.Height / Me.MFG.CellHeight * 0.93
 
 Call zeigdienstplan_Click
 D2 = Now
 Debug.Print "Aufbaudauer: ", D2 - d1
 
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in MDIForm_Load/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Form_Load()

Private Sub Überschrift_Click()
 If obdebug Then Debug.Print "Überschrift_Click("
End Sub ' Überschrift_Click()

Private Sub ucMDIKeys_GotFocus()
 aCtl = ucMDIKeys.name
End Sub ' ucMDIKeys_GotFocus

Private Sub ucMDIKeys_KeyPress(KeyAscii As Integer)
  Static obucMDIKeys%
    If obdebug Then Debug.Print "KeyPress", KeyAscii
    If KeyAscii = 9 Then
     If Not obucMDIKeys Then
      Call StDirekt("{RIGHT}", 0)
      obucMDIKeys = True
     End If
    End If
End Sub

Private Sub ucMDIKeys_KeyUp(KeyCode As Integer, Shift As Integer)
    If obdebug Then Debug.Print "KeyUp", KeyCode, Shift
End Sub ' ucMDIKeys_KeyUp

Private Sub ucMDIKeys_Click()
 If obdebug Then Debug.Print "ucMDIKeys_Click("
' Stop
End Sub ' ucMDIKeys_Click

Private Sub ucMDIKeys_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print "KeyDown", KeyCode, Shift
'    If KeyCode = 9 Then Stop
    Call Key(KeyCode, Shift, Me, Me.MFG, -1)
End Sub ' ucMDIKeys_KeyDown

Private Sub ucMDIKeys_LostFocus()
 If obdebug Then Debug.Print "ucMDIKeys_LostFocus", ActiveControl.name
' Stop
End Sub ' ucMDIKeys_LostFocus

Private Sub Vordergrundfarbe2_Click()
  Dim rAf&
  If obdebug Then Debug.Print "Vordergrundfarbe2_Click("
  FmCD.CmDlg.CancelError = True
'  FmCD.CmDlg.DialogTitle = "Vordergrundfarbe 2"
'  FmCD.CmDlg.Color = 65535
  On Error Resume Next
  FmCD.CmDlg.ShowColor
  If Err.Number = 0 Then
   On Error GoTo fehler
   Call dbv.wCn.Execute("UPDATE `" & tbm(tbei) & "` SET wert = " & FmCD.CmDlg.Color & " WHERE einstellung = 'Vordergrundfarbe 2'", rAf)
   If rAf = 0 Then
    Call dbv.wCn.Execute("INSERT INTO `" & tbm(tbei) & "` VALUES('Vordergrundfarbe 2'," & FmCD.CmDlg.Color & ")", rAf)
   End If
   VGF2 = FmCD.CmDlg.Color
   Call MFGRefresh(Me.MfGTyp)
  End If
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Vordergrundfarbe2_Click/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Vordergrundfarbe2_Click

Private Sub Vordergrundfarbe1_Click()
  Dim rAf&
  If obdebug Then Debug.Print "Vordergrundfarbe1_Click("
  FmCD.CmDlg.CancelError = True
'  FmCD.CmDlg.DialogTitle = "Vordergrundfarbe 2"
'  FmCD.CmDlg.Color = 65535
  On Error Resume Next
  FmCD.CmDlg.ShowColor
  If Err.Number = 0 Then
   On Error GoTo fehler
   Call dbv.wCn.Execute("UPDATE `" & tbm(tbei) & "` SET wert = " & FmCD.CmDlg.Color & " WHERE einstellung = 'Vordergrundfarbe 1'", rAf)
   If rAf = 0 Then
    Call dbv.wCn.Execute("INSERT INTO `" & tbm(tbei) & "` VALUES('Vordergrundfarbe 1'," & FmCD.CmDlg.Color & ")", rAf)
   End If
   VGF1 = FmCD.CmDlg.Color
   Call MFGRefresh(Me.MfGTyp)
  End If
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Vordergrundfarbe1_Click/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Vordergrundfarbe1_Click

Private Sub Zeilenauswahl_Click()
 Dim artnr$, i%, altCol&
 On Error GoTo fehler
 altCol = MFG.Col
 If MfGTyp = azgdp Then
  artnr = InputBox("ArtNr:")
  If LenB(artnr) <> 0 Then
   cRow(azgdp) = MFG.Row
   For i = 1 To MFG.Cols - 1
    cCol(azgdp) = i
    MFG.Col = i
    Me.Cb1.Text = artnr
    Call doChange(norefresh:=True)
   Next i
   MFG.Col = altCol
   Call MFGRefresh(azgdp)
  End If
 End If
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Zeilenauswahl_Click/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Zeilenauswahl_Click()

Private Sub DatenbankErstellen_Click()
 Dim rAf&, erg&, i%
 Const opti = 1 + 2 + 8 ' 32 macht die Auswahl bei PatAuswahl sehr langsam
 Dim Server$, User$, pwd$, sql$, treiber$, db$, Benutzer$
 On Error GoTo fehler
 Server = InputBox("Server für die Erstellung der Tabellen:")
 User = InputBox("Benutzer für die Erstellung der Tabellen auf " & Server)
 pwd = InputBox("Passwort für " & User & " auf " & Server & ":")
 treiber = InputBox("Treiber für die Tabellenerstellung:", , "MySQL ODBC 5.1 Driver") ' MySQL ODBC 5.1 Driver
 db = InputBox("Datenbank für die Tabellenerstellung:", , "dp")
 Benutzer = InputBox("Benutzer für Tabellenverwendung:")
' Set cn = Nothing
' dbv.wCn.Open "DRIVER={" & treiber & "};server=" & server & ";uid=" & user & ";pwd=" & pwd & ";option=" & opti
 On Error Resume Next
 dbv.wCn.Execute ("use `" & db & "`")
 If Err.Number <> 0 Then
  dbv.wCn.Execute ("CREATE DATABASE `" & db & "`")
  dbv.wCn.Execute ("use `" & db & "`")
 End If
 On Error GoTo fehler
' Exit Function
 Call dbv.wCn.Execute("SET foreign_key_checks=0")
 
 Dim k As TbTyp
 For k = MTBeg To tbende - 1
  dbv.wCn.Execute "SHOW TABLES LIKE '" & tbm(k) & "'", rAf
  If rAf = 1 Then
   erg = MsgBox("Tabelle '" & tbm(k) & "' existiert schon. Löschen?", vbYesNoCancel)
   Select Case erg
    Case vbYes
     Call dbv.wCn.Execute("DROP TABLE IF EXISTS `" & tbm(k) & "`")
     rAf = 0
    Case vbCancel: Exit Sub
   End Select
  End If
  If rAf = 0 Then
   Select Case k
    Case tbar
     sql = "CREATE TABLE  `" & tbm(tbar) & "` (" & _
           "`ArtNr` varchar(30) COLLATE latin1_german2_ci COMMENT 'Artnr'," & _
           "`erkl` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Erklärung'," & _
           "`Stdn` double(3,1) unsigned NOT NULL COMMENT 'Stunden'," & _
           "`Farbe` int(4) unsigned NOT NULL," & _
           "`zusatz` tinyint(1) unsigned DEFAULT NULL," & _
           "PRIMARY KEY (`ArtNr`)" & _
           ") ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Dienstarten'"
    Case tbwp
     sql = "CREATE TABLE  `" & tbm(tbwp) & "` (" & _
           "`PersNr` int(5) unsigned NOT NULL DEFAULT '0' COMMENT 'Personal-Nr.'," & _
           "`ab` date NOT NULL DEFAULT '2007-01-01' COMMENT 'Gültigkeitsbeginn'," & _
           "`Mo` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Montag'," & _
           "`Di` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Dienstag'," & _
           "`Mi` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Mittwoch'," & _
           "`Do` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Donnerstag'," & _
           "`Fr` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Freitag'," & _
           "`Sa` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Samstag'," & _
           "`So` varchar(30) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Sonntag'," & _
           "`WAZ` double(3,1) NOT NULL DEFAULT '38.5'," & _
           "`Urlaub` int(2) NOT NULL DEFAULT '28' COMMENT 'Urlaubstage pro Jahr'," & _
           "PRIMARY KEY (`PersNr`,`ab`)," & _
           "KEY `Mo` (`Mo`)," & _
           "KEY `Di` (`Di`)," & _
           "KEY `Mi` (`Mi`)," & _
           "KEY `Do` (`Do`)," & _
           "KEY `Fr` (`Fr`)," & _
           "KEY `Sa` (`Sa`)," & _
           "KEY `So` (`So`),"
     sql = sql & _
           " CONSTRAINT `DiArt` FOREIGN KEY (`Di`) REFERENCES `" & tbm(tbar) & "` (`ArtNr`)," & _
           " CONSTRAINT `DoArt` FOREIGN KEY (`Do`) REFERENCES ``" & tbm(tbar) & "`` (`ArtNr`)," & _
           " CONSTRAINT `FrArt` FOREIGN KEY (`Fr`) REFERENCES `" & tbm(tbar) & "` (`ArtNr`)," & _
           " CONSTRAINT `MiArt` FOREIGN KEY (`Mi`) REFERENCES `" & tbm(tbar) & "` (`ArtNr`)," & _
           " CONSTRAINT `MoArt` FOREIGN KEY (`Mo`) REFERENCES `" & tbm(tbar) & "` (`ArtNr`)," & _
           " CONSTRAINT `Persnr1` FOREIGN KEY (`PersNr`) REFERENCES `" & tbm(tbma) & "` (`PersNr`)," & _
           " CONSTRAINT `SaArt` FOREIGN KEY (`Sa`) REFERENCES `" & tbm(tbar) & "` (`ArtNr`)," & _
           " CONSTRAINT `SoArt` FOREIGN KEY (`So`) REFERENCES `" & tbm(tbar) & "` (`ArtNr`)" & _
           ") ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Wochenplan'"
           
    Case tbpr
     sql = "CREATE TABLE `" & tbm(tbpr) & "` (" & _
 " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT," & _
 " `tag` date NOT NULL," & _
 " `PersNr` int(5) unsigned NOT NULL," & _
 " `ArtNrV` varchar(30) COLLATE latin1_german2_ci," & _
 " `ArtNr` varchar(30) COLLATE latin1_german2_ci," & _
 " `AendDat` datetime DEFAULT NULL," & _
 " `AendPC` varchar(20) COLLATE latin1_german2_ci DEFAULT NULL," & _
 " `AendUser` varchar(25) COLLATE latin1_german2_ci DEFAULT NULL," & _
 " PRIMARY KEY (`ID`)," & _
 " KEY `Persnr` (`PersNr`)," & _
 " KEY `Artnr` (`ArtNr`)," & _
 " CONSTRAINT `ProtArtNr` FOREIGN KEY (`ArtNr`) REFERENCES `" & tbm(tbar) & "` (`ArtNr`)," & _
 " CONSTRAINT `ProtPersnr` FOREIGN KEY (`PersNr`) REFERENCES `" & tbm(tbma) & "` (`PersNr`)" & _
") ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Änderungsprotokoll für Dienstplan';"
'") ENGINE=InnoDB AUTO_INCREMENT=86 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Änderungsprotokoll für Dienstplan'"

    Case tbdp
     sql = "CREATE TABLE  `" & tbm(tbdp) & "` (" & _
  "`ID` int(10) unsigned NOT NULL AUTO_INCREMENT," & _
  "`tag` date NOT NULL," & _
  "`PersNr` int(5) unsigned NOT NULL," & _
  "`ArtNr` varchar(30) COLLATE latin1_german2_ci," & _
  "PRIMARY KEY (`ID`)," & _
  "KEY `Persnr` (`PersNr`)," & _
  "KEY `Artnr` (`ArtNr`)," & _
  "CONSTRAINT `ArtNr` FOREIGN KEY (`ArtNr`) REFERENCES `" & tbm(tbar) & "` (`ArtNr`)," & _
 " CONSTRAINT `Persnr` FOREIGN KEY (`PersNr`) REFERENCES `" & tbm(tbma) & "` (`PersNr`)" & _
") ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Dienstplan';"
'") ENGINE=InnoDB AUTO_INCREMENT=79 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci COMMENT='Dienstplan'"

 Case tbma
  sql = "CREATE TABLE  `" & tbm(tbma) & "` (" & _
  "`PersNr` int(5) unsigned NOT NULL AUTO_INCREMENT COMMENT 'Personal-Nummer'," & _
  "`Kuerzel` varchar(10) COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Kürzel'," & _
  "`Nachname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL," & _
  "`Vorname` varchar(50) COLLATE latin1_german2_ci DEFAULT NULL," & _
  "`Aus` date DEFAULT NULL COMMENT 'Austritt'," & _
 " PRIMARY KEY (`PersNr`)," & _
 " KEY `Aus` (`Aus`)" & _
") ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci;"
'") ENGINE=InnoDB AUTO_INCREMENT=60 DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci;"
 Case tbul
  sql = "CREATE TABLE  `dp`.`" & tbm(tbul) & "` (" & _
  " `user` varchar(45) CHARACTER SET latin1 NOT NULL," & _
  " `Passwort` blob NOT NULL," & _
  " `hinzugefügt` datetime NOT NULL," & _
  " `geändert` datetime NOT NULL," & _
  " `ID` int(10) unsigned NOT NULL AUTO_INCREMENT," & _
  " PRIMARY KEY (`ID`)," & _
  " UNIQUE KEY `user` (`user`)" & _
" ) ENGINE=InnoDB AUTO_INCREMENT=11 DEFAULT CHARSET=latin1 COLLATE=latin1_german1_ci COMMENT='Benutzer für Datenänderungen'"
 Case tbbi
  sql = "CREATE TABLE `" & tbm(tbbi) & "` (" & _
 " `PersNr` int(5) unsigned NOT NULL DEFAULT '0'," & _
 " `Jahr` int(4) unsigned NOT NULL," & _
 " `Urlaub` double(4,1) NOT NULL COMMENT 'Tage'," & _
 " `Überstunden` double(4,1) NOT NULL COMMENT 'Stunden'," & _
 " `Fortbildung` double(4,1) NOT NULL COMMENT 'Tage'," & _
 " PRIMARY KEY (`PersNr`,`Jahr`) USING BTREE," & _
 " CONSTRAINT `PersNr2` FOREIGN KEY (`PersNr`) REFERENCES `" & tbm(tbma) & "` (`PersNr`)" & _
") ENGINE=InnoDB DEFAULT CHARSET=latin1 COMMENT='Urlaubs- und Überstundenbilanzen';"

   End Select
   Call dbv.wCn.Execute(sql)
  End If
 Next k

' Call dbv.wCn.Execute("grant all ON * to '" & user & "' WITH GRANT OPTION")
 On Error Resume Next
 Call dbv.wCn.Execute("INSERT INTO mysql.db VALUES('%','" & db & "','" & User & "','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y')")
 On Error GoTo fehler
 Call dbv.wCn.Execute("FLUSH PRIVILEGES")
 Call dbv.wCn.Execute("GRANT ALL ON * to '" & Benutzer & "' WITH GRANT OPTION")
 Call dbv.wCn.Execute("SET foreign_key_checks=1")
 Call Grundausstatt(nloe:=True)
 Call ViewsErstellen
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Create/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' DatenbankErstellen_Click

Public Function Grundausstatt(Optional nloe% = 0)
 On Error GoTo fehler
 If nloe Then
  Set dbv.wCn = Nothing
  dbv.wCn.Open dbv.CnStr
  Set DBCn = dbv.wCn
  DBCnS = dbv.CnStr
 End If
 dbv.wCn.Execute "SET foreign_key_checks = 0"
 dbv.wCn.Execute "DELETE FROM `" & tbm(tbar) & "` WHERE artnr in ('','-','0,5','1','1,5','2','2,5','3','3,5','4','4,5','5','5,5','6','6,5','7','7,5','8','8,5','9','9,5','10','10,5','11','11,5','12','12,5','13','13,5','14','14,5','15','15,5','ü','u','uw','su','f','fw','k','ki','b','hFT','g','WF')"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('','keine Änderung',0,65535)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('-','nicht eingeplant',0,14671839)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('1','anwesend 1 Stunde',1,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('2','anwesend 2 Stunden',2,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('3','anwesend 3 Stunden',3,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('4','anwesend 4 Stunden',4,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('5','anwesend 5 Stunden',5,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('6','anwesend 6 Stunden',6,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('7','anwesend 7 Stunden',7,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('8','anwesend 8 Stunden',8,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('9','anwesend 9 Stunden',9,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('10','anwesend 10 Stunden',10,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('11','anwesend 11 Stunden',11,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('12','anwesend 12 Stunden',12,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('13','anwesend 13 Stunden',13,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('14','anwesend 14 Stunden',14,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('15','anwesend 15 Stunden',14,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('0,5','anwesend 0,5 Stunden',1.5,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('1,5','anwesend 1,5 Stunden',1.5,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('2,5','anwesend 2,5 Stunden',2.5,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('3,5','anwesend 3,5 Stunden',3.5,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('4,5','anwesend 4,5 Stunden',4.5,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('5,5','anwesend 5,5 Stunden',5.5,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('6,5','anwesend 6,5 Stunden',6.5,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('7,5','anwesend 7,5 Stunden',7.5,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('8,5','anwesend 8,5 Stunden',8.5,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('9,5','anwesend 9,5 Stunden',9.5,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('10,5','anwesend 10,5 Stunden',10.5,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('11,5','anwesend 11,5 Stunden',11.5,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('12,5','anwesend 12,5 Stunden',12.5,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('13,5','anwesend 13,5 Stunden',13.5,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('14,5','anwesend 14,5 Stunden',14.5,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('15,5','anwesend 15,5 Stunden',15.5,16629431)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('b','Betriebsausflug',8,2422722)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('f','Fortbildung',0,11163042)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('fw','Fortbildungswunsch',0,14614751)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('g','geschlossen',0,10524329)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('hFT','Halbfeiertag',0,10461087)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('k','krank',0,8421376)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('ki','Kind krank',0,8421376)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('su','Sonderurlaub',0,8945663)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('u','Urlaub',0,255)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('ü','Überstunden',0,2834420)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('üw','Überstundenausgleichswunsch',0,4870879)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('uw','Urlaubswunsch',0,7303167)"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbar) & "`(`ArtNr`,`erkl`,`Stdn`,`Farbe`) values ('WF','Wochenend/Feiertag',0,7303023)"
 If nloe Then
 dbv.wCn.Execute "DELETE FROM `" & tbm(tbwp) & "` WHERE ab < '2007-10-30'"
 dbv.wCn.Execute "INSERT INTO `" & tbm(tbwp) & "` (`PersNr`,`ab`,`Mo`,`Di`,`Mi`,`Do`,`Fr`,`Sa`,`So`,`WAZ`,`Urlaub`) VALUES " & _
 "(43,'2006-05-01','8','5,5','5,5','-','-','WF','WF',0.0,30)," & _
 "(43,'2006-07-01','10,5','11','6','11','ü','WF','WF',0.0,30)," & _
 "(45,'2006-02-01','-','9,5','5,5','9,5','5,5','WF','WF',30.0,30)," & _
 "(45,'2006-04-01','-','9,5','6','9,5','-','WF','WF',25.0,30)," & _
 "(45,'2006-11-01','-','-','5,5','9,5','-','WF','WF',15.0,30)," & _
 "(45,'2007-06-01','5,5','5,5','6','9','-','WF','WF',26.0,30)," & _
 "(46,'2005-09-19','9','5,5','6','-','6','WF','WF',26.5,26)," & _
 "(46,'2006-09-11','9','9,5','6','-','6','WF','WF',30.5,26)," & _
 "(55,'2006-10-30','5','5','-','5','5','WF','WF',20.0,26)," & _
 "(55,'2007-01-15','9','9,5','-','9,5','6','WF','WF',34.0,26)," & _
 "(59,'2007-07-01','9','9','5,5','9','5,5','WF','WF',38.0,0)," & _
 "(59,'2007-10-29','9','-','5,5','9','5,5','WF','WF',29.0,0);"
 dbv.wCn.Execute "DELETE FROM `" & tbm(tbma) & "`"
dbv.wCn.Execute "INSERT INTO `" & tbm(tbma) & "` (`PersNr`,`Kuerzel`,`Nachname`,`Vorname`,`Aus`) VALUES " & _
 "(43,'wr','Roßmeier','Walburga',NULL)," & _
 "(45,'cr','Reindl','Cornelia',NULL)," & _
 "(46,'tst','Sturm','Tamara',NULL)," & _
 "(55,'ke','Kupinic','Elvisa',NULL)," & _
 "(59,'bga','Gallitz','Benedikt','2008-06-30');"
 dbv.wCn.Execute "SET foreign_key_checks = 1"
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Grundausstatt/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Grundausstatt


Function DtbCreateQueryDef$(QName$, sql$)
 Dim csql As New CString
 Const ifexists$ = "IF EXISTS"
 If sql <> vNS Then
 csql.Append sql
  
  On Error GoTo fehler
  Call myEFrag("DROP TABLE " & ifexists & " `" & QName & "`;")
  Call myEFrag("DROP VIEW " & ifexists & " `" & QName & "`;")
  On Error GoTo fehler
  Dim cvrs As Recordset
  
  Call myEFrag("CREATE OR REPLACE ALGORITHM=UNDEFINED DEFINER=`" & Forms(0).dbv.uid & "`@`%` SQL SECURITY DEFINER VIEW `" & QName & "` AS " & csql)
  Set cvrs = myEFrag("SHOW TABLES WHERE `tables_in_" & DefDB(DBCn) & "` LIKE '" & QName & "'")
  If cvrs.BOF Then
   dbv.ausgeb QName & " konnte nicht erstellt werden.", True
  Else
'   Debug.Print QName & " gibts."
  End If
 
 DtbCreateQueryDef = csql
 End If ' sql = vns
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
If InStrB(Err.Description, "gone away") <> 0 Then
 Call DBCnOpen
 Resume
End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DtbCreateQueryDef/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DtbCreateQueryDef


Private Sub MFG_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim sql$, tt$, rs As ADODB.Recordset, rsd As New ADODB.Recordset
Static altMCol&, altMRow&
'Dim c, r As Long, aktRow&, aktCol&
'Dim H1, H2, W1, W2 As Long
'Dim MyRowName, MyColName  As String
'Dim rs As New ADODB.Recordset
'For r = 0 To MFG.Rows - 1
' H1 = MFG.RowPos(r)
' H2 = MFG.RowHeight(r)
'
' If y >= H1 AND y <= (H1 + H2) Then
'  MyRowName = " Zeile: " & r
'  aktRow = r
'  If aktRow > 159 AND aktRow < 162 Then
'    Debug.Print "aktrow=" & aktRow & ",y=" & y & ",mfg.rowpos()=" & H1 & ",mfg.rowheight()=" & H2
'  End If
'  Exit For
' End If
'Next r
'For c = 0 To MFG.Cols - 1
' W1 = MFG.ColPos(c)
' W2 = MFG.ColWidth(c)
'
' If x >= W1 AND x <= (W1 + W2) Then
'  MyColName = " Spalte: " & c
'  aktCol = c
'  Exit For
' End If
'Next c
Dim i%

dp:
On Error GoTo fehler
With MFG
If .MouseCol <> altMCol Or .MouseRow <> altMRow Then
 altMCol = .MouseCol
 altMRow = .MouseRow
Call vormerken(MfGTyp)
.ToolTipText = "Position: Spalte:" & .MouseCol & ", Zeile:" & .MouseRow ' MyColName & MyRowName
Select Case MfGTyp
 Case azgdp
  If .MouseCol = 0 Then
   If LenB(User) <> 0 Then
'    rs.Open "SELECT kaldb, kaltab, kaldatsp FROM `" & tbm(tbma) & "` ma WHERE nachname LIKE '" & User & "'", dbv.wCn, adOpenStatic, adLockReadOnly
    myFrag rs, "SELECT kaldb, kaltab, kaldatsp FROM `" & tbm(tbma) & "` ma WHERE nachname LIKE '" & User & "'", adOpenStatic, dbv.wCn, adLockReadOnly
    If Not rs.BOF Then
     If IsDate(BegD) Then
      If Not IsNull(rs!kaldb) And Not IsNull(rs!kaltab) And Not IsNull(rs!kaldatsp) Then
       If LenB(rs!kaldb) <> 0 And LenB(rs!kaltab) <> 0 And LenB(rs!kaldatsp) <> 0 Then ' 9.3.09
        sql = "SELECT * FROM `" & rs!kaldb & "`.`" & rs!kaltab & "` WHERE `" & rs!kaldatsp & "` = " & datform(BegD + .MouseRow - 1 - SZZ)
'        rsd.Open sql, dbv.wCn, adOpenStatic, adLockReadOnly
        myFrag rsd, sql, adOpenStatic, dbv.wCn, adLockReadOnly
        If Not rsd.EOF Then
         On Error GoTo f0
         If rsd.BOF Then GoTo f0
         tt = rsd(rs.Fields("kaldatsp").Value) & ": "
         On Error GoTo fehler
'         For i = 1 To 4
'          If Not IsNull(rs.Fields("kalsp" & i)) Then
'           tt = tt & rsd(rs.Fields("kalsp" & i).Value)
'          End If
'         Next i
         .ToolTipText = tt
  '      .ToolTipText = rs!Datum & ": " & rs!geburtstage & "; " & rs!termin & "; " & rs!hg
        End If ' If rs!kaldb <> vns AND rs!kaltab <> vns AND rs!kaldatsp <> vns
       End If ' Not IsNull(rs!kaldb) AND NOT IsNull(rs!kaltab) AND NOT IsNull(rs!kaldatsp)
      End If ' Not IsNull(rs!kaldb) And Not IsNull(rs!kaltab) And Not IsNull(rs!kaldatsp) Then
     End If ' IsDate(BegD) Then
    End If ' Not rs.BOF Then
   End If ' LenB(User) <> 0 Then
  ElseIf .MouseCol - 1 >= 0 And .MouseCol - 1 <= UBound(pn) Then
   If .MouseRow = 0 Then
    .ToolTipText = "Persnr. " & pn(.MouseCol - 1)
   ElseIf .MouseRow = 2 Then
    sql = "SELECT * FROM `" & tbm(tbbi) & "` WHERE persnr = " & pn(.MouseCol - 1) & " AND jahr = " & Me.Jahr
    Set rs = Nothing
'    rs.Open sql, dbv.wCn, adOpenStatic, adLockReadOnly
    myFrag rs, sql, adOpenStatic, dbv.wCn, adLockReadOnly
    If rs.BOF Then
    Else
     .ToolTipText = "Planstunden: " & rs!planstunden
    End If
   ElseIf .MouseRow = 4 Then ' Urlaub
'    .ToolTipText = dbv.wCn.Execute("SELECT group_concat(case DATE_FORMAT(tag,'%w') when 0 then 'So' when 1 then 'Mo' when 2 then 'Di' when 3 then 'Mi' when 4 then 'Do' when 5 then 'Fr' when 6 then 'Sa' else 'un' end,DATE_FORMAT(tag,', %d.%m.%y '),artnr,'\\\n') FROM `" & tbm(tbdp) & "` WHERE persnr=" & pn(.MouseCol - 1) & " and ArtNr in ('u','uw') and YEAR(tag)=" & Me.Jahr).Fields(0)
'''    dbv.wCn.Close
'''    dbv.wCn.Open
'    .ToolTipText = dbv.wCn.Execute("SELECT COALESCE(GROUP_CONCAT(DATE_FORMAT(tag,'%d.%m') separator ','),'') FROM `" & tbm(tbdp) & "` WHERE persnr=" & pn(.MouseCol - 1) & " and ArtNr in ('g','u','uw') and YEAR(tag)=" & Me.Jahr).Fields(0)
    Dim rstag As ADODB.Recordset
    myFrag rstag, "SELECT COALESCE(GROUP_CONCAT(DATE_FORMAT(tag,'%d.%m') separator ','),'') FROM `" & tbm(tbdp) & "` WHERE persnr=" & pn(.MouseCol - 1) & " and ArtNr in ('g','u','uw') and YEAR(tag)=" & Me.Jahr, adOpenStatic, dbv.wCn, adLockReadOnly
    If rstag.State <> 0 Then If Not rstag.BOF Then .ToolTipText = rstag.Fields(0)
   ElseIf .MouseRow > .FixedRows Then
    tt = vNS
    Dim Reihe&
    Reihe = .MouseRow
    sql = "SELECT * FROM `" & tbm(tbpr) & "` WHERE `tag` = " & datform(BegD + Reihe - 1 - SZZ) & " AND `PersNr` = " & pn(.MouseCol - 1)
    Set rs = Nothing
'    dbv.wCn.Close
'    dbv.wCn.Open
'    rs.Open sql, dbv.wCn, adOpenStatic, adLockReadOnly
    myFrag rs, sql, adOpenStatic, dbv.wCn, adLockReadOnly
    If rs.BOF Then
     Set rs = Nothing
     Dim mcol&
     mcol = .MouseCol - 1
     If mcol < LBound(pn) Then mcol = LBound(pn)
     If mcol > UBound(pn) Then mcol = UBound(pn)
'     rs.Open "SELECT vorname, nachname FROM `" & tbm(tbma) & "` WHERE persnr = " & pn(mcol), dbv.wCn, adOpenStatic, adLockReadOnly
     Set rs = Nothing
     myFrag rs, "SELECT vorname, nachname FROM `" & tbm(tbma) & "` WHERE persnr = " & pn(mcol), adOpenStatic, dbv.wCn, adLockReadOnly
     If Not rs.EOF Then
      tt = "Keine Einträge zu " & BegD + Reihe - 1 - SZZ & " zu " & rs!Vorname & " " & rs!Nachname
     End If
    Else
     Do While Not rs.EOF
      tt = tt & rs!ArtNrV & "->" & rs!artnr & "(" & rs!aenddat & "/" & rs!User & "), "
      rs.Move 1
     Loop
    End If
    .ToolTipText = tt
   End If ' .MouseRow = 2 Then elsif
  End If ' .MouseCol
 Case Else
End Select ' select case mfgtyp
End If ' .MouseCol <> altMCol Or .MouseRow <> altMRow Then
End With
Exit Sub
f0:
Set rsd = Nothing
Set rs = Nothing
GoTo dp
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MouseMove/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub 'MouseMove

