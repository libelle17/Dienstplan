VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Anmeldung"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Benutzername:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Kennwort:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public obdefinier%
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'Globale Variable auf False setzen,
    'um eine fehlgeschlagene Anmeldung zu kennzeichnen.
    LoginSucceeded = False
    Me.Hide
End Sub ' cmdCancel_Click()

Private Sub cmdOK_Click()
  Dim rs As New ADODB.Recordset
  On Error GoTo fehler
  If Not obdefinier Then
    'Auf korrektes Kennwort überprüfen
    MDI.dbv.wCn.Close
    MDI.dbv.wCn.Open
    rs.Open "SELECT aes_decrypt(passwort,'0&F54') = '" & Me.txtPassword & "' as obrichtig,u.* FROM `" & tbm(tbul) & "` u WHERE user = '" & Me.txtUserName & "'", MDI.dbv.wCn, adOpenStatic, adLockReadOnly
    If rs.BOF Or IsNull(rs!obrichtig) Then
        MsgBox "Ungültiger Benutzer. Bitte versuchen Sie es noch einmal!", , "Anmeldung"
        Me.txtUserName.SetFocus
        SendKeys "{Home}+{End}"
    Else
     If rs!obrichtig Then
        'Geben Sie hier Code ein, um den Erfolg
        'an die aufrufende Unterroutine weiterzuleiten.
        'Setzen der globalen Variablen ist leicht.
        LoginSucceeded = True
        Me.Hide
     Else
        MsgBox "Ungültiges Kennwort. Bitte versuchen Sie es noch einmal!", , "Anmeldung"
        txtPassword.SetFocus
'        SendKeys "{Home}+{End}"
     End If
    End If
  Else
   Me.Hide
  End If
  Exit Sub
fehler:
  Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in cmdOK_Click/" + App.Path)
   Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
   Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
   Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
  End Select
End Sub ' Sub cmdOK_Click()

Private Sub Form_Activate()
 On Error GoTo fehler
 If Me.txtUserName <> vNS Then
  Me.txtPassword.SetFocus
 Else
  Me.txtUserName.SetFocus
 End If
  Exit Sub
fehler:
  Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Form_Activate/" + App.Path)
   Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
   Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
   Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
  End Select
End Sub ' Form_Activate()

