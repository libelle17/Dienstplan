VERSION 5.00
Begin VB.UserControl ucMDIKeys 
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1560
   ScaleWidth      =   1800
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   135
      Top             =   450
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "MDIKeys"
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   645
   End
End
Attribute VB_Name = "ucMDIKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *** Den Artikel zu diesem Modul finden Sie unter http://www.aboutvb.de/kom/artikel/kommdikeys.htm ***

Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Sub tmr_Timer()
    Const SW_NORMAL = 1
    
    With UserControl
        If GetForegroundWindow() = .Parent.hwnd Then
            If GetFocus <> .hwnd Then
                If .Parent.ActiveForm Is Nothing Then
                    ShowWindow .hwnd, SW_NORMAL
                    SetParent .hwnd, .Parent.hwnd
                    SetFocusAPI .hwnd
                End If
            End If
        End If
    End With
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Ambient.UserMode Then
        UserControl.BackStyle = 0
        lbl.Visible = False
        tmr.Enabled = True
    End If
End Sub

Private Sub UserControl_Resize()
    If Not Ambient.UserMode Then
        With lbl
            UserControl.Size .Width + 2 * .Left, .Height + 2 * .Top
        End With
    End If
End Sub

