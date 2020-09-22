VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "BalPoint SimpleSkin"
   ClientHeight    =   5310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4395
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMain.frx":0442
   MousePointer    =   99  'Custom
   ScaleHeight     =   5310
   ScaleWidth      =   4395
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMainSkin 
      Height          =   5130
      Left            =   0
      Picture         =   "frmMain.frx":110C
      ScaleHeight     =   5070
      ScaleWidth      =   4140
      TabIndex        =   0
      Top             =   0
      Width           =   4200
      Begin VB.Timer tmOnTop 
         Interval        =   1
         Left            =   75
         Top             =   75
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BalPoint Skin
'Created by R.Baldewsingh 2002
'--------------------------------------------
'Run this program and press F1 for more Info.
'--------------------------------------------
Option Explicit
'Always on top Contsants
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Dim result As Long
    result = SetWindowPos(frmMain.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    End
End If
End Sub

Private Sub Form_Load()
    Dim WindowRegion As Long
    ' I set all these settings here so you won't forget
    ' them and have a non-working demo... Set them in
    ' design time
    picMainSkin.ScaleMode = vbPixels
    picMainSkin.AutoRedraw = True
    picMainSkin.AutoSize = True
    picMainSkin.BorderStyle = vbBSNone
    Me.BorderStyle = vbBSNone
    'Set picMainSkin.Picture = LoadPicture(App.Path & "\Graphics\gifMain.gif") or ...
    Set picMainSkin.Picture = picMainSkin.Picture
    Me.Width = picMainSkin.Width
    Me.Height = picMainSkin.Height
    WindowRegion = MakeRegion(picMainSkin)
    SetWindowRgn Me.hwnd, WindowRegion, True
End Sub
Private Sub picMainSkin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
        Dim result As Long
        result = SetWindowPos(frmMain.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
        End
    Else
If KeyCode = vbKeyF1 Then
        Dim wrap$
        wrap$ = Chr$(10) + Chr$(13)
        MsgBox "This is a real working form!" & wrap$ & wrap$ & "This form displays a simple way to create skinned applications." & wrap$ & "It also shows how to make a form movable with just a simple line of code," & wrap$ & "but that's not all. It also shows how to always keep your form ontop." & wrap$ & wrap$ & "Intructions:" & wrap$ & "1) Drag the form to move it." & wrap$ & "2) To see this msgbox press the F1 key." & wrap$ & "3) Press the Escape key to Exit." & wrap$ & wrap$ & "Contact me: mrfloat@msn.com", vbOKOnly + vbInformation, "Information"
End If
End If
End Sub
Private Sub picMainSkin_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      ' Pass the handling of the mouse down message to
      ' the (non-existing really) form caption, so that
      ' the form itself will be dragged when the picture is dragged.
      '
      ' If you have Win 98, Make sure that the "Show window
      ' contents while dragging" display setting is on for nice results.
      ReleaseCapture
      SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
Private Sub tmOnTop_Timer()
Dim result As Long
result = SetWindowPos(frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

' Examples:
' When Exiting a form:
' Dim result As Long
' result = SetWindowPos(frmMain.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
' End
'
' When disabling always on top:
' tmOntop.Enabled = False
' Dim result As Long
' result = SetWindowPos(frmMain.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
