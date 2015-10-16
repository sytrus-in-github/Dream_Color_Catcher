VERSION 5.00
Begin VB.Form UI_Screen 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3540
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   5085
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3540
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Frame frm_Display 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   3000
      TabIndex        =   0
      Top             =   960
      Width           =   2055
      Begin VB.Label lbl_Explanation 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Click to take the color   Press [Esc] to return"
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label lbl_Color 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FFFFFF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.PictureBox Img 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   1575
      ScaleWidth      =   2655
      TabIndex        =   3
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "UI_Screen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TODO Magnifier 15x15 size 8
Option Explicit
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Sub Form_Unload(Cancel As Integer)
        UI_Main.WindowState = vbNormal
End Sub

'to avoid mouse selection blockage
Private Sub frm_Display_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frm_Display.Left = Me.Width - 2055 - frm_Display.Left
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me: UI_Main.Show
End Sub

Private Sub Img_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me: UI_Main.Show
End Sub

Private Sub Form_Load()
    Dim scr_hWnd As Long, scr_hDC As Long
    'get window handle and DC handle
    scr_hWnd = GetDesktopWindow
    scr_hDC = GetDC(scr_hWnd)
    'resize UI form to full screen size
    Me.Height = Screen.Height
    Me.Width = Screen.Width
    frm_Display.Left = Me.Width - 2055
    frm_Display.Top = Me.Height - 2535
    'take a screen shot to imagebox
    BitBlt Me.hdc, 0, 0, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY, GetDC(0), 0, 0, vbSrcCopy
    ReleaseDC scr_hWnd, scr_hDC
    Me.Img.Picture = Me.Image
End Sub

Private Sub Img_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim c As Long
    c = GetPixel(Img.hdc, X \ Screen.TwipsPerPixelX, Y \ Screen.TwipsPerPixelY)
    lbl_Color.BackColor = c
    lbl_Color.Caption = Right("00000" + Hex(c), 6)
    If (c And 255) + (c And 65280) \ 256 + (c And 16711680) \ 65536 > 381 Then lbl_Color.ForeColor = 0 Else lbl_Color.ForeColor = 16777215
End Sub

Private Sub Img_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim c As Long
    c = lbl_Color.BackColor
    UI_Main.intR = c And 255
    UI_Main.intG = (c And 65280) \ 256
    UI_Main.intB = (c And 16711680) \ 65536
    UI_Main.UpdateDisplay
    Unload Me
    UI_Main.Show
End Sub
