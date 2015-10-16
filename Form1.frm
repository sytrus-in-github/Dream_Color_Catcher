VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form UI_Image 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Image:"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6735
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin MSComCtl2.FlatScrollBar fsb_H 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   450
      _Version        =   393216
      MousePointer    =   9
      Appearance      =   2
      Arrows          =   65536
      LargeChange     =   16
      Max             =   255
      Orientation     =   1572865
      SmallChange     =   16
   End
   Begin MSComCtl2.FlatScrollBar fsb_V 
      Height          =   4455
      Left            =   6480
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   7858
      _Version        =   393216
      MousePointer    =   7
      Appearance      =   2
      LargeChange     =   16
      Max             =   255
      Orientation     =   1572864
      SmallChange     =   16
   End
   Begin VB.Frame lbl_Scroll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   3
      Top             =   4440
      Width           =   255
   End
   Begin VB.PictureBox Img 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   4455
      ScaleWidth      =   6495
      TabIndex        =   2
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "UI_Image"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pic As IPictureDisp
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long


Private Sub Form_Load()
    Err.Clear
    On Error Resume Next
    Me.Img.Picture = pic
    If Err.Number <> 0 Then
        Me.Img.Picture = LoadPicture()
        Err.Clear
    Else
        refreshUI
    End If
End Sub

Private Sub refreshUI()
    If Img.Width > Me.Width - 240 Or Img.Height > Me.Height - 540 Then
        If Me.Width < 750 Then Me.Width = 750
        If Me.Height < 1050 Then Me.Height = 1050
        With fsb_H
            .Visible = True
            .Top = Me.Height - 795
            .Width = Me.Width - 495
            .Min = .Width \ Screen.TwipsPerPixelX
            .Max = Img.Width \ Screen.TwipsPerPixelX
            If .Max <= .Min Then .Max = .Min + 1
            If .Min * 2 < .Max Then .LargeChange = .Min Else .LargeChange = .Max - .Min
        End With
        With fsb_V
            .Visible = True
            .Left = Me.Width - 495
            .Height = Me.Height - 795
             .Min = .Height \ Screen.TwipsPerPixelY
            .Max = Img.Height \ Screen.TwipsPerPixelY
            If .Max <= .Min Then .Max = .Min + 1
            If .Min * 2 < .Max Then .LargeChange = .Min Else .LargeChange = .Max - .Min
        End With
        With Img
            If .Left + .Width < fsb_H.Width Then
                If fsb_H.Width < .Width Then .Left = fsb_H.Width - .Width: fsb_H.Value = fsb_H.Min - (.Left \ Screen.TwipsPerPixelX) Else .Left = 0: fsb_H.Value = fsb_H.Min
            End If
            If .Top + .Height < fsb_V.Height Then
                If fsb_V.Height < .Height Then .Top = fsb_V.Height - .Height: fsb_V.Value = fsb_V.Min - (.Top \ Screen.TwipsPerPixelY) Else .Top = 0: fsb_V.Value = fsb_V.Min
            End If
        End With
        lbl_Scroll.Visible = True
        lbl_Scroll.Top = fsb_H.Top
        lbl_Scroll.Left = fsb_V.Left
    Else
        fsb_H.Visible = False
        fsb_V.Visible = False
        lbl_Scroll.Visible = False
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then UI_Main.Show
End Sub

Private Sub Form_Resize()
    refreshUI
End Sub

Private Sub fsb_H_Change()
    Img.Left = (fsb_H.Min - fsb_H.Value) * Screen.TwipsPerPixelX
End Sub

Private Sub fsb_V_Change()
    Img.Top = (fsb_V.Min - fsb_V.Value) * Screen.TwipsPerPixelY
End Sub

Private Sub Img_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim c As Long
    c = GetPixel(Img.hdc, X \ Screen.TwipsPerPixelX, Y \ Screen.TwipsPerPixelY)
    lbl_Scroll.BackColor = c
    Me.Caption = Right("00000" + Hex(c), 6)
End Sub

Private Sub Img_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim c As Long
    c = lbl_Scroll.BackColor
    UI_Main.intR = c And 255
    UI_Main.intG = (c And 65280) \ 256
    UI_Main.intB = (c And 16711680) \ 65536
    UI_Main.UpdateDisplay
    Unload Me
    UI_Main.Show
End Sub
