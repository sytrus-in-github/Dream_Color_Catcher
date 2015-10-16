VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form UI_Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dream Color Catcher"
   ClientHeight    =   4695
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   6975
   Icon            =   "UI_Main.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6975
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Frame frm_Compare 
      Caption         =   "Compare"
      Height          =   1575
      Left            =   2400
      TabIndex        =   21
      Top             =   2880
      Width           =   4335
      Begin VB.CommandButton btn_Load 
         Caption         =   "Load"
         Height          =   360
         Left            =   3495
         TabIndex        =   38
         Top             =   240
         Width           =   720
      End
      Begin VB.OptionButton obn_Compare 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   2760
         TabIndex        =   36
         Top             =   720
         Value           =   -1  'True
         Width           =   480
      End
      Begin VB.OptionButton obn_Compare 
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   3720
         TabIndex        =   35
         Top             =   720
         Width           =   480
      End
      Begin VB.OptionButton obn_Compare 
         BackColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   3240
         TabIndex        =   34
         Top             =   720
         Width           =   480
      End
      Begin VB.OptionButton obn_Compare 
         BackColor       =   &H0000FFFF&
         Height          =   240
         Index           =   5
         Left            =   3720
         TabIndex        =   33
         Top             =   960
         Width           =   480
      End
      Begin VB.OptionButton obn_Compare 
         BackColor       =   &H00FF00FF&
         Height          =   240
         Index           =   4
         Left            =   3240
         TabIndex        =   32
         Top             =   960
         Width           =   480
      End
      Begin VB.OptionButton obn_Compare 
         BackColor       =   &H00FFFF00&
         Height          =   240
         Index           =   3
         Left            =   2760
         TabIndex        =   31
         Top             =   960
         Width           =   480
      End
      Begin VB.OptionButton obn_Compare 
         BackColor       =   &H00FF0000&
         Height          =   240
         Index           =   8
         Left            =   3720
         TabIndex        =   28
         Top             =   1200
         Width           =   480
      End
      Begin VB.OptionButton obn_Compare 
         BackColor       =   &H0000FF00&
         Height          =   240
         Index           =   7
         Left            =   3240
         TabIndex        =   27
         Top             =   1200
         Width           =   480
      End
      Begin VB.OptionButton obn_Compare 
         BackColor       =   &H000000FF&
         Height          =   240
         Index           =   6
         Left            =   2760
         TabIndex        =   26
         Top             =   1200
         Width           =   480
      End
      Begin VB.CommandButton btn_Save 
         Caption         =   "Save"
         Height          =   360
         Left            =   2745
         TabIndex        =   24
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lbl_Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   2750
         TabIndex        =   37
         Top             =   705
         Width           =   1470
      End
      Begin VB.Label lbl_CompareReference 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "FFFFFF"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lbl_CompareCurrent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "FFFFFF"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   1320
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lbl_Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1240
         Left            =   110
         TabIndex        =   25
         Top             =   230
         Width           =   2440
      End
   End
   Begin MSComDlg.CommonDialog cdg_Palette 
      Left            =   6720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frm_Display 
      Caption         =   "Display"
      Height          =   4335
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2055
      Begin VB.CommandButton btn_Inverse 
         Caption         =   "Inverse"
         Height          =   360
         Left            =   120
         TabIndex        =   30
         Top             =   3840
         Width           =   1800
      End
      Begin VB.CommandButton btn_Random 
         Caption         =   "Random"
         Height          =   360
         Left            =   120
         TabIndex        =   29
         Top             =   3360
         Width           =   1800
      End
      Begin VB.CommandButton btn_Image 
         Caption         =   "Image.."
         Height          =   360
         Left            =   120
         TabIndex        =   20
         Top             =   2880
         Width           =   1800
      End
      Begin VB.CommandButton btn_Screen 
         Caption         =   "Screen.."
         Height          =   360
         Left            =   120
         TabIndex        =   19
         Top             =   2520
         Width           =   1800
      End
      Begin VB.CommandButton btn_Palette 
         Caption         =   "Palette.."
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   1800
      End
      Begin VB.Label lbl_Display 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame frm_Edit 
      Caption         =   "Edit"
      Height          =   2655
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin MSComCtl2.UpDown sbn_B 
         Height          =   495
         Left            =   2520
         TabIndex        =   40
         Top             =   2040
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   873
         _Version        =   393216
         Value           =   255
         Max             =   255
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown sbn_G 
         Height          =   495
         Left            =   2520
         TabIndex        =   39
         Top             =   1440
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   873
         _Version        =   393216
         Value           =   255
         Max             =   255
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown sbn_R 
         Height          =   495
         Left            =   2520
         TabIndex        =   18
         Top             =   840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   873
         _Version        =   393216
         Value           =   255
         Max             =   255
         Enabled         =   -1  'True
      End
      Begin VB.TextBox tbx_BHex 
         Alignment       =   2  'Center
         BackColor       =   &H00FFCCCC&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00AA0000&
         Height          =   495
         Left            =   3600
         TabIndex        =   14
         Text            =   "FF"
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox tbx_BDec 
         Alignment       =   2  'Center
         BackColor       =   &H00FFCCCC&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00AA0000&
         Height          =   495
         Left            =   2760
         TabIndex        =   13
         Text            =   "255"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox tbx_GHex 
         Alignment       =   2  'Center
         BackColor       =   &H00CCFFCC&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000AA00&
         Height          =   495
         Left            =   3600
         TabIndex        =   12
         Text            =   "FF"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox tbx_GDec 
         Alignment       =   2  'Center
         BackColor       =   &H00CCFFCC&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000AA00&
         Height          =   495
         Left            =   2760
         TabIndex        =   11
         Text            =   "255"
         Top             =   1440
         Width           =   735
      End
      Begin MSComctlLib.Slider sld_Blue 
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   2040
         Width           =   2165
         _ExtentX        =   3810
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   16
         Max             =   255
         SelStart        =   255
         TickFrequency   =   4
         Value           =   255
      End
      Begin MSComctlLib.Slider sld_Green 
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   1440
         Width           =   2165
         _ExtentX        =   3810
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   16
         Max             =   255
         SelStart        =   255
         TickFrequency   =   4
         Value           =   255
      End
      Begin MSComctlLib.Slider sld_Red 
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   2165
         _ExtentX        =   3810
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   16
         Max             =   255
         SelStart        =   255
         TickFrequency   =   4
         Value           =   255
      End
      Begin VB.TextBox tbx_RHex 
         Alignment       =   2  'Center
         BackColor       =   &H00CCCCFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000AA&
         Height          =   495
         Left            =   3600
         TabIndex        =   7
         Text            =   "FF"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox tbx_RDec 
         Alignment       =   2  'Center
         BackColor       =   &H00CCCCFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000AA&
         Height          =   495
         Left            =   2760
         TabIndex        =   6
         Text            =   "255"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox tbx_VBColor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "FFFFFF"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lbl_Blue 
         Alignment       =   2  'Center
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00AA0000&
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label lbl_Green 
         Alignment       =   2  'Center
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000AA00&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lbl_Red 
         Alignment       =   2  'Center
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000AA&
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lbl_VBColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Color (BBGGRR):"
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
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Option "
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuCatchColor 
         Caption         =   "&Catch Color"
         Begin VB.Menu mnuPalette 
            Caption         =   "from Palette.."
         End
         Begin VB.Menu mnuScreen 
            Caption         =   "from Screen.."
         End
         Begin VB.Menu mnuImage 
            Caption         =   "from Image.."
         End
         Begin VB.Menu mnuRandom 
            Caption         =   "Random Color"
         End
      End
      Begin VB.Menu mnuInverseColor 
         Caption         =   "&Inverse Color"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About.."
      NegotiatePosition=   1  'Left
   End
End
Attribute VB_Name = "UI_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TODO screen/image select, RGB/BGR/YUV(combo box)
Option Explicit
Public intR As Integer, intG As Integer, intB As Integer
Private Updating As Boolean
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Load()
    intR = 255
    intG = 255
    intB = 255
    Updating = False
    Randomize
End Sub

Public Sub UpdateDisplay()
    If Not Updating Then
        Updating = True
        sld_Red.Value = intR
        sld_Green.Value = intG
        sld_Blue.Value = intB
        sbn_R.Value = intR
        sbn_G.Value = intG
        sbn_B.Value = intB
        tbx_RDec.Text = intR
        tbx_GDec.Text = intG
        tbx_BDec.Text = intB
        tbx_RHex.Text = Right("0" + Hex(intR), 2)
        tbx_GHex.Text = Right("0" + Hex(intG), 2)
        tbx_BHex.Text = Right("0" + Hex(intB), 2)
        tbx_VBColor.Text = tbx_BHex.Text + tbx_GHex.Text + tbx_RHex.Text
        lbl_Display.BackColor = CLng("&H" + tbx_VBColor.Text)
        lbl_CompareCurrent.BackColor = lbl_Display.BackColor
        lbl_CompareCurrent.Caption = tbx_VBColor.Text
        If intR + intG + intB > 381 Then lbl_CompareCurrent.ForeColor = 0 Else lbl_CompareCurrent.ForeColor = 16777215
        Updating = False
    End If
End Sub

Private Sub UpdateReferenceColor(r As Integer, g As Integer, b As Integer)
    lbl_CompareReference.BackColor = CLng(r) + 256 * CLng(g) + 65536 * CLng(b)
    If r + g + b > 381 Then lbl_CompareReference.ForeColor = 0 Else lbl_CompareReference.ForeColor = 16777215
    lbl_CompareReference.Caption = Right("0" + Hex(b), 2) + Right("0" + Hex(g), 2) + Right("0" + Hex(r), 2)
End Sub

Private Sub btn_Palette_Click()
    Dim l As Long
    cdg_Palette.ShowColor
    l = cdg_Palette.Color
    intR = l And &HFF
    intG = (l And 65280) \ 256
    intB = (l And 16711680) \ 65536
    UpdateDisplay
End Sub

Private Sub btn_Image_Click()
    cdg_Palette.Filter = "All Pictures|*.bmp;*.dib;*.jpg;*.gif;*.ico;*.cur;*.wmf;*.emf|Bitmap(*.bmp;*.dib)|*.bmp;*.dib|JPEG file (*.jpg)|*.jpg|metafile (*.wmf;*.emf)|*.wmf;*.emf|GIF file (*.gif)|*.gif|icon (*.ico;*.cur)|*.ico;*.cur"
    cdg_Palette.ShowOpen
    If cdg_Palette.FileName <> "" Then
        Err.Clear
        On Error Resume Next
        Set UI_Image.pic = LoadPicture(cdg_Palette.FileName)
        If Err.Number <> 0 Then
            MsgBox "Failed to open the specified file.", vbExclamation, "Error"
        Else
            Me.Hide
            UI_Image.Caption = "Image: " & cdg_Palette.FileTitle
            UI_Image.Show
        End If
    End If
End Sub

Private Sub btn_Screen_Click()
    Me.Hide
    Sleep 500
    UI_Screen.Show
End Sub

Private Sub btn_Random_Click()
    intR = Int(256 * Rnd())
    intG = Int(256 * Rnd())
    intB = Int(256 * Rnd())
    UpdateDisplay
End Sub

Private Sub btn_Inverse_Click()
    intR = 255 - intR
    intG = 255 - intG
    intB = 255 - intB
    UpdateDisplay
End Sub

Private Sub mnuAbout_Click()
    Me.Hide
    UI_About.Show
End Sub

Private Sub mnuPalette_Click()
    btn_Palette_Click
End Sub

Private Sub mnuScreen_Click()
    btn_Screen_Click
End Sub

Private Sub mnuImage_Click()
    btn_Image_Click
End Sub

Private Sub mnuRandom_Click()
    btn_Random_Click
End Sub

Private Sub mnuInverseColor_Click()
    btn_Inverse_Click
End Sub

Private Sub sld_Red_change()
    intR = sld_Red.Value
    UpdateDisplay
End Sub

Private Sub sld_Green_change()
    intG = sld_Green.Value
    UpdateDisplay
End Sub

Private Sub sld_Blue_change()
    intB = sld_Blue.Value
    UpdateDisplay
End Sub

Private Sub sbn_R_Change()
    intR = sbn_R.Value
    UpdateDisplay
End Sub

Private Sub sbn_G_Change()
    intG = sbn_G.Value
    UpdateDisplay
End Sub

Private Sub sbn_B_Change()
    intB = sbn_B.Value
    UpdateDisplay
End Sub

Private Sub tbx_RDec_Change()
    Dim c As Integer
    c = CInt("0" + tbx_RDec.Text)
    If c > 255 Then
        c = 255
    End If
    intR = c
    UpdateDisplay
End Sub

Private Sub tbx_GDec_Change()
   Dim c As Integer
    c = CInt("0" + tbx_GDec.Text)
    If c > 255 Then
        c = 255
    End If
    intG = c
    UpdateDisplay
End Sub

Private Sub tbx_BDec_Change()
    Dim c As Integer
    c = CInt("0" + tbx_BDec.Text)
    If c > 255 Then
        c = 255
    End If
    intB = c
    UpdateDisplay
End Sub

Private Sub tbx_RHex_Change()
    Dim s As String
    s = tbx_RHex.Text
    s = Right("00" + s, 2)
    intR = CInt("&H" + s)
    UpdateDisplay
End Sub

Private Sub tbx_GHex_Change()
    Dim s As String
    s = tbx_GHex.Text
    s = Right("00" + s, 2)
    intG = CInt("&H" + s)
    UpdateDisplay
End Sub

Private Sub tbx_BHex_Change()
    Dim s As String
    s = tbx_BHex.Text
    s = Right("00" + s, 2)
    intB = CInt("&H" + s)
    UpdateDisplay
End Sub

Private Sub tbx_RDec_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep
End Sub

Private Sub tbx_GDec_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep
End Sub

Private Sub tbx_BDec_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep
End Sub

Private Sub tbx_RHex_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And InStr("0123456789abcdefABCDEF", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep
End Sub

Private Sub tbx_GHex_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And InStr("0123456789abcdefABCDEF", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep
End Sub

Private Sub tbx_BHex_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And InStr("0123456789abcdefABCDEF", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep
End Sub

Private Sub btn_Save_Click()
    Dim i As Integer
    For i = 0 To 8
        If obn_Compare(i).Value = True Then obn_Compare(i).BackColor = CLng("&H" + tbx_VBColor.Text): UpdateReferenceColor intR, intG, intB: Exit For
    Next
End Sub

Private Sub btn_Load_Click()
    Dim i As Integer
    Dim l As Long
    For i = 0 To 8
        If obn_Compare(i).Value = True Then l = obn_Compare(i).BackColor: Exit For
    Next
    intR = l And &HFF
    intG = (l And 65280) \ 256
    intB = (l And &HFF0000) \ 65536
    UpdateDisplay
End Sub

Private Sub obn_Compare_Click(Index As Integer)
    Dim l As Long
    l = obn_Compare(Index).BackColor
    UpdateReferenceColor (l And &HFF), (l And 65280) \ 256, (l And 16711680) \ 65536
End Sub
