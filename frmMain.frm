VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Winamp XP 1.00"
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   2880
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgOpenFiles 
      Left            =   3360
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open file(s)"
      Filter          =   $"frmMain.frx":2BE0
   End
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2880
      Top             =   2280
   End
   Begin PicClip.PictureClip clipTitlebar 
      Left            =   1440
      Top             =   1800
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin PicClip.PictureClip clipCButtons 
      Left            =   0
      Top             =   1800
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin VB.PictureBox Mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1740
      Left            =   3840
      Picture         =   "frmMain.frx":2C67
      ScaleHeight     =   1740
      ScaleWidth      =   4125
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   4125
   End
   Begin PicClip.PictureClip clipShufRep 
      Left            =   1920
      Top             =   1800
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin PicClip.PictureClip clipPosBar 
      Left            =   480
      Top             =   1800
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin PicClip.PictureClip clipVolume 
      Left            =   960
      Top             =   1800
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin PicClip.PictureClip clipBalance 
      Left            =   2400
      Top             =   1800
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin PicClip.PictureClip clipPlayPaus 
      Left            =   2880
      Top             =   1800
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin PicClip.PictureClip clipNumbers 
      Left            =   3360
      Top             =   1800
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin PicClip.PictureClip clipMonoster 
      Left            =   3840
      Top             =   1800
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin VB.Image imgMonoSter 
      Height          =   255
      Index           =   1
      Left            =   3480
      Top             =   600
      Width           =   495
   End
   Begin VB.Image imgMonoSter 
      Height          =   255
      Index           =   0
      Left            =   3000
      Top             =   600
      Width           =   495
   End
   Begin MediaPlayerCtl.MediaPlayer mediaPlay 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   30
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   0   'False
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Image imgNumber 
      Height          =   255
      Index           =   3
      Left            =   1320
      Top             =   360
      Width           =   255
   End
   Begin VB.Image imgNumber 
      Height          =   255
      Index           =   2
      Left            =   1080
      Top             =   360
      Width           =   255
   End
   Begin VB.Image imgNumber 
      Height          =   255
      Index           =   1
      Left            =   840
      Top             =   360
      Width           =   255
   End
   Begin VB.Image imgNumber 
      Height          =   255
      Index           =   0
      Left            =   600
      Top             =   360
      Width           =   255
   End
   Begin VB.Image imgBalanceBut 
      Height          =   255
      Index           =   1
      Left            =   2760
      Top             =   840
      Width           =   255
   End
   Begin VB.Image imgVolumeBut 
      Height          =   255
      Index           =   1
      Left            =   1800
      Top             =   840
      Width           =   255
   End
   Begin VB.Image imgPlayPaus 
      Height          =   255
      Index           =   0
      Left            =   360
      Top             =   360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgPlayPaus 
      Height          =   255
      Index           =   1
      Left            =   480
      Top             =   360
      Width           =   135
   End
   Begin VB.Image imgPosSlide 
      Height          =   135
      Index           =   1
      Left            =   720
      Top             =   1080
      Width           =   615
   End
   Begin VB.Image imgBalanceBut 
      Height          =   255
      Index           =   0
      Left            =   2880
      Top             =   840
      Width           =   255
   End
   Begin VB.Image imgBalance 
      Height          =   135
      Left            =   2760
      Top             =   840
      Width           =   495
   End
   Begin VB.Image imgVolumeBut 
      Height          =   255
      Index           =   0
      Left            =   2040
      Top             =   840
      Width           =   255
   End
   Begin VB.Image imgVolume 
      Height          =   135
      Left            =   1680
      Top             =   840
      Width           =   975
   End
   Begin VB.Image imgAmp 
      Height          =   375
      Left            =   3720
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image imgPosSlide 
      Height          =   135
      Index           =   0
      Left            =   1200
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgPosBar 
      Height          =   135
      Left            =   240
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Image imgSufRep 
      Height          =   255
      Index           =   3
      Left            =   3600
      Tag             =   "0"
      Top             =   840
      Width           =   375
   End
   Begin VB.Image imgSufRep 
      Height          =   255
      Index           =   2
      Left            =   3240
      Tag             =   "0"
      Top             =   840
      Width           =   375
   End
   Begin VB.Image imgSufRep 
      Height          =   255
      Index           =   0
      Left            =   3120
      Tag             =   "0"
      Top             =   1320
      Width           =   495
   End
   Begin VB.Image imgSufRep 
      Height          =   255
      Index           =   1
      Left            =   2640
      Tag             =   "0"
      Top             =   1320
      Width           =   495
   End
   Begin VB.Image imgMenu 
      Height          =   255
      Index           =   3
      Left            =   3840
      Top             =   0
      Width           =   255
   End
   Begin VB.Image imgMenu 
      Height          =   255
      Index           =   2
      Left            =   3600
      Top             =   0
      Width           =   255
   End
   Begin VB.Image imgMenu 
      Height          =   255
      Index           =   1
      Left            =   3360
      Top             =   0
      Width           =   255
   End
   Begin VB.Image imgMenu 
      Height          =   255
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Image imgPlayBut 
      Height          =   255
      Index           =   5
      Left            =   2040
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image imgPlayBut 
      Height          =   375
      Index           =   4
      Left            =   1680
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image imgPlayBut 
      Height          =   375
      Index           =   3
      Left            =   1320
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image imgPlayBut 
      Height          =   375
      Index           =   2
      Left            =   960
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image imgPlayBut 
      Height          =   375
      Index           =   1
      Left            =   600
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image imgPlayBut 
      Height          =   375
      Index           =   0
      Left            =   240
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image imgMove 
      Height          =   375
      Left            =   1920
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgTitlebar 
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Const RGN_OR = 2
Private lngRegion As Long

Private intPosMove As Integer
Private intVolMove As Integer
Private intBalMove As Integer

Private intSec As Integer
Private intMin As Integer
Private blnRepos As Boolean         '= True when you use the posbar
Private Sub SecondsToTime(lSeconds As Double)
    intSec = Abs(Fix(lSeconds)) Mod 60
    intMin = Fix(Abs(Fix(lSeconds)) / 60)
End Sub
Private Function RegionFromBitmap(picSource As PictureBox, Optional lngTransColor As Long) As Long
    Dim lngRetr As Long, lngHeight As Long, lngWidth As Long
    Dim lngRgnFinal As Long, lngRgnTmp As Long
    Dim lngStart As Long, lngRow As Long
    Dim lngCol As Long
    If lngTransColor& < 1 Then
        lngTransColor& = GetPixel(picSource.hDC, 0, 0)
    End If
    lngHeight& = picSource.Height / Screen.TwipsPerPixelY
    lngWidth& = picSource.Width / Screen.TwipsPerPixelX
    lngRgnFinal& = CreateRectRgn(0, 0, 0, 0)
    For lngRow& = 0 To lngHeight& - 1
        lngCol& = 0
        Do While lngCol& < lngWidth&
            Do While lngCol& < lngWidth& And GetPixel(picSource.hDC, lngCol&, lngRow&) = lngTransColor&
                lngCol& = lngCol& + 1
            Loop
            If lngCol& < lngWidth& Then
                lngStart& = lngCol&
                Do While lngCol& < lngWidth& And GetPixel(picSource.hDC, lngCol&, lngRow&) <> lngTransColor&
                    lngCol& = lngCol& + 1
                Loop
                If lngCol& > lngWidth& Then lngCol& = lngWidth&
                lngRgnTmp& = CreateRectRgn(lngStart&, lngRow&, lngCol&, lngRow& + 1)
                lngRetr& = CombineRgn(lngRgnFinal&, lngRgnFinal&, lngRgnTmp&, RGN_OR)
                DeleteObject (lngRgnTmp&)
            End If
        Loop
    Next
    RegionFromBitmap& = lngRgnFinal&
End Function

Public Sub ChangeMask(Optional lngTransColor As Long = &HFF00FF)
    On Error Resume Next
    
    Dim lngRetr As Long
    lngRegion& = RegionFromBitmap(Mask, lngTransColor)
    lngRetr& = SetWindowRgn(Me.hwnd, lngRegion&, True)
End Sub
Public Sub MoveControls()
    imgPlayBut(0).Move 16 * Screen.TwipsPerPixelX, 88 * Screen.TwipsPerPixelY, 22 * Screen.TwipsPerPixelX, 17 * Screen.TwipsPerPixelY
    imgPlayBut(1).Move 39 * Screen.TwipsPerPixelX, 88 * Screen.TwipsPerPixelY, 22 * Screen.TwipsPerPixelX, 17 * Screen.TwipsPerPixelY
    imgPlayBut(2).Move 62 * Screen.TwipsPerPixelX, 88 * Screen.TwipsPerPixelY, 22 * Screen.TwipsPerPixelX, 17 * Screen.TwipsPerPixelY
    imgPlayBut(3).Move 85 * Screen.TwipsPerPixelX, 88 * Screen.TwipsPerPixelY, 22 * Screen.TwipsPerPixelX, 17 * Screen.TwipsPerPixelY
    imgPlayBut(4).Move 108 * Screen.TwipsPerPixelX, 88 * Screen.TwipsPerPixelY, 22 * Screen.TwipsPerPixelX, 17 * Screen.TwipsPerPixelY
    imgPlayBut(5).Move 136 * Screen.TwipsPerPixelX, 89 * Screen.TwipsPerPixelY, 21 * Screen.TwipsPerPixelX, 15 * Screen.TwipsPerPixelY
    
    imgTitlebar.Move 0, 0, 274 * Screen.TwipsPerPixelX, 12 * Screen.TwipsPerPixelY
    
    imgMenu(0).Move 6 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY, 9 * Screen.TwipsPerPixelX, 9 * Screen.TwipsPerPixelY
    imgMenu(1).Move 244 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY, 9 * Screen.TwipsPerPixelX, 9 * Screen.TwipsPerPixelY
    imgMenu(2).Move 264 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY, 9 * Screen.TwipsPerPixelX, 9 * Screen.TwipsPerPixelY
    imgMenu(3).Move 254 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY, 9 * Screen.TwipsPerPixelX, 9 * Screen.TwipsPerPixelY
    
    imgSufRep(0).Move 210 * Screen.TwipsPerPixelX, 89 * Screen.TwipsPerPixelY, 29 * Screen.TwipsPerPixelX, 15 * Screen.TwipsPerPixelY
    imgSufRep(1).Move 164 * Screen.TwipsPerPixelX, 89 * Screen.TwipsPerPixelY, 47 * Screen.TwipsPerPixelX, 15 * Screen.TwipsPerPixelY
    
    imgSufRep(2).Move 219 * Screen.TwipsPerPixelX, 58 * Screen.TwipsPerPixelY, 23 * Screen.TwipsPerPixelX, 12 * Screen.TwipsPerPixelY
    imgSufRep(3).Move 242 * Screen.TwipsPerPixelX, 58 * Screen.TwipsPerPixelY, 23 * Screen.TwipsPerPixelX, 12 * Screen.TwipsPerPixelY
    
    imgPosBar.Move 16 * Screen.TwipsPerPixelX, 72 * Screen.TwipsPerPixelY, 249 * Screen.TwipsPerPixelX, 10 * Screen.TwipsPerPixelY
    imgPosSlide(0).Move 16 * Screen.TwipsPerPixelX, 72 * Screen.TwipsPerPixelY, 29 * Screen.TwipsPerPixelX, 10 * Screen.TwipsPerPixelY
    imgPosSlide(1).Move 16 * Screen.TwipsPerPixelX, 72 * Screen.TwipsPerPixelY, 29 * Screen.TwipsPerPixelX, 10 * Screen.TwipsPerPixelY
    
    imgAmp.Move 251 * Screen.TwipsPerPixelX, 91 * Screen.TwipsPerPixelY, 14 * Screen.TwipsPerPixelX, 16 * Screen.TwipsPerPixelY
    
    imgVolume.Move 107 * Screen.TwipsPerPixelX, 57 * Screen.TwipsPerPixelY, 68 * Screen.TwipsPerPixelX, 13 * Screen.TwipsPerPixelY
    imgVolumeBut(0).Move 135 * Screen.TwipsPerPixelX, 58 * Screen.TwipsPerPixelY, 14 * Screen.TwipsPerPixelX, 11 * Screen.TwipsPerPixelY
    imgVolumeBut(1).Move 135 * Screen.TwipsPerPixelX, 58 * Screen.TwipsPerPixelY, 14 * Screen.TwipsPerPixelX, 11 * Screen.TwipsPerPixelY
    
    imgBalance.Move 177 * Screen.TwipsPerPixelX, 57 * Screen.TwipsPerPixelY, 38 * Screen.TwipsPerPixelX, 13 * Screen.TwipsPerPixelY
    imgBalanceBut(0).Move 189 * Screen.TwipsPerPixelX, 58 * Screen.TwipsPerPixelY, 14 * Screen.TwipsPerPixelX, 11 * Screen.TwipsPerPixelY
    imgBalanceBut(1).Move 189 * Screen.TwipsPerPixelX, 58 * Screen.TwipsPerPixelY, 14 * Screen.TwipsPerPixelX, 11 * Screen.TwipsPerPixelY
    
    imgPlayPaus(0).Move 24 * Screen.TwipsPerPixelX, 28 * Screen.TwipsPerPixelY, 9 * Screen.TwipsPerPixelX, 9 * Screen.TwipsPerPixelY
    imgPlayPaus(1).Move 26 * Screen.TwipsPerPixelX, 28 * Screen.TwipsPerPixelY, 3 * Screen.TwipsPerPixelX, 9 * Screen.TwipsPerPixelY
    
    imgNumber(0).Move 48 * Screen.TwipsPerPixelX, 26 * Screen.TwipsPerPixelY, 9 * Screen.TwipsPerPixelX, 13 * Screen.TwipsPerPixelY
    imgNumber(1).Move 60 * Screen.TwipsPerPixelX, 26 * Screen.TwipsPerPixelY, 9 * Screen.TwipsPerPixelX, 13 * Screen.TwipsPerPixelY
    imgNumber(2).Move 78 * Screen.TwipsPerPixelX, 26 * Screen.TwipsPerPixelY, 9 * Screen.TwipsPerPixelX, 13 * Screen.TwipsPerPixelY
    imgNumber(3).Move 90 * Screen.TwipsPerPixelX, 26 * Screen.TwipsPerPixelY, 9 * Screen.TwipsPerPixelX, 13 * Screen.TwipsPerPixelY
    
    imgMonoSter(1).Move 212 * Screen.TwipsPerPixelX, 41 * Screen.TwipsPerPixelY, 29 * Screen.TwipsPerPixelX, 12 * Screen.TwipsPerPixelY
    imgMonoSter(0).Move 239 * Screen.TwipsPerPixelX, 41 * Screen.TwipsPerPixelY, 29 * Screen.TwipsPerPixelX, 12 * Screen.TwipsPerPixelY
End Sub

Private Sub Form_GotFocus()
    Dim q As Integer
    GetClip obtTitlebar, 1, obsDefault
    For q = 4 To 7
        GetClip obtTitlebar, q, obsDefault
    Next q
End Sub

Private Sub Form_Load()
    LoadSkin "C:\Documents and Settings\Daniel\Mijn documenten\Mijn downloads\Winamp\viskin"      'phoe   'viskin     'baseskin
    GoStop False
End Sub

Private Sub Form_LostFocus()
    Dim q As Integer
    GetClip obtTitlebar, 1, obsPressed
    For q = 0 To 3
        imgMenu(q).Picture = LoadPicture()
    Next q
End Sub

Private Sub imgBalanceBut_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 1 And Button = vbLeftButton Then
        intBalMove = X
        GetClip obtBalance, 30, obsDefault
    End If
End Sub


Private Sub imgBalanceBut_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngA As Single, intA As Integer
    Dim sngMin As Single
    sngMin = 177 * Screen.TwipsPerPixelX
    Dim sngMax As Single
    sngMax = 215 * Screen.TwipsPerPixelX - imgBalanceBut(0).Width
    
    If intBalMove And Index = 1 Then
        sngA = X + imgBalanceBut(1).Left - intBalMove
        If sngA < sngMin + 3 * Screen.TwipsPerPixelX Then
            imgBalanceBut(0).Left = sngMin
        ElseIf sngA > sngMax - 3 * Screen.TwipsPerPixelX Then
            imgBalanceBut(0).Left = sngMax
        ElseIf sngA > (sngMax - sngMin) / 2 - 3 * Screen.TwipsPerPixelX + sngMin And sngA < (sngMax - sngMin) / 2 + 3 * Screen.TwipsPerPixelX + sngMin Then
            imgBalanceBut(0).Left = (sngMax - sngMin) / 2 + sngMin
        Else
            imgBalanceBut(0).Left = sngA
        End If
        If imgBalanceBut(0).Left - sngMin < (sngMax - sngMin) / 2 Then
            intA = 28 - Round((imgBalanceBut(0).Left - sngMin) / ((sngMax - sngMin) / 2 / 28), 0)
        ElseIf imgBalanceBut(0).Left - sngMin > (sngMax - sngMin) / 2 Then
            intA = Round((imgBalanceBut(0).Left - sngMin - (sngMax - sngMin) / 2) / ((sngMax - sngMin) / 2 / 28), 0)
        Else
            intA = 1
        End If
        If intA <= 0 Then intA = 1
        If intA >= 28 Then intA = 28
        GetClip obtBalance, intA, obsDefault
        'mediaPlay.Balance = (6000 / 28) * intA - 3000
        'mediaPlay.Balance = (3000 / 28) * intA - 3000
        mediaPlay.Balance = ((imgBalanceBut(0).Left - sngMin) / (sngMax - sngMin)) * 6000 - 3000
    End If
End Sub


Private Sub imgBalanceBut_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 1 And intBalMove Then
        intBalMove = 0
        GetClip obtBalance, 29, obsDefault
        imgBalanceBut(1).Left = imgBalanceBut(0).Left
    End If
End Sub


Private Sub imgMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetClip obtTitlebar, Index + 4, obsPressed
End Sub

Private Sub imgMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetClip obtTitlebar, Index + 4, obsDefault
End Sub


Private Sub imgMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        Call SendMessage(Me.hwnd, &HA1, 2, 0)
    End If
End Sub


Private Sub imgPlayBut_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetClip obtCButtons, Index + 1, obsPressed
End Sub


Private Sub imgPlayBut_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo 0
    On Error GoTo stpError
    
    GetClip obtCButtons, Index + 1, obsDefault
    Select Case Index
        Case 0
            
        Case 1
            GoPlay
        Case 2
            GoPause
        Case 3
            GoStop
        Case 4
            
        Case 5
            dlgOpenFiles.ShowOpen
            mediaPlay.FileName = dlgOpenFiles.FileName
            GoPlay False
    End Select
    Exit Sub
stpError:
    Select Case Err
        Case 32755
    End Select
End Sub


Private Sub imgPosSlide_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 1 And Button = vbLeftButton Then
        intPosMove = X
        GetClip obtPosBar, 2, obsPressed
        blnRepos = True
    End If
End Sub

Private Sub imgPosSlide_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngA As Single
    Dim sngMin As Single
    sngMin = 16 * Screen.TwipsPerPixelX
    Dim sngMax As Single
    sngMax = 264 * Screen.TwipsPerPixelX - imgPosSlide(0).Width
    
    If intPosMove And Index = 1 Then
        sngA = X + imgPosSlide(1).Left - intPosMove
        If sngA < sngMin Then
            imgPosSlide(0).Left = sngMin
        ElseIf sngA > sngMax Then
            imgPosSlide(0).Left = sngMax
        Else
            imgPosSlide(0).Left = sngA
        End If
    End If
End Sub

Private Sub imgPosSlide_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 1 And intPosMove Then
        mediaPlay.CurrentPosition = ((imgPosSlide(0).Left - imgPosBar.Left) / (imgPosBar.Width - imgPosSlide(0).Width)) * mediaPlay.Duration
        intPosMove = 0
        GetClip obtPosBar, 2, obsDefault
        imgPosSlide(1).Left = imgPosSlide(0).Left
        blnRepos = False
    End If
End Sub

Private Sub imgSufRep_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgSufRep(Index).Tag = "0" Then
        GetClip obtShufRep, Index + 1, obsPressed
    Else
        GetClip obtShufRep, Index + 1, obsOnPressed
    End If
End Sub


Private Sub imgSufRep_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgSufRep(Index).Tag = "0" Then
        GetClip obtShufRep, Index + 1, obsOn
        imgSufRep(Index).Tag = "1"
    Else
        GetClip obtShufRep, Index + 1, obsDefault
        imgSufRep(Index).Tag = "0"
    End If
End Sub


Private Sub imgVolume_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        imgVolumeBut(0).Left = X + imgVolume.Left - imgVolumeBut(0).Width / 2
        imgVolumeBut(1).Left = imgVolumeBut(0).Left
    End If
End Sub


Private Sub imgVolumeBut_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 1 And Button = vbLeftButton Then
        intVolMove = X
        GetClip obtVolume, 30, obsDefault
    End If
End Sub

Private Sub imgVolumeBut_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngA As Single, intA As Integer
    Dim sngMin As Single
    sngMin = 107 * Screen.TwipsPerPixelX
    Dim sngMax As Single
    sngMax = 172 * Screen.TwipsPerPixelX - imgVolumeBut(0).Width
    
    If intVolMove And Index = 1 Then
        sngA = X + imgVolumeBut(1).Left - intVolMove
        If sngA < sngMin Then
            imgVolumeBut(0).Left = sngMin
        ElseIf sngA > sngMax Then
            imgVolumeBut(0).Left = sngMax
        Else
            imgVolumeBut(0).Left = sngA
        End If
        intA = Round((imgVolumeBut(0).Left - sngMin) / ((sngMax - sngMin) / 28), 0)
        If intA <= 0 Then intA = 1
        If intA >= 28 Then intA = 28
        GetClip obtVolume, intA, obsDefault
        mediaPlay.Volume = ((imgVolumeBut(0).Left - sngMin) / (sngMax - sngMin)) * 3000 - 3000
    End If
End Sub


Private Sub imgVolumeBut_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 1 And intVolMove Then
        intVolMove = 0
        GetClip obtVolume, 29, obsDefault
        imgVolumeBut(1).Left = imgVolumeBut(0).Left
    End If
End Sub



Private Sub tmrTime_Timer()
    If mediaPlay.PlayState = mpPlaying Then
        SecondsToTime mediaPlay.CurrentPosition
        ShowTime
        If Not blnRepos Then
            imgPosSlide(0).Move (mediaPlay.CurrentPosition / mediaPlay.Duration) * (imgPosBar.Width - imgPosSlide(0).Width) + imgPosBar.Left
            imgPosSlide(1).Move imgPosSlide(0).Left
        End If
    End If
    Select Case mediaPlay.PlayState
        Case mpClosed Or mpStopped
            GoStop False
        Case mpPaused
            GoPause False
        Case mpPlaying
            GoPlay False
    End Select
End Sub
Public Sub ShowTime()
    Dim a As String, b As Integer, c As Integer
    a = "0" & CStr(intSec)
    b = CInt(Right(a, 1))
    c = CInt(Left(Right(a, 2), 1))
    GetClip obtNumers, b, 3
    GetClip obtNumers, c, 2
    
    a = "0" & CStr(intMin)
    b = CInt(Right(a, 1))
    c = CInt(Left(Right(a, 2), 1))
    GetClip obtNumers, b, 1
    GetClip obtNumers, c, 0
End Sub

Public Sub GoPlay(Optional blnDo As Boolean = True)
    If blnDo Then mediaPlay.Play
    GetClip obtPlayPaus, 1, obsDefault
    GetClip obtPlayPaus, 5, obsDefault
    imgPlayPaus(0).Visible = True
    tmrTime.Enabled = True
    For q = 0 To 3
        imgNumber(q).Visible = True
    Next q
    imgPosSlide(0).Visible = True
    imgPosSlide(1).Visible = True
End Sub

Public Sub GoPause(Optional blnDo As Boolean = True)
    If blnDo Then mediaPlay.Pause
    GetClip obtPlayPaus, 2, obsDefault
    imgPlayPaus(0).Visible = False
    tmrTime.Enabled = False
    For q = 0 To 3
        imgNumber(q).Visible = True
    Next q
    imgPosSlide(0).Visible = True
    imgPosSlide(1).Visible = True
End Sub

Public Sub GoStop(Optional blnDo As Boolean = True)
    If blnDo Then mediaPlay.Stop
    mediaPlay.CurrentPosition = 0
    GetClip obtPlayPaus, 3, obsDefault
    imgPlayPaus(0).Visible = False
    For q = 0 To 3
        imgNumber(q).Visible = False
    Next q
    tmrTime.Enabled = False
    imgPosSlide(0).Visible = False
    imgPosSlide(1).Visible = False
End Sub
