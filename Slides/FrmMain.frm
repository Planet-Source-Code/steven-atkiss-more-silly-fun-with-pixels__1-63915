VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   9225
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicRange 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   1695
      Index           =   2
      Left            =   300
      Picture         =   "FrmMain.frx":0000
      ScaleHeight     =   1695
      ScaleWidth      =   2580
      TabIndex        =   5
      Top             =   7320
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   8940
      Top             =   660
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "End"
      Height          =   555
      Left            =   7380
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8100
      Width           =   1575
   End
   Begin VB.PictureBox Canvas 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   180
      Picture         =   "FrmMain.frx":E406
      ScaleHeight     =   1695
      ScaleWidth      =   2625
      TabIndex        =   3
      Top             =   120
      Width           =   2625
   End
   Begin VB.PictureBox PicRange 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Index           =   0
      Left            =   60
      Picture         =   "FrmMain.frx":1D9B4
      ScaleHeight     =   1695
      ScaleWidth      =   2775
      TabIndex        =   2
      Top             =   3780
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox PicOutPut 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   60
      Picture         =   "FrmMain.frx":2CF62
      ScaleHeight     =   1695
      ScaleWidth      =   2775
      TabIndex        =   1
      Top             =   1980
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox PicRange 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   1695
      Index           =   1
      Left            =   60
      Picture         =   "FrmMain.frx":3C510
      ScaleHeight     =   1695
      ScaleWidth      =   1710
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   1710
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'NOTE:  YOU CAN HAVE AS MANY DIFERANT PICTURES AS YOU WANT, JUST ADD MORE CONTROLS TO THE PICRANGE CONTROL
'       ARRAY AND LOAD EM UP WITH YOUR PICTURES. THE PROGRAM AUTOMATICALLY COMPENSATES FOR HOW MANY PICTURES
'       YOU HAVE. EG    PicRange(0), PicRange(1), PicRange(2), PicRange(3), PicRange(4)


Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Byte

'View Mode Variables
Private Mode As Single
Private ChangeView As Boolean
Private Steps As Single



Private Sub Command1_Click()
    
    'EErrrrr
    End
    
End Sub

Private Sub SlideIcons()
    
    Dim Movement As Single
    Movement = 5
    
    If ChangeView = True Then Exit Sub
    
    ChangeView = True
    
    Dim LP  As Single
    Dim Start As Single
    
    Mode = Mode + 1
    
    
    If Mode > PicRange.UBound Then Mode = 0
    
    Dim X As Long, Y As Long, Rep As Long
    
    'Move Slide Out
    For Start = 0 To Steps - 1
        For Rep = 0 To T2P(Canvas.Width) Step Movement
            For Y = Start To T2P(Canvas.Height) Step Steps
                For X = 0 To T2P(Canvas.Width) - Rep
                    SetPixelV Canvas.hDC, X, Y, GetPixel(PicOutPut.hDC, X + Rep, Y)
                Next X
            Next Y
            DoEvents
        Next Rep
    Next Start
    
    
    
    
    'Move The Canvas
    Canvas.Picture = LoadPicture("")
    Canvas.Top = ((Screen.Height - Canvas.Height) * Rnd)
    Canvas.Left = ((Screen.Width - Canvas.Width) * Rnd)
    
    'Adjusting The Steps Variable Adjusts How The Trasition Sweeps In\Out
    Steps = (Int(T2P(Canvas.Height) / 2) * Rnd) + 1
    
    'Move Next Slide In
    For Start = 0 To Steps - 1
        For Rep = 0 To T2P(Canvas.Width) Step Movement
            For Y = Start To T2P(Canvas.Height) Step Steps
                For X = 0 To Rep
                    SetPixelV Canvas.hDC, T2P(Canvas.Width) - Rep + X, Y, GetPixel(PicRange(Mode).hDC, X, Y)
                Next X
            Next Y
            DoEvents
        Next Rep
    Next Start
    
    
    
    PicOutPut.Picture = PicRange(Mode).Picture
    Canvas.Picture = PicRange(Mode).Picture
    
    ChangeView = False
    
End Sub

Private Function T2P(Twip As Long) As Long
    
    'Convert Twip To Pixel For Graphics Routines
    T2P = Int(Twip / 15)

End Function

Private Sub Form_Load()

Randomize Timer

    Steps = (20 * Rnd) + 1
    
    Command1.Top = Screen.Height - Command1.Height - 100
    Command1.Left = Screen.Width - Command1.Width - 100
    
End Sub

Private Sub Timer1_Timer()

    SlideIcons
    
End Sub



