VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sprite demo"
   ClientHeight    =   6525
   ClientLeft      =   660
   ClientTop       =   600
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   435
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   Begin VB.VScrollBar VS 
      Height          =   1620
      LargeChange     =   10
      Left            =   315
      Max             =   0
      Min             =   300
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4275
      Width           =   240
   End
   Begin VB.PictureBox picCopy 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   285
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   74
      TabIndex        =   3
      Top             =   2970
      Width           =   1110
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   255
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   1
      Top             =   1590
      Width           =   1170
   End
   Begin VB.PictureBox picP 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   270
      Picture         =   "Form1.frx":4916
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   0
      Top             =   240
      Width           =   1170
   End
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   9060
      Left            =   1920
      Picture         =   "Form1.frx":922C
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   2
      Top             =   270
      Width           =   12060
   End
   Begin VB.Label Label1 
      Caption         =   "Delay"
      Height          =   255
      Left            =   660
      TabIndex        =   6
      Top             =   4620
      Width           =   465
   End
   Begin VB.Label LabET 
      Height          =   255
      Left            =   615
      TabIndex        =   4
      Top             =   5010
      Width           =   645
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Sprite Demo

' Starting point was VBHelper "howto_overlay_moving_picture"
' Sprite made with Region Selector (PSC  CodeId=60865)

Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private PrevX As Long
Private PrevY As Long
Private LX As Long
Private LY As Long

Private W As Long
Private H As Long
Private Wmax As Long
Private Hmax As Long
Private StepX As Long
Private StepY As Long

Private aDone As Boolean
Private ET As Long

Dim tmDelay As CTiming

Private Sub Form_Load()
   Set tmDelay = New CTiming
   
   ET = 40  ' Start Delay
   VS.Value = 40
   Show
   ' Small image with black surround
   W = picP.Width
   H = picP.Height
   ' Make picCopy same size
   picCopy.Width = W
   picCopy.Height = H
   
   ' Max X & Y that sprite can go
   Wmax = picDisplay.Width - W
   Hmax = picDisplay.Height - H
   ' Starting positions
   PrevX = 20
   PrevY = 25
   LX = 20
   LY = 25
   
   StepX = 1
   StepY = 1
   
   ' Make copy of image where sprite is going
   BitBlt picCopy.hDC, _
       0, 0, W, H, _
       picDisplay.hDC, PrevX, PrevY, vbSrcCopy
   picCopy.Refresh
   
   aDone = False
Do
   LX = LX + StepX
   LY = LY + StepY
   
   ' Ensure stays on picDisplay
   If LX < 0 Then
       LX = 0
       StepX = -StepX     ' Reverse steps
   End If
   If LX > Wmax Then
       LX = Wmax
       StepX = -StepX
   End If
   If LY < 0 Then
       LY = 0
       StepY = -StepY
   End If
   If LY > Hmax Then
       LY = Hmax
       StepY = -StepY
   End If
   
   DrawPicture
   
   tmDelay.Reset
   Do
   Loop Until tmDelay.Elapsed > ET
   
   DoEvents
Loop Until aDone

End Sub

Private Sub DrawPicture()
    ' Put back part of the image that was covered.
    BitBlt picDisplay.hDC, _
        PrevX, PrevY, W, H, _
        picCopy.hDC, 0, 0, vbSrcCopy
    
    PrevX = LX
    PrevY = LY
    ' Get new background picture
    BitBlt picCopy.hDC, _
        0, 0, W, H, _
        picDisplay.hDC, PrevX, PrevY, vbSrcCopy
    picCopy.Refresh
    
    ' Paint on the new image.
    BitBlt picDisplay.hDC, _
        LX, LY, W, H, _
        picMask.hDC, 0, 0, vbSrcAnd
    BitBlt picDisplay.hDC, _
        LX, LY, W, H, _
        picP.hDC, 0, 0, vbSrcPaint

    ' Update the display.
    picDisplay.Refresh
End Sub

Private Sub VS_Change()
   Call VS_Scroll
End Sub

Private Sub VS_Scroll()
   ET = VS.Value
   LabET = Str$(ET)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   aDone = True
   Set tmDelay = Nothing
   Unload Me
   End
End Sub

Private Sub Command1_Click()
' Close
   aDone = True
   Set tmDelay = Nothing
   Unload Me
   End
End Sub

