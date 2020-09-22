VERSION 5.00
Begin VB.Form frmBGRSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Brush gradient colors settings"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboSize 
      Height          =   315
      ItemData        =   "frmBGRSettings.frx":0000
      Left            =   1080
      List            =   "frmBGRSettings.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   615
      Left            =   1320
      TabIndex        =   6
      Top             =   3600
      Width           =   1575
   End
   Begin VB.PictureBox picPreview 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   1440
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.PictureBox picClr2 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   1680
      ScaleHeight     =   555
      ScaleWidth      =   915
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox picClr1 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   1680
      ScaleHeight     =   555
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblSize 
      Alignment       =   2  'Center
      Caption         =   "Gradient size:"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Width           =   4095
   End
   Begin VB.Label lblPreview 
      Alignment       =   2  'Center
      Caption         =   "Preview:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Label lblColor2 
      Caption         =   "Color 2:"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblColor1 
      Caption         =   "Color 1:"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "frmBGRSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Color1 As Long
Dim Color2 As Long


Private Sub cboSize_Click()
If cboSize.Text = "Small" Then
    picPreview.Width = 1215
    picPreview.Left = Me.Width \ 2 - 1215 \ 2
ElseIf cboSize.Text = "Medium" Then
    picPreview.Width = 2415
    picPreview.Left = Me.Width \ 2 - 2415 \ 2
ElseIf cboSize.Text = "Large" Then
    picPreview.Width = 3855
    picPreview.Left = Me.Width \ 2 - 3855 \ 2
End If

PaintPreview
End Sub

Private Sub cmdOK_Click()
brush_gradient_colors_internal_OnSettOK
End Sub

Private Sub Form_Load()
PaintPreview

cboSize_Click

picClr1.BackColor = Color1
picClr2.BackColor = Color2
End Sub

Private Sub picClr1_Click()
frmMain.cdlColor.Color = Color1
frmMain.cdlColor.ShowColor

Color1 = frmMain.cdlColor.Color
picClr1.BackColor = Color1

PaintPreview
End Sub

Private Sub picClr2_Click()
frmMain.cdlColor.Color = Color2
frmMain.cdlColor.ShowColor

Color2 = frmMain.cdlColor.Color
picClr2.BackColor = Color2

PaintPreview
End Sub

Sub PaintPreview()

Dim r1 As Double, r2 As Double
Dim g1 As Double, g2 As Double
Dim b1 As Double, b2 As Double

Dim rCurr As Double
Dim gCurr As Double
Dim bCurr As Double

Dim mixCurr As Long

Dim rStep As Double
Dim gStep As Double
Dim bStep As Double

Dim rAmount As Double
Dim gAmount As Double
Dim bAmount As Double

r1 = Color1 Mod 256
r2 = Color2 Mod 256

g1 = (Color1 \ 256) Mod 256
g2 = (Color2 \ 256) Mod 256

b1 = (Color1 \ 256 \ 256) Mod 256
b2 = (Color2 \ 256 \ 256) Mod 256

Debug.Print "********************************"
Debug.Print "r1=" & r1 & ";r2=" & r2 & ";g1=" & g1 & ";g2=" & g2 & ";b1=" & b1 & ";b2=" & b2

rAmount = r2 - r1
gAmount = g2 - g1
bAmount = b2 - b1

Debug.Print "rAmount=" & rAmount & ";gAmount=" & gAmount & ";bAmount=" & bAmount

rStep = rAmount / picPreview.ScaleWidth
gStep = gAmount / picPreview.ScaleWidth
bStep = bAmount / picPreview.ScaleWidth

'rStep = SafeDiv(rAmount, picPreview.ScaleWidth)
'gStep = SafeDiv(gAmount, picPreview.ScaleWidth)
'bStep = SafeDiv(bAmount, picPreview.ScaleWidth)

Debug.Print "rStep=" & rStep & ";gStep=" & gStep & ";bStep=" & bStep

rCurr = r1
gCurr = g1
bCurr = b1

For i = 1 To picPreview.ScaleWidth
    rCurr = rCurr + rStep
    gCurr = gCurr + gStep
    bCurr = bCurr + bStep

    mixCurr = RGB(rCurr, gCurr, bCurr)
    
    picPreview.Line (i, 0)-(i, picPreview.ScaleHeight), mixCurr
Next i

End Sub

Public Function SafeDiv(N1 As Double, N2 As Double) As Long
On Error Resume Next

If N1 > N2 Then SafeDiv = N1 / N2 Else SafeDiv = N2 \ N1
End Function
