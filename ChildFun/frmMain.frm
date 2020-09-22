VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "ChildFun"
   ClientHeight    =   10275
   ClientLeft      =   435
   ClientTop       =   1950
   ClientWidth     =   15240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10275
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSprayer 
      Height          =   855
      Index           =   5
      Left            =   3480
      Picture         =   "frmMain.frx":0ABA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdSprayer 
      Height          =   855
      HelpContextID   =   4
      Index           =   4
      Left            =   2640
      Picture         =   "frmMain.frx":3B38
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdSprayer 
      Height          =   855
      Index           =   3
      Left            =   1800
      Picture         =   "frmMain.frx":6712
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdSprayer 
      Caption         =   " "
      Height          =   855
      Index           =   2
      Left            =   960
      Picture         =   "frmMain.frx":9B20
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdSprayer 
      Caption         =   " "
      Height          =   855
      Index           =   1
      Left            =   120
      Picture         =   "frmMain.frx":CF2E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdRight 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   14400
      Picture         =   "frmMain.frx":11C50
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdPlugin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "no plugin"
      Height          =   855
      Left            =   12240
      Picture         =   "frmMain.frx":127BA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdLeft 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   11400
      Picture         =   "frmMain.frx":12F4C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   240
      Width           =   855
   End
   Begin MSComDlg.CommonDialog cdlColor 
      Left            =   5880
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select color"
   End
   Begin VB.PictureBox picSprayer 
      Height          =   1095
      Left            =   7800
      ScaleHeight     =   1035
      ScaleWidth      =   1395
      TabIndex        =   11
      Top             =   240
      Width           =   1455
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   8775
      Left            =   0
      ScaleHeight     =   581
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1021
      TabIndex        =   10
      Top             =   1440
      Width           =   15375
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   9480
      TabIndex        =   9
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   9480
      TabIndex        =   8
      Top             =   360
      Width           =   1695
   End
   Begin VB.PictureBox picBrushPreview 
      Height          =   855
      Left            =   4680
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.HScrollBar hscSize 
      Height          =   255
      LargeChange     =   5
      Left            =   5760
      Max             =   700
      Min             =   1
      TabIndex        =   5
      Top             =   840
      Value           =   1
      Width           =   1815
   End
   Begin VB.Image imgSprRight 
      Height          =   240
      Left            =   4200
      Picture         =   "frmMain.frx":13AB6
      Top             =   120
      Width           =   150
   End
   Begin VB.Image imgSprLeft 
      Height          =   240
      Left            =   120
      Picture         =   "frmMain.frx":13CF8
      Top             =   120
      Width           =   150
   End
   Begin VB.Label lblSize 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Size:"
      Height          =   255
      Left            =   5880
      TabIndex        =   7
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblSprayer 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   3975
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "POPUP"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupPrimaryColor 
         Caption         =   "Select primary color..."
      End
      Begin VB.Menu mnuPopupSecondaryColor 
         Caption         =   "Select secondary color..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' It was a simple program for drawing.
'
' I am ten. I was maked this program for my little brother and I got an idea to submit it to Planet-Source-Code.
'
'
' It has an brush function. Click on brush preview picturebox and brush will be activated. Change size by scrrolbar.
' You can draw with random, single or two colors that are switching when you move the mouse. To use single color,
' select primary and secondary colors be same. Set primary/secondary color:

' Right click on cnavas, select "Select primary" or "Select secondary color" and pick up shelled color.
' Click Random colors to spray random colors on canvas.

' It has an image sprayer function. Click on an image and spray on cnavas.

'"Clear" to clear canvas,  "Save" to save image.

' If You like this code = True Then
'   Call Vote_For_Geomaster
'   Enjoy
' Else
'   Enjoy
' End If

Dim Plugins() As String

Const ShellType = "bmp"
Const TopConst = 1440
Const SaveType = "bmp"

Public SprayerCount As Integer
Public CurrentSprayer As Integer
Public CurrentPlugin As Integer
Public Brush As Boolean
Public CurrentSpray As Integer
Public Size As Integer
Public Random As Boolean
Public PrimaryColor As Long
Public SecondaryColor As Long
Public Dragging As Boolean
Public Switch As Boolean

'Private Sub chkRandom_Click()
'Random = chkRandom.Value ' Set the checkbox value to be in variable
'End Sub



Private Sub cmdClear_Click()
picCanvas.Cls 'Clear the canvas
End Sub

Private Sub cmdLeft_Click()
CurrentPlugin = CurrentPlugin - 1 ' Decrement plugin counter
If CurrentPlugin < 0 Then CurrentPlugin = PluginCount ' If counter is smaller than 0, set to count from last plugin

cmdPlugin.Caption = Plugins(CurrentPlugin) ' Write plugin name

pPluginName = IIf(Dir(App.Path & "\plugins\" & Plugins(CurrentPlugin) & ".bmp") = "", "unknown", Plugins(CurrentPlugin))
cmdPlugin.Picture = LoadPicture(App.Path & "\plugins\" & pPluginName & ".bmp") ' Set plugin icon

End Sub

Private Sub cmdPlugin_Click()
internal_SwitchPlugin Plugins(CurrentPlugin) ' Trigger switching event in plugin manager
End Sub

Private Sub cmdRight_Click()
CurrentPlugin = CurrentPlugin + 1
If CurrentPlugin > PluginCount Then CurrentPlugin = 0

cmdPlugin.Caption = Plugins(CurrentPlugin)

pPluginName = IIf(Dir(App.Path & "\plugins\" & Plugins(CurrentPlugin) & ".bmp") = "", "unknown", Plugins(CurrentPlugin))
cmdPlugin.Picture = LoadPicture(App.Path & "\plugins\" & pPluginName & ".bmp")

End Sub

Private Sub cmdSave_Click()
SavePicture picCanvas.Image, App.Path & "\Saved images\Image captured " & Replace(Date, "/", "-") & " at " & Replace(Time, ":", "-") & "." & SaveType ' Save the picture. Filename has Time and Date stamp.
cmdClear_Click
End Sub

Private Sub cmdSprayer_Click(Index As Integer)
On Error Resume Next

'Set SprayerPict = LoadPicture(App.Path & "\sprayer\" & Index & "." & ShellType) ' Load an appropriate image for sprayer
'Set MaskPict = LoadPicture(App.Path & "\sprayer\" & Index & "_mask." & ShellType) ' Load an appropriate mask for sprayer
Dim rnFName As String

For i = 1 To 5
    rnFName = IIf(Dir(App.Path & "\sprayer\" & AllSprayers(CurrentSprayer) & "\" & Index & "_" & i & "." & ShellType) = "", App.Path & "\sprayer\" & AllSprayers(CurrentSprayer) & "\" & Index & "_1." & ShellType, App.Path & "\sprayer\" & AllSprayers(CurrentSprayer) & "\" & Index & "_" & i & "." & ShellType)
    Set SprayerPict(i) = LoadPicture(rnFName)
    
    rnFName = IIf(Dir(App.Path & "\sprayer\" & AllSprayers(CurrentSprayer) & "\" & Index & "_" & i & "_mask." & ShellType) = "", App.Path & "\sprayer\" & AllSprayers(CurrentSprayer) & "\" & Index & "_1_mask." & ShellType, App.Path & "\sprayer\" & AllSprayers(CurrentSprayer) & "\" & Index & "_" & i & "_mask." & ShellType)
    Set MaskPict(i) = LoadPicture(rnFName)
Next i

picSprayer.Picture = SprayerPict(1) ' Show the user current sprayer picture, so he always see current picture
Brush = False ' Turn off Brush mode
End Sub

Private Sub Form_Load()
On Error Resume Next

Open App.Path & "\sprayer\sprayers.cf" For Input As #1
    Line Input #1, strnam
    ReDim AllSprayers(1 To CInt(strnam))

    For i = 1 To CInt(strnam)
        Line Input #1, strData
        AllSprayers(i) = strData
    Next i
Close #1

SprayerCount = CInt(strnam)

ReDim Plugins(0 To PluginCount)

For i = 0 To PluginCount - 1
    Plugins(i + 1) = Split(InstalledPlugins, ";")(i)
Next i

Plugins(0) = "no plugin"

For i = 1 To 5
    cmdSprayer(i).Picture = LoadPicture(App.Path & "\sprayer\Default\" & i & "_1." & ShellType) ' Set appropriate image to all sprayer commandbuttons (default imagesprayer set)
Next i

cmdRight_Click
cmdLeft_Click

imgSprRight_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next

picCanvas.Move 120, TopConst, Me.Width - 120 * 3, Me.Height - 120 * 17 ' Align canvas
'If Me.WindowState = vbMaximized Then picCanvas.Move 0, TopConst, Screen.Width, Screen.Height ' I had some problems with aligning when window is maximized (width and height aren't showing real values) so we will align to Screen.Width and Height
End Sub

Private Sub mnuPopupRandom_Click()

End Sub

Private Sub imgSprLeft_Click()
On Error Resume Next

CurrentSprayer = CurrentSprayer - 1
If CurrentSprayer < 1 Then CurrentSprayer = SprayerCount

For i = 1 To 5
    cmdSprayer(i).Picture = LoadPicture(App.Path & "\sprayer\" & AllSprayers(CurrentSprayer) & "\" & i & "_1.bmp")
Next i

lblSprayer.Caption = AllSprayers(CurrentSprayer)
End Sub

Private Sub imgSprRight_Click()
On Error Resume Next

CurrentSprayer = CurrentSprayer + 1
If CurrentSprayer > SprayerCount Then CurrentSprayer = 1

Debug.Print CurrentSprayer

For i = 1 To 5
    cmdSprayer(i).Picture = LoadPicture(App.Path & "\sprayer\" & AllSprayers(CurrentSprayer) & "\" & i & "_1.bmp")
Next i


    
lblSprayer.Caption = AllSprayers(CurrentSprayer)
End Sub

Private Sub picBrushPreview_Paint()
hscSize_Change
End Sub

'Private Sub mnuPopupRandom_Click()
'mnuPopupRandom.Checked = Not (mnuPopupRandom.Checked) ' Invert the checked value
'chkRandom.Value = IIf(mnuPopupRandom.Checked, 1, 0) ' Set the checkbox value to our menu check value
'chkRandom_Click ' Trigger the click event to move checked value into variable
'End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
OnCanvasMouseDown Button, Shift, CInt(X), CInt(Y)

If Button = vbRightButton Then
    PopupMenu mnuPopup, , X * 15, Y * 15 + TopConst ' Popup menu if user clicks right button
    Exit Sub ' Break sub, so it will not draw object
End If

DrawShelledObject X, Y ' Draw the object on X and Y
Dragging = True ' Set global dragging variable to true
OnStartDragging CInt(X), CInt(Y), Shift
End Sub

Sub SprayPict(X As Single, Y As Single, Picture As IPictureDisp, MaskPict As IPictureDisp, ByRef pic As PictureBox)
On Error GoTo handle_errormessage
pic.PaintPicture MaskPict, X, Y, , , , , , , vbMergePaint ' Paint mask on canvas
pic.PaintPicture Picture, X, Y, , , , , , , vbSrcAnd ' Paint picture on canvas

handle_errormessage:
Exit Sub
End Sub

Sub BrushSpray(X As Single, Y As Single, Size As Integer, Color As Long, pic As PictureBox)
On Error Resume Next
Dim BeforeDWidth As Integer

BeforeDWidth = pic.DrawWidth ' Set this variable to current drawwidth

pic.DrawWidth = Size ' Set drawwidth to shelled size
pic.PSet (X, Y), Color ' Spray the dot of size and color

pic.DrawWidth = BeforeDWidth ' Return the previous drawwidth

End Sub

Sub DrawShelledObject(X As Single, Y As Single)
Dim pX As Long
Dim pY As Long
Dim bSkip As Boolean

pX = X
pY = Y

If Brush Then
    OnDrawBrush pX, pY, PrimaryColor, SecondaryColor, CLng(Size), Switch, 0, bSkip
    If bSkip = True Then Exit Sub
    
    BrushSpray X, Y, Size, PrimaryColor, picCanvas ' Spray brush
    If Switch Then BrushSpray X, Y, Size, SecondaryColor, picCanvas ' If is turn when drawing in secondary color, just overwrite with secondary color
    If Random Then BrushSpray X, Y, Size, RGB(Rnd * 255, Rnd * 255, Rnd * 255), picCanvas ' If random colors, overwrite with random color
    Switch = Not (Switch) ' Invert switch ( change turn )
Else
    CurrentSpray = CurrentSpray + 1
    If CurrentSpray > 5 Then CurrentSpray = 1
    
    OnImageSpray 0, pX, pY, SprayerPict(CurrentSpray), MaskPict(CurrentSpray), 0, bSkip
    If bSkip = True Then Exit Sub
    
    SprayPict X, Y, SprayerPict(CurrentSpray), MaskPict(CurrentSpray), picCanvas ' If we have Image Sprayer, spray picture on coords
End If

End Sub

Private Sub hscSize_Change()
picBrushPreview.Cls ' Clear brush preview
BrushSpray picBrushPreview.Width / 2, picBrushPreview.Height / 2, hscSize.Value, vbBlack, picBrushPreview ' Draw brush preview on picBrushPreview
Size = hscSize.Value ' Change size
End Sub

Private Sub mnuPopupPrimaryColor_Click()
cdlColor.Color = PrimaryColor ' Set the default color to old primary color
cdlColor.ShowColor ' Show color select dialog
PrimaryColor = cdlColor.Color ' Set primary color to selected color
End Sub

Private Sub mnuPopupSecondaryColor_Click()
cdlColor.Color = SecondaryColor '  Same thing, just with secondary color
cdlColor.ShowColor
SecondaryColor = cdlColor.Color
End Sub

Private Sub picBrushPreview_Click()
Brush = True ' Activate brush
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
OnCanvasMouseMove Button, Shift, CInt(X), CInt(Y)

If Dragging Then
    DrawShelledObject X, Y ' If dragging, draw selected object on coords
End If

End Sub

Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
OnCanvasMouseUp Button, Shift, CInt(X), CInt(Y)

Dragging = False ' Turn off dragging
OnStopDragging CInt(X), CInt(Y), Shift
End Sub


