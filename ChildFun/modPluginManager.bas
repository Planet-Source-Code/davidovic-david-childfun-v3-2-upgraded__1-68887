Attribute VB_Name = "modPluginManager"
Dim ActivePlugin As String ' This variable contains active plugin name

' This constant contains names of all plugins seperated by ";". If you want to add plugin,
' just add ';<your plugin name>'. Then, create appropriate Case commands in subs OnDrawBrush,
' OnImageSpray,OnChangeColor,OnStartDragging,OnStopDragging if you need. Plugin "NEW" is
' here to just show how to add 'Case' commands in subs and it doesn't do anything.
Public Const InstalledPlugins As String = "Brush random colors;Brush gradient colors;Mirror function;Mirror function horizontal;Effects;Brush random size"

' When you add new plugin, you need to increment this number
Public Const PluginCount As Integer = 6

Dim brc_OldPrim As Long
Dim brc_OldSec As Long
Dim brs_RC As Boolean
Dim brs_Max As Long
Dim bgr_BrushSeq() As Long
Dim Decrement As Boolean
Dim bgr_CurrClr As Integer
Dim e_IsPickMode As Boolean
Dim e_PickStep As Integer
Dim e_IsBoundSelected As Boolean
Dim e_XBound As Long
Dim e_YBound As Long
Public Sub internal_SwitchPlugin(NewPluginName As String)
OnDeactivatePlugin
ActivePlugin = NewPluginName
OnActivatePlugin
End Sub

Public Sub OnDrawBrush(X As Long, Y As Long, PrimaryColor As Long, SecondaryColor As Long, Size As Long, IsSecColorTurn As Boolean, Shift As Integer, Skip As Boolean)
Dim nCurrentCol As Long

Select Case ActivePlugin
    Case "Brush random colors"
        frmMain.BrushSpray CSng(X), CSng(Y), CInt(Size), RGB(Rnd * 255, Rnd * 255, Rnd * 255), frmMain.picCanvas
        Skip = True
    Case "Brush random size"
        frmMain.BrushSpray CSng(X), CSng(Y), CInt(Rnd * (brs_Max - 1)) + 1, IIf(brs_RC = True, RGB(Rnd * 255, Rnd * 255, Rnd * 255), IIf(Not frmMain.Switch, frmMain.PrimaryColor, frmMain.SecondaryColor)), frmMain.picCanvas
        frmMain.Switch = Not frmMain.Switch
        Skip = True
    Case "Brush gradient colors"
        If Not Decrement Then bgr_CurrClr = bgr_CurrClr + 1 Else bgr_CurrClr = bgr_CurrClr + -1
        If bgr_CurrClr >= UBound(bgr_BrushSeq) - 1 Then Decrement = True
        If bgr_CurrClr <= 1 Then Decrement = False
        
         nCurrentCol = bgr_BrushSeq(bgr_CurrClr)
        
        frmMain.BrushSpray CSng(X), CSng(Y), frmMain.Size, nCurrentCol, frmMain.picCanvas
        Skip = True
    Case "Effects"
        If e_IsPickMode Then Skip = True
End Select
End Sub

Public Sub OnImageSpray(ImageIndex As Integer, X As Long, Y As Long, SprayPicture As IPictureDisp, MaskPicture As IPictureDisp, Shift As Integer, Skip As Boolean)
Select Case ActivePlugin
    Case "Effects"
        If e_IsPickMode Then Skip = True
End Select
End Sub

Public Sub OnChangeColor(IsPrimaryColorChanged As Boolean, OldColor As Long, NewColor As Long, Skip As Boolean)

End Sub

Public Sub OnStartDragging(X As Long, Y As Long, Shift As Integer)

End Sub

Public Sub OnStopDragging(X As Long, Y As Long, Shift As Integer)

End Sub

Public Sub OnCanvasClick()

End Sub

Public Sub OnCanvasMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case ActivePlugin
    Case "Effects"
        Debug.Print e_IsPickMode
        If e_IsPickMode = True Then
            e_IsBoundSelected = True: e_XBound = CLng(X): e_YBound = CLng(Y)
        End If
End Select
End Sub

Public Sub OnCanvasMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Public Sub OnCanvasMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Sub Main()
frmMain.Show
End Sub

Sub OnDeactivatePlugin()
 Select Case ActivePlugin
    Case "Brush random colors"
'        frmMain.PrimaryColor = brc_OldPrim
'        frmMain.SecondaryColor = brc_OldSec
End Select
End Sub

Sub OnActivatePlugin()
 Select Case ActivePlugin
    Case "Mirror function"
        For i = 1 To frmMain.picCanvas.ScaleWidth \ 2
            For j = 1 To frmMain.picCanvas.ScaleHeight
                nr = frmMain.picCanvas.Point(i, j)
                
                pnm = frmMain.picCanvas.ScaleWidth - i
                
                frmMain.picCanvas.PSet (pnm, j), nr
            Next j
            DoEvents
        Next i
    Case "Mirror function horizontal"
        For j = 1 To frmMain.picCanvas.ScaleHeight \ 2
            For i = 1 To frmMain.picCanvas.ScaleWidth
                nr = frmMain.picCanvas.Point(i, j)
                
                pnm = frmMain.picCanvas.ScaleHeight - (j)
                
                frmMain.picCanvas.PSet (i, pnm), nr
            Next i
            DoEvents
        Next j
    Case "Brush random size"
        frmBRSSettings.Show vbModal
    Case "Brush gradient colors"
        frmBGRSettings.Show vbModal
    Case "Effects"
        frmESettings.Show vbModal
End Select
End Sub

Sub brush_random_size_internal_OnSettOK()
On Error Resume Next

brs_RC = frmBRSSettings.chkRandomColors.Value
brs_Max = Val(frmBRSSettings.txtMax.Text)

If brs_Max = 0 Then brs_Max = 2

Unload frmBRSSettings
End Sub

Sub brush_gradient_colors_internal_OnSettOK()
Erase bgr_BrushSeq
ReDim bgr_BrushSeq(1 To frmBGRSettings.picPreview.ScaleWidth)

For i = 1 To frmBGRSettings.picPreview.ScaleWidth
    bgr_BrushSeq(i) = frmBGRSettings.picPreview.Point(i, 0)
Next i

Unload frmBGRSettings

bgr_CurrClr = 1
Decrement = False

End Sub

Sub effects_internal_OnSettOK()
Dim Effect As String
Dim PickBounds As Boolean
Dim UpLeftX As Long
Dim DownRightX As Long
Dim UpLeftY As Long
Dim DownRightY As Long

Effect = frmESettings.cboEffect
PickBounds = frmESettings.chkBounds.Value = 1

Unload frmESettings

If PickBounds = False Then UpLeftX = 0: UpLeftY = 0: DownRightX = frmMain.picCanvas.ScaleWidth: DownRightY = frmMain.picCanvas.ScaleHeight: GoTo Skippy

e_IsPickMode = True
MsgBox "Click OK and then click the upper left bound of area to be applied with effect.", vbInformation

Do While Not e_IsBoundSelected = True
    DoEvents
Loop

UpLeftX = e_XBound
UpLeftY = e_YBound

e_IsBoundSelected = False

MsgBox "Click OK and then click the down right bound of area to be applied with effect.", vbInformation

Do While Not e_IsBoundSelected = True
    DoEvents
Loop

DownRightX = e_XBound
DownRightY = e_YBound

e_IsBoundSelected = False

e_IsPickMode = False

Skippy:

Select Case Effect
    Case "XOR Effect"
        MsgBox "XOR Effect has special color choosing. After clicking OK, pick up desired color.", vbExclamation
        frmMain.cdlColor.ShowColor
        pxclr = frmMain.cdlColor.Color
        
        For i = UpLeftX To DownRightX
            For j = UpLeftY To DownRightY
                cpoint = frmMain.picCanvas.Point(i, j)
                newcpoint = cpoint Xor pxclr
                
                frmMain.picCanvas.PSet (i, j), newcpoint
            Next j
            DoEvents
        Next i
    Case "Invert colors"
        For i = UpLeftX To DownRightX
            For j = UpLeftY To DownRightY
                cpoint = frmMain.picCanvas.Point(i, j)
                newcpoint = Not cpoint
                
                frmMain.picCanvas.PSet (i, j), newcpoint
            Next j
            DoEvents
        Next i
    Case "Cheapy TV"
        For i = UpLeftX To DownRightX
            For j = UpLeftY To DownRightY
                cpoint = frmMain.picCanvas.Point(i, j)
                
                r = cpoint Mod 256
                g = (cpoint \ 256) Mod 256
                b = (cpoint \ 256 \ 256) Mod 256
                
                If Rnd * 1000 < 500 Then
                    r2 = r \ 2
                    g2 = g \ 2
                    b2 = b \ 2
                Else
                    r2 = (r + 255) \ 2
                    g2 = (g + 255) \ 2
                    b2 = (b + 255) \ 2
                End If
                
                frmMain.picCanvas.PSet (i, j), RGB(r2, g2, b2)
            Next j
            DoEvents
        Next i
    Case "Interleave"
        For i = UpLeftX To DownRightX Step 2
            frmMain.picCanvas.Line (i, UpLeftY)-(i, DownRightY), vbWhite
        Next i
    Case "Interleave Horizontal"
        For i = UpLeftY To DownRightY Step 2
            frmMain.picCanvas.Line (UpLeftX, i)-(DownRightX, i), vbWhite
        Next i
End Select

End Sub
