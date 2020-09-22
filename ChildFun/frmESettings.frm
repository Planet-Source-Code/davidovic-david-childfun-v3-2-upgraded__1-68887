VERSION 5.00
Begin VB.Form frmESettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Effects plugin settings"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Frame fraExample 
      Caption         =   "Example"
      Height          =   4215
      Left            =   480
      TabIndex        =   4
      Top             =   1800
      Width           =   4455
      Begin VB.Image imgAfter 
         BorderStyle     =   1  'Fixed Single
         Height          =   1335
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label lblAfter 
         Caption         =   "After:"
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   2160
         Width           =   495
      End
      Begin VB.Image imgBefore 
         BorderStyle     =   1  'Fixed Single
         Height          =   1335
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblBefore 
         Caption         =   "Before:"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fraEffect 
      Caption         =   "Effect"
      Height          =   1575
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox chkBounds 
         Caption         =   "Allow me to choose area the effect will be applied to."
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   4215
      End
      Begin VB.ComboBox cboEffect 
         Height          =   315
         ItemData        =   "frmESettings.frx":0000
         Left            =   120
         List            =   "frmESettings.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label lblEffect 
         Caption         =   "Select an effect to apply:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmESettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboEffect_Click()
On Error Resume Next

imgBefore.Picture = LoadPicture(App.Path & "\plugins\effects_plugin_data\" & cboEffect.Text & " Before.bmp")
imgAfter.Picture = LoadPicture(App.Path & "\plugins\effects_plugin_data\" & cboEffect.Text & " After.bmp")
End Sub

Private Sub cmdOK_Click()
If cboEffect.Text = "" Then cboEffect.SetFocus: Exit Sub
effects_internal_OnSettOK
End Sub

