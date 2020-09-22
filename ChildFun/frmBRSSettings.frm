VERSION 5.00
Begin VB.Form frmBRSSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Brush random size plugin settings"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMax 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Text            =   "70"
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CheckBox chkRandomColors 
      Caption         =   "Random colors"
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblMax 
      Caption         =   "Maximum size:"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "frmBRSSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
brush_random_size_internal_OnSettOK
End Sub

Private Sub Text1_Change()

End Sub

