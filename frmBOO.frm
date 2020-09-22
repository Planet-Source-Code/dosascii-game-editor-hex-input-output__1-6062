VERSION 5.00
Begin VB.Form frmBOO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Boo!"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   2910
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Raaaaahhhhhh!"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmBOO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
'Close this form or dialpg box, and return
'focus to the main frame.
Unload Me
End Sub

