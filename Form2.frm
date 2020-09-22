VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3165
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   1
      Text            =   "2000"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Number of URLs to Sipder"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Form1.MaxURL = Text1.Text
    Unload Me
    Form1.ProgressBar1.Max = Text1.Text
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Text1.Text = Form1.MaxURL
End Sub
