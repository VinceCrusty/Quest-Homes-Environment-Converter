VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'Kein
   Caption         =   "About"
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5640
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'Kein
      Height          =   3975
      Left            =   840
      Picture         =   "Form2.frx":838B
      ScaleHeight     =   3975
      ScaleWidth      =   3975
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "(c) Vince Crusty 2020 for Quest Homes Discord"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Picture1_Click()

Form2.Hide

End Sub
