VERSION 5.00
Begin VB.Form Form8 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0025221F&
   BorderStyle     =   0  'Kein
   Caption         =   "Form8"
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   ClipControls    =   0   'False
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00404040&
      Height          =   2205
      Left            =   6840
      ScaleHeight     =   2205
      ScaleWidth      =   45
      TabIndex        =   5
      Top             =   0
      Width           =   50
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00404040&
      Height          =   2205
      Left            =   0
      ScaleHeight     =   2205
      ScaleWidth      =   45
      TabIndex        =   4
      Top             =   0
      Width           =   50
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00404040&
      Height          =   50
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   6975
      TabIndex        =   3
      Top             =   2200
      Width           =   6975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   50
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   6975
      TabIndex        =   2
      Top             =   0
      Width           =   6975
   End
   Begin Projekt1.lvButtons_H Command4 
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   1440
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1085
      Caption         =   "Yes"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   16777215
      cBhover         =   4194368
      Focus           =   0   'False
      LockHover       =   1
      cGradient       =   4210752
      Mode            =   0
      Value           =   0   'False
      cBack           =   4210752
   End
   Begin Projekt1.lvButtons_H lvButtons_H1 
      Height          =   615
      Left            =   3720
      TabIndex        =   6
      Top             =   1440
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1085
      Caption         =   "No"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   16777215
      cBhover         =   4194368
      Focus           =   0   'False
      LockHover       =   1
      cGradient       =   4210752
      Mode            =   0
      Value           =   0   'False
      cBack           =   4210752
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H0025221F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   6375
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private storedx As Integer
Private storedy As Integer

Private Sub Command4_Click()

answer = True
Me.Hide

End Sub

Private Sub Form_Load()

On Error Resume Next

Me.Top = (Form1.Top + (Form1.Height / 2) - (Form3.Height / 2))
Me.Left = (Form1.Left + (Form1.Width / 2) - (Form3.Width / 2))
Command4.HoverBackColor = Form1.Command4.HoverBackColor
Command4.HoverForeColor = Form1.Command4.HoverForeColor
Command4.ForeColor = Form1.Command4.ForeColor
lvButtons_H1.HoverBackColor = Form1.Command4.HoverBackColor
lvButtons_H1.HoverForeColor = Form1.Command4.HoverForeColor
lvButtons_H1.ForeColor = Form1.Command4.ForeColor

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

storedx = x
storedy = y

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then
    Me.Left = x - storedx + Me.Left
    Me.Top = y - storedy + Me.Top
End If

End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

storedx = x
storedy = y

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then
    Me.Left = x - storedx + Me.Left
    Me.Top = y - storedy + Me.Top
End If

End Sub


Private Sub lvButtons_H1_Click()

answer = False
Me.Hide

End Sub

