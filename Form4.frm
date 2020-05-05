VERSION 5.00
Begin VB.Form Form4 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0025221F&
   BorderStyle     =   0  'Kein
   Caption         =   "Form3"
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   ClipControls    =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
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
      Left            =   2640
      TabIndex        =   0
      Top             =   1440
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1085
      Caption         =   "SAVE"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
   Begin VB.Label Label3 
      BackColor       =   &H0025221F&
      Caption         =   "New Port"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   1080
      Width           =   960
   End
   Begin VB.Label Label2 
      BackColor       =   &H0025221F&
      Caption         =   "Old Port"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   600
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H0025221F&
      Caption         =   "Enter new ADB Wireless Port  (e.g. 5555/5557):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private storedx As Integer
Private storedy As Integer

Private Sub Command4_Click()

On Error Resume Next

Text2.Text = Trim(Text2.Text)
If Text2.Text <> Text1.Text And Text2.Text <> "" And IsNumeric(Text2.Text) = True Then
   PutINISetting "QuestIP", "Port", Text2.Text, App.path & "\files\config.ini"
End If
If IsNumeric(Text2.Text) = False Then Text2.Text = ""
Beep
Me.Hide

End Sub

Private Sub Form_Load()

On Error Resume Next

Me.Top = (Form1.Top + (Form1.Height / 2) - (Form3.Height / 2))
Me.Left = (Form1.Left + (Form1.Width / 2) - (Form3.Width / 2))
Command4.HoverBackColor = Form1.Command4.HoverBackColor
Command4.HoverForeColor = Form1.Command4.HoverForeColor
Text1.Text = GetINISetting("QuestIP", "Port", App.path & "\files\config.ini")

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

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

storedx = x
storedy = y

End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then
    Me.Left = x - storedx + Me.Left
    Me.Top = y - storedy + Me.Top
End If

End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

storedx = x
storedy = y

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then
    Me.Left = x - storedx + Me.Left
    Me.Top = y - storedy + Me.Top
End If

End Sub
