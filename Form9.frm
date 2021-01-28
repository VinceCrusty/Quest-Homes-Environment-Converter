VERSION 5.00
Begin VB.Form Form9 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0025221F&
   BorderStyle     =   0  'Kein
   Caption         =   "Form9"
   ClientHeight    =   2610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   ClipControls    =   0   'False
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture6 
      Appearance      =   0  '2D
      BackColor       =   &H0025221F&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   240
      Picture         =   "Form9.frx":0000
      ScaleHeight     =   1095
      ScaleWidth      =   975
      TabIndex        =   9
      Top             =   1320
      Width           =   972
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   5520
      Picture         =   "Form9.frx":4102
      ScaleHeight     =   1095
      ScaleWidth      =   1215
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00404040&
      Height          =   2805
      Left            =   6840
      ScaleHeight     =   2805
      ScaleWidth      =   45
      TabIndex        =   5
      Top             =   0
      Width           =   50
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00404040&
      Height          =   2925
      Left            =   0
      ScaleHeight     =   2925
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
      Top             =   2565
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
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   1920
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   873
      Caption         =   "START"
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
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H0025221F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   1000
      Width           =   6375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H0025221F&
      Caption         =   "© Vince Crusty for Quest Homes Discord"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   660
      Width           =   6375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H0025221F&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Environment Converter v1.9.7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private storedx As Integer
Private storedy As Integer

Private Sub Command4_Click()

Me.hide

End Sub

Private Sub Form_Load()

On Error Resume Next
'Picture5.ScaleMode = 3
'Picture5.AutoRedraw = True
'Picture5.PaintPicture Picture5.Picture, _
'0, 0, Picture5.ScaleWidth, Picture5.ScaleHeight, _
'0, 0, Picture5.Picture.Width / 21.2, _
'Picture5.Picture.Height / 21.2
'Picture5.Picture = Picture5.Image
Label3.Caption = "Special thanks goes to Elin from Quest Homes Discord" & vbNewLine
Label3.Caption = Label3.Caption & "for many hours of alpha testing, many" & vbNewLine
Label3.Caption = Label3.Caption & "suggestions and the wonderful tutorial"
Me.Top = (Form1.Top + (Form1.Height / 2) - (Form3.Height / 2))
Me.Left = (Form1.Left + (Form1.Width / 2) - (Form3.Width / 2))
Command4.HoverBackColor = Form1.Command4.HoverBackColor
Command4.HoverForeColor = Form1.Command4.HoverForeColor
Command4.ForeColor = Form1.Command4.ForeColor

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

Private Sub Picture6_DblClick()

Form2.Show (vbModal)

End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

storedx = x
storedy = y

End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then
    Me.Left = x - storedx + Me.Left
    Me.Top = y - storedy + Me.Top
End If

End Sub

Private Sub Picture6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

storedx = x
storedy = y

End Sub

Private Sub Picture6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then
    Me.Left = x - storedx + Me.Left
    Me.Top = y - storedy + Me.Top
End If

End Sub
