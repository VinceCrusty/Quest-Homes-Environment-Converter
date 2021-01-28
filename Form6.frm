VERSION 5.00
Begin VB.Form Form6 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0025221F&
   BorderStyle     =   0  'Kein
   Caption         =   "Form6"
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   ClipControls    =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer4 
      Interval        =   400
      Left            =   360
      Top             =   1800
   End
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
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   1080
      Width           =   5295
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
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   1560
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   873
      Caption         =   "SAVE"
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
      BackColor       =   &H0025221F&
      Caption         =   "Filename:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H0025221F&
      Caption         =   "Please enter Pure APK-Filename without WinterLodge/ClasicHome etc."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   660
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private storedx As Integer
Private storedy As Integer

Private Sub Command4_Click()

On Error Resume Next

Text2.Text = Trim(Text2.Text)
If Text2.Text = "" Then Text2.Text = "untitled"
idr2 = Text2.Text
Unload Me

End Sub

Private Sub Form_Load()

'On Error Resume Next

Dim idt As String

Me.Top = (Form1.Top + (Form1.Height / 2) - (Form3.Height / 2))
Me.Left = (Form1.Left + (Form1.Width / 2) - (Form3.Width / 2))
Command4.HoverBackColor = Form1.Command4.HoverBackColor
Command4.HoverForeColor = Form1.Command4.HoverForeColor
Command4.ForeColor = Form1.Command4.ForeColor

If renunp = False Then
   Text2.Text = idr2
Else
   idr2 = Replace(Form1.Label8.Caption, "winterlodge", "", 1, , vbTextCompare)
   idr2 = Replace(idr2, "winter.lodge", "", 1, , vbTextCompare)
   idr2 = Replace(idr2, "winter lodge", "", 1, , vbTextCompare)
   idr2 = Replace(idr2, "winter_lodge", "", 1, , vbTextCompare)
   idr2 = Replace(idr2, "spacestation", "", 1, , vbTextCompare)
   idr2 = Replace(idr2, "space station", "", 1, , vbTextCompare)
   idr2 = Replace(idr2, "space.station", "", 1, , vbTextCompare)
   idr2 = Replace(idr2, "space_station", "", 1, , vbTextCompare)
   idr2 = Replace(idr2, "classichome", "", 1, , vbTextCompare)
   idr2 = Replace(idr2, "classic home", "", 1, , vbTextCompare)
   idr2 = Replace(idr2, "classic_home", "", 1, , vbTextCompare)
   idr2 = Replace(idr2, "classic.home", "", 1, , vbTextCompare)
   idr2 = Replace(idr2, ".apk", "", 1, , vbTextCompare)
   If Mid$(idr2, Len(idr2), 1) = "." Then idr2 = Left$(idr2, Len(idr2) - 1)
   Text2.Text = idr2
End If
Beep
If Mak = True Then Call Command4_Click

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

Private Sub Timer4_Timer()

Text2.SetFocus
Text2.SelStart = Len(Text2.Text)
Timer4.Enabled = False

End Sub
