VERSION 5.00
Begin VB.UserControl DropList 
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2790
   ControlContainer=   -1  'True
   ForwardFocus    =   -1  'True
   ScaleHeight     =   660
   ScaleWidth      =   2790
   ToolboxBitmap   =   "DropList.ctx":0000
   Begin VB.ComboBox cbo 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   90
      Style           =   2  'Dropdown-Liste
      TabIndex        =   0
      Top             =   135
      Width           =   2055
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2250
      Top             =   105
   End
End
Attribute VB_Name = "DropList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *** Den Artikel zu diesem Modul finden Sie unter http://www.aboutvb.de/kom/artikel/komdroplist.htm ***

Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Private mStartFocus As Long

Public Event Closed()

Private pEnabled As Boolean

Public Property Get Combo() As Object
    Set Combo = cbo
End Property

Public Property Get Enabled() As Boolean
    Enabled = pEnabled
End Property

Public Property Let Enabled(New_Enabled As Boolean)
    pEnabled = New_Enabled
    If pEnabled = False Then
        Me.DropDown False
    End If
End Property

Public Property Get IsDropped() As Boolean
    Const CB_GETDROPPEDSTATE = &H157
    
    IsDropped = CBool(SendMessage(cbo.hwnd, CB_GETDROPPEDSTATE, 0, 0))
End Property

Public Sub DropDown(Optional ByVal ShowHide As Boolean = True)
    Const CB_SHOWDROPDOWN = &H14F
  
    If ShowHide Then
        If Not pEnabled Then
            Exit Sub
        End If
        mStartFocus = GetFocus()
    End If
    SendMessage cbo.hwnd, CB_SHOWDROPDOWN, ShowHide, 0
    If ShowHide Then
        UserControl.Enabled = True
        'cbo.SetFocus
        tmr.Enabled = True
    End If
End Sub

Private Sub tmr_Timer()
    If Not Me.IsDropped Then
        SetFocusAPI mStartFocus
        tmr.Enabled = False
        UserControl.Enabled = False
        RaiseEvent Closed
    End If
End Sub

Private Sub UserControl_Initialize()
    pEnabled = True
End Sub

Private Sub UserControl_InitProperties()
    If Ambient.UserMode Then
        With UserControl
            .BackStyle = 0
            .Enabled = False
        End With
    Else
        With cbo
            .AddItem Ambient.DisplayName
            .ListIndex = 0
            .ForeColor = vbHighlightText
            .BackColor = vbHighlight
        End With
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl_InitProperties
    pEnabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Enabled", pEnabled, True
End Sub

Private Sub UserControl_Resize()
    Static sInProc As Boolean
        
    If sInProc Then
        Exit Sub
    Else
        sInProc = True
    End If
    If Ambient.UserMode Then
        UserControl.Height = 0
        With cbo
            .Move 0, -.Height, UserControl.ScaleWidth
        End With
    Else
        With UserControl
            .Height = cbo.Height
            cbo.Move 0, 0, .ScaleWidth
        End With
    End If
    sInProc = False
End Sub

