Attribute VB_Name = "Colorpicker"
Private Type ChooseColorStruct
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" _
    (lpChoosecolor As ChooseColorStruct) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor _
    As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Private Const CC_RGBINIT = &H1&
Private Const CC_FULLOPEN = &H2&
Private Const CC_PREVENTFULLOPEN = &H4&
Private Const CC_SHOWHELP = &H8&
Private Const CC_ENABLEHOOK = &H10&
Private Const CC_ENABLETEMPLATE = &H20&
Private Const CC_ENABLETEMPLATEHANDLE = &H40&
Private Const CC_SOLIDCOLOR = &H80&
Private Const CC_ANYCOLOR = &H100&
Private Const CLR_INVALID = &HFFFF


' Show the common dialog for choosing a color.
' Return the chosen color, or -1 if the dialog is canceled
'
' hParent is the handle of the parent form
' bFullOpen specifies whether the dialog will be open with the Full style
' (allows to choose many more colors)
' InitColor is the color initially selected when the dialog is open

' Example:
'    Dim oleNewColor As OLE_COLOR
'    oleNewColor = ShowColorsDialog(Me.hwnd, True, vbRed)
'    If oleNewColor <> -1 Then Me.BackColor = oleNewColor

Public Function ShowColorDialog(Optional ByVal hParent As Long, _
    Optional ByVal bFullOpen As Boolean, Optional ByVal InitColor As OLE_COLOR) _
    As Long
    Dim CC As ChooseColorStruct
    Dim aColorRef(15) As Long
    Dim lInitColor As Long

    ' translate the initial OLE color to a long value
    If InitColor <> 0 Then
        If OleTranslateColor(InitColor, 0, lInitColor) Then
            lInitColor = CLR_INVALID
        End If
    End If

    'fill the ChooseColorStruct struct
    With CC
        .lStructSize = Len(CC)
        .hwndOwner = hParent
        .lpCustColors = VarPtr(aColorRef(0))
        .rgbResult = lInitColor
        .flags = CC_SOLIDCOLOR Or CC_ANYCOLOR Or CC_RGBINIT Or IIf(bFullOpen, _
            CC_FULLOPEN, 0)
    End With

    ' Show the dialog
    If ChooseColor(CC) Then
        'if not canceled, return the color
        ShowColorDialog = CC.rgbResult
    Else
        'else return -1
        ShowColorDialog = -1
    End If
End Function

Public Function Lux(ByVal pColor As Long) As Integer

Dim iR As Integer, iG As Integer, iB As Integer
Dim nRPct As Single, nGPct As Single, nBPct As Single
Dim nMax As Single, nMin As Single
Dim nLumPct As Single

iR = pColor Mod 256
pColor = pColor \ 256
iG = pColor Mod 256
pColor = pColor \ 256
iB = pColor Mod 256
nRPct = iR / 255
nGPct = iG / 255
nBPct = iB / 255
If nRPct > nGPct And nRPct > nBPct Then
   nMax = nRPct
ElseIf nGPct > nBPct Then
   nMax = nGPct
Else
    nMax = nBPct
End If
If nRPct < nGPct And nRPct < nBPct Then
   nMin = nRPct
ElseIf nGPct < nBPct Then
   nMin = nGPct
Else
   nMin = nBPct
End If
nLumPct = (nMin + nMax) / 2
' return the luminosity min (darkest) = 0 max (lightest) = 240
Lux = nLumPct * 240
    
End Function
