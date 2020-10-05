VERSION 5.00
Begin VB.Form Form7 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0025221F&
   BorderStyle     =   0  'Kein
   Caption         =   "Form7"
   ClientHeight    =   8895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8925
   ClipControls    =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer TimerPic 
      Interval        =   100
      Left            =   720
      Top             =   8160
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   4210
      Left            =   240
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   4185
      ScaleWidth      =   8385
      TabIndex        =   5
      Top             =   240
      Width           =   8420
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00404040&
      Height          =   9285
      Left            =   8880
      ScaleHeight     =   9285
      ScaleWidth      =   45
      TabIndex        =   4
      Top             =   0
      Width           =   50
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00404040&
      Height          =   9285
      Left            =   0
      ScaleHeight     =   9285
      ScaleWidth      =   45
      TabIndex        =   3
      Top             =   0
      Width           =   50
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00404040&
      Height          =   50
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   9855
      TabIndex        =   2
      Top             =   8850
      Width           =   9855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   50
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   9975
      TabIndex        =   1
      Top             =   0
      Width           =   9975
   End
   Begin Projekt1.lvButtons_H Command4 
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   8160
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   873
      Caption         =   "Build Panorama"
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
      Enabled         =   0   'False
      cBack           =   4210752
   End
   Begin Projekt1.lvButtons_H lvButtons_H1 
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   6600
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   16777215
      cFDown          =   4194368
      cBhover         =   4194368
      Focus           =   0   'False
      LockHover       =   1
      cGradient       =   4210752
      Mode            =   1
      Value           =   -1  'True
      cBack           =   8421504
   End
   Begin Projekt1.lvButtons_H lvButtons_H2 
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   6960
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   16777215
      cFDown          =   4194368
      cBhover         =   4194368
      Focus           =   0   'False
      LockHover       =   1
      cGradient       =   4210752
      Mode            =   1
      Value           =   0   'False
      cBack           =   8421504
   End
   Begin Projekt1.lvButtons_H lvButtons_H3 
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   7320
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   16777215
      cFDown          =   4194368
      cBhover         =   4194368
      Focus           =   0   'False
      LockHover       =   1
      cGradient       =   4210752
      Mode            =   1
      Value           =   0   'False
      cBack           =   8421504
   End
   Begin Projekt1.lvButtons_H lvButtons_H4 
      Height          =   255
      Left            =   7080
      TabIndex        =   13
      Top             =   6600
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   16777215
      cFDown          =   4194368
      cBhover         =   4194368
      Focus           =   0   'False
      LockHover       =   1
      cGradient       =   4210752
      Mode            =   1
      Value           =   0   'False
      cBack           =   8421504
   End
   Begin Projekt1.lvButtons_H lvButtons_H5 
      Height          =   495
      Left            =   4440
      TabIndex        =   15
      Top             =   8160
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   873
      Caption         =   "Cancel"
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
   Begin Projekt1.lvButtons_H Check1 
      Height          =   255
      Left            =   2280
      TabIndex        =   18
      Top             =   6600
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   16777215
      cFDown          =   4194368
      cBhover         =   4194368
      Focus           =   0   'False
      LockHover       =   1
      cGradient       =   4210752
      Mode            =   1
      Value           =   -1  'True
      Enabled         =   0   'False
      cBack           =   8421504
   End
   Begin Projekt1.lvButtons_H Check2 
      Height          =   255
      Left            =   2280
      TabIndex        =   19
      Top             =   6960
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   16777215
      cFDown          =   4194368
      cBhover         =   4194368
      Focus           =   0   'False
      LockHover       =   1
      cGradient       =   4210752
      Mode            =   1
      Value           =   0   'False
      cBack           =   8421504
   End
   Begin Projekt1.lvButtons_H Check3 
      Height          =   255
      Left            =   2280
      TabIndex        =   20
      Top             =   7320
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   16777215
      cFDown          =   4194368
      cBhover         =   4194368
      Focus           =   0   'False
      LockHover       =   1
      cGradient       =   4210752
      Mode            =   1
      Value           =   0   'False
      cBack           =   8421504
   End
   Begin Projekt1.lvButtons_H lvButtons_H 
      Height          =   315
      Left            =   4320
      TabIndex        =   22
      Top             =   6570
      Width           =   705
      _ExtentX        =   1296
      _ExtentY        =   556
      Caption         =   "19"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
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
      Enabled         =   0   'False
      cBack           =   4210752
   End
   Begin Projekt1.DropList DropList1 
      Height          =   315
      Left            =   4320
      TabIndex        =   25
      Top             =   6860
      Width           =   705
      _ExtentX        =   1296
      _ExtentY        =   556
   End
   Begin Projekt1.lvButtons_H Check4 
      Height          =   255
      Left            =   2280
      TabIndex        =   26
      Top             =   7680
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   16777215
      cFDown          =   4194368
      cBhover         =   4194368
      Focus           =   0   'False
      LockHover       =   1
      cGradient       =   4210752
      Mode            =   1
      Value           =   -1  'True
      Enabled         =   0   'False
      cBack           =   8421504
   End
   Begin VB.Label Label10 
      BackColor       =   &H0025221F&
      Caption         =   "Encode OGG Audio File (lower filesize)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   27
      ToolTipText     =   "Re-encode the OGG-audio file to reduce the size"
      Top             =   7680
      Width           =   3615
   End
   Begin VB.Label Label5 
      BackColor       =   &H0025221F&
      Caption         =   "Silent audio file"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   24
      ToolTipText     =   "Exchange the audio file with an empty (silent) audio file"
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H0025221F&
      Caption         =   "Default audio file (fireplace)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   23
      ToolTipText     =   "Exchange the audio file with the standard fireplace sound"
      Top             =   6960
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H0025221F&
      Caption         =   "Decrase volume by             dB (19 recommned)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   21
      ToolTipText     =   "Most audio files are too loud for an environment"
      Top             =   6600
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   840
      TabIndex        =   17
      Top             =   4560
      Width           =   7815
   End
   Begin VB.Label Label7 
      BackColor       =   &H0025221F&
      Caption         =   "Audio:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H0025221F&
      Caption         =   "Auto Install"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7440
      TabIndex        =   14
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label Label17 
      BackColor       =   &H0025221F&
      Caption         =   "SpaceStation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackColor       =   &H0025221F&
      Caption         =   "ClassicHome"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackColor       =   &H0025221F&
      Caption         =   "WinterLodge"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0025221F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   360
      TabIndex        =   6
      Top             =   4920
      Width           =   8175
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Combo5 As ComboBox
Attribute Combo5.VB_VarHelpID = -1

Private storedx As Integer
Private storedy As Integer
Private pat As String



Private Sub Form_Load()

On Error Resume Next

Dim i As Long

Set Combo5 = DropList1.Combo
For i = 1 To 32
Set Combo5 = DropList1.Combo
Combo5.AddItem (i)
Next i
Combo5.Text = "19"
rota(1) = "0.4998510181903839,0.49992236495018005,0.5001489520072937,0.5000776052474976"
rota(2) = "0.5608572363853455,0.5609186887741089,0.4306264817714691,0.4305464029312134"
rota(3) = "0.6122670769691467,0.6123175024986267,0.3537358343601227,0.35364842414855957"
rota(4) = "0.6532008647918701,0.6532394289970398,0.270792692899704,0.2706994414329529"
rota(5) = "0.68295818567276,0.6829842925071716,0.18321619927883148,0.18311870098114014"
rota(6) = "0.701029896736145,0.7010430097579956,0.09250480681657791,0.09240473806858063"
rota(7) = "0.7071067690849304,0.7071067690849304,0.0002106766332872212,0.00010974617907777429"
rota(8) = "0.7010848522186279,0.7010716795921326,-0.09208710491657257,-0.09218716621398926"
rota(9) = "0.6830672025680542,0.6830410957336426,-0.18280920386314392,-0.18290668725967407"
rota(10) = "0.6533620953559875,0.6533234715461731,-0.2704033851623535,-0.2704966366291046"
rota(11) = "0.6124777793884277,0.612427294254303,-0.35337090492248535,-0.35345831513404846"
rota(12) = "0.5611137747764587,0.5610523223876953,-0.43029215931892395,-0.4303722083568573"
rota(13) = "0.5001490116119385,0.5000776052474976,-0.4998509883880615,-0.49992233514785767"
rota(14) = "0.4306264817714691,0.430546373128891,-0.5608572959899902,-0.5609186887741089"
rota(15) = "0.3537357747554779,0.3536483645439148,-0.6122671365737915,-0.6123175621032715"
rota(16) = "0.27079257369041443,0.27069932222366333,-0.6532009243965149,-0.6532394886016846"
rota(17) = "0.18321600556373596,0.18311850726604462,-0.6829582452774048,-0.6829842925071716"
rota(18) = "0.09250465780496597,0.09240458160638809,-0.701029896736145,-0.7010430693626404"
rota(19) = "0.00021035069948993623,0.00010942023800453171,-0.7071067690849304,-0.7071067690849304"
rota(20) = "-0.09208755195140839,-0.09218761324882507,-0.7010847926139832,-0.7010716199874878"
rota(21) = "-0.18280979990959167,-0.18290729820728302,-0.6830670833587646,-0.6830409169197083"
rota(22) = "-0.2704041302204132,-0.2704973518848419,-0.6533617973327637,-0.6533231735229492"
rota(23) = "-0.35337159037590027,-0.3534590005874634,-0.6124773621559143,-0.6124268770217896"
rota(24) = "-0.43029290437698364,-0.430372953414917,-0.5611132383346558,-0.5610517859458923"
gltf1 = "{!asset!:{!generator!:!Khronos glTF Blender I/O v1.1.46!,!version!:!2.0!},!scene!:0,!scenes!:[{!name!:!Scene!,!nodes!:[0]}],!nodes!:[{!mesh!:0,!name!:!Sphere.011!,!rotation!:["
gltf2 = "],!scale!:[978.0592651367188,978.0593872070312,978.0591430664062]}],!materials!:[{!emissiveFactor!:[0,0,0],!name!:!material!,!pbrMetallicRoughness!:{!baseColorTexture!:{!inde" & _
        "x!:0,!texCoord!:0},!metallicFactor!:0,!roughnessFactor!:0.5}}],!meshes!:[{!name!:!Material2.011!,!primitives!:[{!attributes!:{!POSITION!:0,!NORMAL!:1,!TEXCOORD_0!:2},!indice" & _
        "s!:3,!material!:0}]}],!textures!:[{!source!:0}],!images!:[{!mimeType!:!image/jpeg!,!name!:!pano!,!uri!:!pano.jpg!}],!accessors!:[{!bufferView!:0,!componentType!:5126,!c" & _
        "ount!:4410,!max!:[1,0.9994649887084961,0.9994649887084961],!min!:[-1,-0.9994649887084961,-0.9994649887084961],!type!:!VEC3!},{!bufferView!:1,!componentType!:5126,!count!:441" & _
        "0,!type!:!VEC3!},{!bufferView!:2,!componentType!:5126,!count!:4410,!type!:!VEC2!},{!bufferView!:3,!componentType!:5123,!count!:24192,!type!:!SCALAR!}],!bufferViews!:[{!buffe" & _
        "r!:0,!byteLength!:52920,!byteOffset!:0},{!buffer!:0,!byteLength!:52920,!byteOffset!:52920},{!buffer!:0,!byteLength!:35280,!byteOffset!:105840},{!buffer!:0,!byteLength!:48384" & _
        ",!byteOffset!:141120}],!buffers!:[{!byteLength!:189504,!uri!:!pano.bin!}]}"
        
gltf1 = Replace(gltf1, "!", Chr$(34))
gltf2 = Replace(gltf2, "!", Chr$(34))
Label2.Caption = "1. Drag/Drop your .jpg or .png Image file (And Audio if desired)" & vbNewLine & _
                 "2. Click on Image to set your desired viewpoint (Green Line)" & vbNewLine & _
                 "3. Choose WinterLodge/ClassicHome/SpaceStation" & vbNewLine & _
                 "3. Setup Audio if desired (default is silent-audio)" & vbNewLine & _
                 "5. Choose Auto Install if you want (Connect Quest)" & vbNewLine & _
                 "4. Click 'Build Panorama'"
                 
Me.Top = (Form1.Top + 600)
Me.Left = (Form1.Left + 1270)
Command4.HoverBackColor = Form1.Command4.HoverBackColor
Command4.HoverForeColor = Form1.Command4.HoverForeColor
lvButtons_H1.CheckDownColor = Form1.Command4.HoverBackColor: lvButtons_H1.HoverBackColor = Form1.Command4.HoverBackColor: lvButtons_H1.Enabled = False: lvButtons_H1.Enabled = True
lvButtons_H2.CheckDownColor = Form1.Command4.HoverBackColor: lvButtons_H2.HoverBackColor = Form1.Command4.HoverBackColor: lvButtons_H2.Enabled = False: lvButtons_H2.Enabled = True
lvButtons_H3.CheckDownColor = Form1.Command4.HoverBackColor: lvButtons_H3.HoverBackColor = Form1.Command4.HoverBackColor: lvButtons_H3.Enabled = False: lvButtons_H3.Enabled = True
lvButtons_H4.CheckDownColor = Form1.Command4.HoverBackColor: lvButtons_H4.HoverBackColor = Form1.Command4.HoverBackColor: lvButtons_H4.Enabled = False: lvButtons_H4.Enabled = True
Check1.HoverBackColor = Form1.Command4.HoverBackColor: Check2.HoverBackColor = Form1.Command4.HoverBackColor: Check3.HoverBackColor = Form1.Command4.HoverBackColor: Check4.HoverBackColor = Form1.Command4.HoverBackColor
Check1.CheckDownColor = Form1.Command4.HoverBackColor: Check2.CheckDownColor = Form1.Command4.HoverBackColor: Check3.CheckDownColor = Form1.Command4.HoverBackColor: Check4.CheckDownColor = Form1.Command4.HoverBackColor:
Check1.Enabled = False: Check1.Enabled = True: Check2.Enabled = False: Check2.Enabled = True: Check3.Enabled = False: Check3.Enabled = True: Check4.Enabled = False: Check4.Enabled = True:
lvButtons_H1.Value = Form1.Check6.Value: lvButtons_H2.Value = Form1.Check7.Value: lvButtons_H3.Value = Form1.Check0.Value: lvButtons_H4.Value = Form1.Check8.Value
Check1.Value = Form1.Check1.Value: Check2.Value = Form1.Check2.Value: Check3.Value = Form1.Check3.Value: Check4.Value = Form1.Check4.Value
Check1.Enabled = Form1.Check1.Enabled: Check2.Enabled = Form1.Check2.Enabled: Check3.Enabled = Form1.Check3.Enabled: Check4.Enabled = Form1.Check4.Enabled
Label1.Enabled = Form1.Label1.Enabled: Label4.Enabled = Form1.Label4.Enabled: Label5.Enabled = Form1.Label5.Enabled: Label10.Enabled = Form1.Label10.Enabled:
lvButtons_H5.CheckDownColor = Form1.Command4.HoverBackColor: lvButtons_H5.HoverBackColor = Form1.Command4.HoverBackColor
lvButtons_H.CheckDownColor = Form1.Command4.HoverBackColor: lvButtons_H.HoverBackColor = Form1.Command4.HoverBackColor
lvButtons_H.Enabled = Form1.lvButtons_H.Enabled

End Sub

Private Sub Check1_Click()

If Check1.Value = True Then
   'Combo5.Enabled = True ': Form1.Combo1.Enabled = True
   lvButtons_H.Enabled = True: Form1.lvButtons_H.Enabled = True
   If Check4.Enabled = True Then Check4.Value = True: Form1.Check4.Value = True
Else
   'Combo5.Enabled = False ': Form1.Combo1.Enabled = False
   Form1.Check4.Enabled = False
   lvButtons_H.Enabled = False: Form1.lvButtons_H.Enabled = False
End If
Form1.Check1.Value = Check1.Value

End Sub

Private Sub lvButtons_H_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error Resume Next

DropList1.DropDown

End Sub

Private Sub DropList1_Closed()

On Error Resume Next

If Combo5.Text = "" Then Exit Sub
lvButtons_H.Caption = Combo5.Text: Form1.lvButtons_H.Caption = Combo5.Text

End Sub

Private Sub lvButtons_H5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Me.hide

End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Command4.Enabled = False Then Exit Sub
Picture5.Cls
If x > 557 Then x = 557
If x < 1 Then x = 1
sva = x
Picture5.DrawWidth = 3
Picture5.Line (sva, 0)-(sva, 500), vbGreen
vp = Int(x / 23.25) + 1

End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Command4.Enabled = False Then Exit Sub
If Button <> 1 Then Exit Sub
If x > 557 Then x = 557
If x < 1 Then x = 1
Picture5.Cls
sva = x
Picture5.DrawWidth = 3
Picture5.Line (sva, 0)-(sva, 500), vbGreen
vp = Int(x / 23.25) + 1

End Sub

Private Sub Picture5_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

On Error Resume Next

Dim intFile As Integer
Dim fields() As String
Dim H1 As Long
Dim W1 As Long
Dim tw As Boolean
Dim k As String

tw = False
With Data
     For intFile = 1 To .Files.Count
         pat = Data.Files.Item(intFile)
     Next intFile
End With
u = ExtractFile(pat)
k = LCase(Right$(u, 3))
If k = "png" Or k = "jpg" Or k = "peg" Then
   ReadImageInfo (pat)
   If m_ImageType < 2 Or m_ImageType > 3 Then
      MessageBeep (16)
      Message "File type '" & k & "' not supported, sorry!"
      Exit Sub
   End If
   If m_Width > 8193 Then
      MessageBeep (16)
      Message "Warning! Image dimension (" & m_Width & "x" & m_Height & ") might be too large for the Quest!", True
   End If
   If (m_Width / m_Height) <> 2 Then
      MessageBeep (16)
      Message "Image aspect ratio wrong, will be distorted in Quest! Aspect Ratio have to be 2:1", True
   End If
   LoadPicture2 pat, Picture5
   use_pic = pat
   Picture5.DrawWidth = 2
   Picture5.Line (4030, 0)-(4030, 5000), vbGreen
Else
   If k = "mp3" Then tw = True
   If k = "aif" Then tw = True
   If k = "iff" Then tw = True
   If k = "ogg" Then tw = True
   If k = "lac" Then tw = True
   If k = "fla" Then tw = True
   If k = "wav" Then tw = True
   If k = "ave" Then tw = True
   If k = "pcm" Then tw = True
   If tw = False Then
      MessageBeep (16)
      Message "File or Audio type '" & k & "' not supported, sorry!", True
      Exit Sub
   End If
End If
If tw = True Then
   Beep
   aud = pat
   pataud = u
   If k = "ogg" Then
      Check4.Enabled = True
      Label10.Enabled = True
   End If
   Label9.Caption = u: Form1.Label9.Caption = u
   Form1.txtOutputs.Text = Form1.txtOutputs.Text & vbNewLine & "Added Audio-file: " & aud & vbNewLine & vbNewLine
   Form1.txtOutputs.SelStart = Len(Form1.txtOutputs.Text)
   Check1.Enabled = True: Form1.Check1.Enabled = True
   Label1.Enabled = True: Form1.Label1.Enabled = True
   Check2.Enabled = False: Form1.Check2.Enabled = False
   Label4.Enabled = False: Form1.Label4.Enabled = False
   Check3.Enabled = False: Form1.Check3.Enabled = False
   Label5.Enabled = False: Form1.Label5.Enabled = False
   Exit Sub
End If
Form1.txtOutputs.Text = Form1.txtOutputs.Text & vbNewLine & "Added Image-File: " & pat & vbNewLine & vbNewLine
Form1.txtOutputs.SelStart = Len(Form1.txtOutputs.Text)
'HScroll1.Enabled = True
Picture5.Cls
sva = Int(23.36 * 12) - 12
vp = 12
Picture5.DrawWidth = 3
Picture5.Line (sva, 0)-(sva, 500), vbGreen
Command4.Enabled = True

End Sub

Private Sub Command4_Click()

On Error Resume Next

Form1.txtOutputs.Text = Form1.txtOutputs.Text & vbNewLine & "Building Panorama Files" & vbNewLine & vbNewLine
Form1.txtOutputs.SelStart = Len(Form1.txtOutputs.Text)
FileCopy App.path & "\files\pano.bin", BuildPath & "\pano.bin"
FileCopy use_pic, BuildPath & "\pano.jpg"
If vp < 1 Then vp = 1
If vp > 24 Then vp = 24
If Dir(BuildPath & "\pano.gltf") <> "" Then Kill BuildPath & "\pano.gltf"
Open BuildPath & "\pano.gltf" For Output As #1
Print #1, gltf1 & rota(vp) & gltf2
Close #1
start_pano = True
Me.hide

End Sub

Private Sub Check2_Click()

On Error Resume Next

If Check2.Value = True Then Check3.Value = False: Form1.Check3.Value = False
If Form1.Label8.Caption <> "" Then Form1.Command1.Enabled = True
Form1.Check2.Value = Check2.Value

End Sub

Private Sub Check3_Click()

On Error Resume Next

If Check3.Value = True Then Check2.Value = False: Form1.Check2.Value = False
If Form1.Label8.Caption <> "" Then Form1.Command1.Enabled = True
Form1.Check3.Value = Check3.Value

End Sub

Private Sub Check4_Click()

On Error Resume Next

If Check4.Value = False Then Check1.Value = False: Form1.Check1.Value = False
Form1.Check4.Value = Check4.Value

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


Private Sub TimerPic_Timer()

'Picture1.Cls
Picture5.CurrentX = 900
Picture5.CurrentY = 1600
Picture5.Font = "Arial"
Picture5.ForeColor = vbWhite
Picture5.FontSize = 23
Picture5.Print "Drag/Drop your 360° Image here"
Picture5.CurrentX = 150
Picture5.Print "Also Drag/Drop Audio file here if desired"
TimerPic.Enabled = False

End Sub

Private Sub lvButtons_H1_Click()

On Error Resume Next

If lvButtons_H1.Value = False Then lvButtons_H1.Value = True: Exit Sub
lvButtons_H2.Value = False: lvButtons_H3.Value = False
Form1.Check6.Value = lvButtons_H1.Value
Form1.Check7.Value = lvButtons_H2.Value
Form1.Check0.Value = lvButtons_H3.Value

End Sub

Private Sub lvButtons_H2_Click()

On Error Resume Next

If lvButtons_H2.Value = False Then lvButtons_H2.Value = True: Exit Sub
lvButtons_H1.Value = False: lvButtons_H3.Value = False
Form1.Check6.Value = lvButtons_H1.Value
Form1.Check7.Value = lvButtons_H2.Value
Form1.Check0.Value = lvButtons_H3.Value

End Sub

Private Sub lvButtons_H3_Click()

On Error Resume Next

If lvButtons_H3.Value = False Then lvButtons_H3.Value = True: Exit Sub
lvButtons_H1.Value = False: lvButtons_H2.Value = False
Form1.Check6.Value = lvButtons_H1.Value
Form1.Check7.Value = lvButtons_H2.Value
Form1.Check0.Value = lvButtons_H3.Value

End Sub

Private Sub lvButtons_H4_Click()

On Error Resume Next

Form1.Check8.Value = lvButtons_H4.Value
'Form1.Check6.Value = lvButtons_H1.Value
'Form1.Check7.Value = lvButtons_H2.Value
'Form1.Check0.Value = lvButtons_H3.Value

End Sub
