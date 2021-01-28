VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0025221F&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Enviroment Converter"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14415
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   14415
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11400
      Top             =   6360
   End
   Begin VB.Timer Timer2 
      Left            =   11880
      Top             =   8760
   End
   Begin VB.PictureBox Boarder1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'Kein
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   17775
      TabIndex        =   45
      Top             =   0
      Width           =   17775
      Begin Projekt1.lvButtons_H lvButtons_H12 
         Height          =   375
         Left            =   9120
         TabIndex        =   92
         Top             =   90
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   661
         CapAlign        =   2
         Shape           =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   1
         Value           =   -1  'True
         Image           =   "Form1.frx":1AC21
         cBack           =   4210752
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'Kein
         Enabled         =   0   'False
         Height          =   300
         Left            =   9560
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   53
         Top             =   110
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'Kein
         Height          =   300
         Left            =   10040
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   52
         Top             =   110
         Width           =   300
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'Kein
         Height          =   300
         Left            =   10520
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   51
         Top             =   110
         Width           =   300
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'Kein
         Height          =   300
         Left            =   11000
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   48
         Top             =   110
         Width           =   300
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   110
         Picture         =   "Form1.frx":1E72B
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   47
         Top             =   60
         Width           =   375
      End
      Begin VB.Label Label21 
         BackColor       =   &H00404040&
         Caption         =   "Environment Converter and Builder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   520
         TabIndex        =   46
         Top             =   110
         Width           =   5295
      End
   End
   Begin Projekt1.DropList DropList1 
      Height          =   315
      Left            =   5520
      TabIndex        =   21
      Top             =   3270
      Width           =   705
      _ExtentX        =   1296
      _ExtentY        =   556
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   240
      OLEDropMode     =   1  'Manuell
      Picture         =   "Form1.frx":1EF45
      ScaleHeight     =   2745
      ScaleWidth      =   2850
      TabIndex        =   19
      Top             =   720
      Width           =   2880
   End
   Begin Projekt1.lvButtons_H Command2 
      Height          =   375
      Left            =   10320
      TabIndex        =   18
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "Settings"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
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
   Begin Projekt1.lvButtons_H Command1 
      Height          =   855
      Left            =   240
      TabIndex        =   14
      ToolTipText     =   "Start Converter"
      Top             =   4560
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1508
      Caption         =   "Start Converter"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H0025221F&
      Caption         =   "Enviroment Builder ( Create Environment with Blender and Build )"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   3480
      TabIndex        =   13
      Top             =   4440
      Width           =   7695
      Begin Projekt1.lvButtons_H Command5 
         Height          =   510
         Left            =   5685
         TabIndex        =   17
         ToolTipText     =   "Creates WinterLodge/ Classic Home and SpaceStation with Audio and the same with silent Audio"
         Top             =   1440
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   900
         Caption         =   "Create 8 releases with and without Audio"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         cBack           =   4210752
      End
      Begin Projekt1.lvButtons_H Command4 
         Height          =   510
         Left            =   5685
         TabIndex        =   16
         Top             =   240
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   900
         Caption         =   "Open Build folder"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         cBack           =   4210752
      End
      Begin Projekt1.lvButtons_H Command3 
         Height          =   972
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Put your model-files in folder .\Build"
         Top             =   360
         Width           =   1452
         _ExtentX        =   2566
         _ExtentY        =   2143
         Caption         =   "Build and Install Enviroment"
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
         cBack           =   4210752
      End
      Begin Projekt1.lvButtons_H Check6 
         Height          =   255
         Left            =   1680
         TabIndex        =   26
         Top             =   480
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
      Begin Projekt1.lvButtons_H Check7 
         Height          =   255
         Left            =   1680
         TabIndex        =   28
         Top             =   840
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
      Begin Projekt1.lvButtons_H Check8 
         Height          =   255
         Left            =   3360
         TabIndex        =   30
         ToolTipText     =   "Install APK after build"
         Top             =   480
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
      Begin Projekt1.lvButtons_H Check9 
         Height          =   255
         Left            =   3360
         TabIndex        =   32
         ToolTipText     =   "Automatically detects when Blender exports new files to .\Build"
         Top             =   840
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
      Begin Projekt1.lvButtons_H Check0 
         Height          =   255
         Left            =   1680
         TabIndex        =   37
         Top             =   1200
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
      Begin Projekt1.lvButtons_H Check10 
         Height          =   255
         Left            =   3360
         TabIndex        =   49
         ToolTipText     =   "Compress all Textures in ../Build to KTX when pressing Build"
         Top             =   1200
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
         Enabled         =   0   'False
         cBack           =   8421504
      End
      Begin Projekt1.lvButtons_H Check13 
         Height          =   255
         Left            =   3360
         TabIndex        =   64
         ToolTipText     =   "Current textures are saved and copied back with every build"
         Top             =   1560
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
         Enabled         =   0   'False
         cBack           =   8421504
      End
      Begin Projekt1.lvButtons_H Command12 
         Height          =   510
         Left            =   5685
         TabIndex        =   68
         Top             =   840
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   900
         Caption         =   "    Delete all Files in      Build Folder"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         cBack           =   4210752
      End
      Begin Projekt1.lvButtons_H Check19 
         Height          =   255
         Left            =   1680
         TabIndex        =   84
         Top             =   1560
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
      Begin Projekt1.lvButtons_H lvButtons_H8 
         Height          =   372
         Left            =   120
         TabIndex        =   86
         ToolTipText     =   "Put your model-files in folder .\Build"
         Top             =   1440
         Width           =   1452
         _ExtentX        =   2566
         _ExtentY        =   1508
         Caption         =   "Reset Audio"
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
         cBack           =   4210752
      End
      Begin VB.Label Label31 
         BackColor       =   &H0025221F&
         Caption         =   "CyberCity"
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
         Left            =   1995
         TabIndex        =   85
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label25 
         BackColor       =   &H0025221F&
         Caption         =   "Texture Protection"
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
         Left            =   3720
         TabIndex        =   65
         ToolTipText     =   "Current textures are saved and copied back with every build"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label22 
         BackColor       =   &H0025221F&
         Caption         =   "Compress to KTX"
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
         Left            =   3720
         TabIndex        =   50
         ToolTipText     =   "Compress all Textures in ../Build to KTX when pressing Build"
         Top             =   1200
         Width           =   1935
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
         Left            =   1995
         TabIndex        =   38
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackColor       =   &H0025221F&
         Caption         =   "Auto Build and Install"
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
         Left            =   3720
         TabIndex        =   33
         ToolTipText     =   "Automatically detects when Blender exports new files to .\Build"
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label15 
         BackColor       =   &H0025221F&
         Caption         =   "Auto Install after build"
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
         Left            =   3720
         TabIndex        =   31
         ToolTipText     =   "Install APK after build"
         Top             =   480
         Width           =   1935
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
         Left            =   1995
         TabIndex        =   29
         Top             =   840
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
         Left            =   1995
         TabIndex        =   27
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   12480
      Top             =   8760
   End
   Begin VB.TextBox txtOutputs 
      Appearance      =   0  '2D
      BackColor       =   &H80000008&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H00FF80FF&
      Height          =   2535
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   10
      Top             =   6645
      Width           =   10935
   End
   Begin Projekt1.lvButtons_H lvButtons_H 
      Height          =   315
      Left            =   5520
      TabIndex        =   20
      Top             =   2970
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
      cBack           =   4210752
   End
   Begin Projekt1.lvButtons_H Check2 
      Height          =   255
      Left            =   3480
      TabIndex        =   22
      Top             =   3360
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
      Left            =   3480
      TabIndex        =   23
      Top             =   3720
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
   Begin Projekt1.lvButtons_H Check1 
      Height          =   255
      Left            =   3480
      TabIndex        =   24
      Top             =   3000
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
   Begin Projekt1.lvButtons_H Check4 
      Height          =   255
      Left            =   3480
      TabIndex        =   25
      Top             =   4080
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
   Begin Projekt1.lvButtons_H Check5 
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   3675
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
   Begin Projekt1.lvButtons_H Option1 
      Height          =   255
      Left            =   3480
      TabIndex        =   35
      Top             =   720
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
   Begin Projekt1.lvButtons_H Option2 
      Height          =   255
      Left            =   3480
      TabIndex        =   36
      Top             =   1320
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
   Begin Projekt1.lvButtons_H lvButtons_H1 
      Height          =   255
      Left            =   3840
      TabIndex        =   39
      Top             =   1020
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
      Left            =   5520
      TabIndex        =   41
      Top             =   1020
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
      Left            =   7320
      TabIndex        =   43
      Top             =   1020
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
   Begin Projekt1.lvButtons_H Command6 
      Height          =   520
      Left            =   11880
      TabIndex        =   54
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   926
      Caption         =   "Change Hover Color"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
   Begin Projekt1.lvButtons_H Command7 
      Height          =   525
      Left            =   11880
      TabIndex        =   55
      Top             =   1320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   926
      Caption         =   "Change Title Bar Color"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
   Begin Projekt1.lvButtons_H Command8 
      Height          =   525
      Left            =   11880
      TabIndex        =   56
      Top             =   3120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   926
      Caption         =   "Change Build Folder Location"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
   Begin Projekt1.lvButtons_H Command9 
      Height          =   375
      Left            =   13080
      TabIndex        =   57
      Top             =   8760
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "About"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
   Begin Projekt1.lvButtons_H Command10 
      Height          =   525
      Left            =   11880
      TabIndex        =   58
      Top             =   1920
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   926
      Caption         =   "Change Console output Color"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
   Begin Projekt1.lvButtons_H Command11 
      Height          =   525
      Left            =   11880
      TabIndex        =   59
      Top             =   2520
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   926
      Caption         =   "Set default Colors"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
      Gradient        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   4210752
   End
   Begin Projekt1.lvButtons_H Check11 
      Height          =   255
      Left            =   11880
      TabIndex        =   60
      Top             =   3960
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
   Begin Projekt1.lvButtons_H Check12 
      Height          =   255
      Left            =   11880
      TabIndex        =   63
      Top             =   4560
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
   Begin Projekt1.lvButtons_H Check14 
      Height          =   255
      Left            =   11880
      TabIndex        =   66
      Top             =   5160
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
   Begin Projekt1.lvButtons_H Check15 
      Height          =   255
      Left            =   11880
      TabIndex        =   69
      Top             =   5760
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
   Begin Projekt1.lvButtons_H Check16 
      Height          =   255
      Left            =   11880
      TabIndex        =   71
      Top             =   6360
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
   Begin Projekt1.lvButtons_H Command13 
      Height          =   525
      Left            =   11880
      TabIndex        =   73
      Top             =   10560
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   926
      Caption         =   "Change ADB Wireless Port"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
   Begin Projekt1.lvButtons_H Command14 
      Height          =   375
      Left            =   9360
      TabIndex        =   74
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "Help"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
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
   Begin Projekt1.lvButtons_H Check17 
      Height          =   255
      Left            =   11880
      TabIndex        =   77
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
      Value           =   -1  'True
      cBack           =   8421504
   End
   Begin Projekt1.lvButtons_H Check18 
      Height          =   255
      Left            =   11880
      TabIndex        =   78
      Top             =   9480
      Visible         =   0   'False
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
   Begin Projekt1.lvButtons_H Command15 
      Height          =   855
      Left            =   240
      TabIndex        =   79
      ToolTipText     =   "Start Converter"
      Top             =   5640
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1508
      Caption         =   "Panorama Builder"
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
   Begin Projekt1.lvButtons_H lvButtons_H4 
      Height          =   375
      Left            =   4560
      TabIndex        =   80
      ToolTipText     =   "Unpack APK to Build folder"
      Top             =   1680
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Caption         =   "Unpack APK to ..\Build"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
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
   Begin Projekt1.lvButtons_H lvButtons_H5 
      Height          =   375
      Left            =   8160
      TabIndex        =   81
      Tag             =   "Install APK"
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Only install APK"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
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
   Begin Projekt1.lvButtons_H lvButtons_H6 
      Height          =   375
      Left            =   9720
      TabIndex        =   82
      Tag             =   "Uninstall APK"
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Uninstall APK"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
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
   Begin Projekt1.lvButtons_H lvButtons_H7 
      Height          =   390
      Left            =   9360
      TabIndex        =   83
      Top             =   1200
      Width           =   1810
      _ExtentX        =   3201
      _ExtentY        =   688
      Caption         =   "Save last Quest Video"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
      cBack           =   4210752
   End
   Begin Projekt1.lvButtons_H lvButtons_H9 
      Height          =   255
      Left            =   11880
      TabIndex        =   87
      Top             =   9960
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
   Begin Projekt1.lvButtons_H lvButtons_H10 
      Height          =   375
      Left            =   6720
      TabIndex        =   89
      ToolTipText     =   "Convert all ktx textures to png in Build folder"
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "PNG Convert"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
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
   Begin Projekt1.lvButtons_H lvButtons_H11 
      Height          =   255
      Left            =   11880
      TabIndex        =   90
      Top             =   7560
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
   Begin VB.Label Label33 
      BackColor       =   &H0025221F&
      Caption         =   "Auto Reset Audio after Build Install"
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
      Height          =   495
      Left            =   12240
      TabIndex        =   91
      ToolTipText     =   "Stres the state of all check boxes/Buttons"
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Label Label32 
      BackColor       =   &H0025221F&
      Caption         =   "Mipmapping KTX"
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
      Left            =   12240
      TabIndex        =   88
      ToolTipText     =   "Stres the state of all check boxes/Buttons"
      Top             =   9960
      Width           =   3135
   End
   Begin VB.Label Label30 
      BackColor       =   &H0025221F&
      Caption         =   "WiFi Auto Connect"
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
      Left            =   12240
      TabIndex        =   76
      ToolTipText     =   "Automatically connects to the Quest via WiFi when the converter is started"
      Top             =   9480
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label29 
      BackColor       =   &H0025221F&
      Caption         =   "Pack all 6 releases to zip file"
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
      Height          =   495
      Left            =   12240
      TabIndex        =   75
      ToolTipText     =   "Untick if you don´t want to zip the releases"
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label Label28 
      BackColor       =   &H0025221F&
      Caption         =   "Kill ADB Server on Exit"
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
      Height          =   495
      Left            =   12240
      TabIndex        =   72
      ToolTipText     =   "Useful if you want to save time and keep the adb connection"
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label27 
      BackColor       =   &H0025221F&
      Caption         =   "Auto delete Files in ..\Build after Build"
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
      Height          =   495
      Left            =   12240
      TabIndex        =   70
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label26 
      BackColor       =   &H0025221F&
      Caption         =   "Delete protected Textures on exit"
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
      Height          =   495
      Left            =   12240
      TabIndex        =   67
      ToolTipText     =   "Delete folder with protected textures (only for builder)"
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label24 
      BackColor       =   &H0025221F&
      Caption         =   "Store Button state"
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
      Left            =   12240
      TabIndex        =   62
      ToolTipText     =   "Stres the state of all check boxes/Buttons"
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Label Label23 
      BackColor       =   &H0025221F&
      Caption         =   "Store Window Position on exit"
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
      Height          =   495
      Left            =   12240
      TabIndex        =   61
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label20 
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
      Left            =   7680
      TabIndex        =   44
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Label Label19 
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
      Left            =   5880
      TabIndex        =   42
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Label Label18 
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
      Left            =   4200
      TabIndex        =   40
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackColor       =   &H0025221F&
      Caption         =   "Install Environment-APK         after conversion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   660
      TabIndex        =   12
      Top             =   3570
      Width           =   2490
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   3480
      X2              =   11160
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label11 
      BackColor       =   &H0025221F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   780
      Left            =   7560
      TabIndex        =   11
      Top             =   3480
      Width           =   3495
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
      Left            =   3840
      TabIndex        =   9
      ToolTipText     =   "Re-encode the OGG-audio file to reduce the size"
      Top             =   4080
      Width           =   3615
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
      Left            =   4560
      TabIndex        =   8
      Top             =   2640
      Width           =   6615
   End
   Begin VB.Label Label8 
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
      Left            =   4560
      TabIndex        =   7
      Top             =   2160
      Width           =   6615
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
      Left            =   3840
      TabIndex        =   6
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H0025221F&
      Caption         =   "APK   :"
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
      Left            =   3840
      TabIndex        =   5
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H0025221F&
      Caption         =   "Replace audio with silent audio file"
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
      Left            =   3840
      TabIndex        =   4
      ToolTipText     =   "Exchange the audio file with an empty (silent) audio file"
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackColor       =   &H0025221F&
      Caption         =   "Replace audio with default audio file"
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
      Left            =   3840
      TabIndex        =   3
      ToolTipText     =   "Exchange the audio file with the standard fireplace sound"
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H0025221F&
      Caption         =   "Replace only Audio"
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
      Left            =   3840
      TabIndex        =   2
      ToolTipText     =   "Only change the audio file and not the environment"
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H0025221F&
      Caption         =   "Switch WinterLodge/ ClassicHome/ SpaceStation (and audio)"
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
      Left            =   3840
      TabIndex        =   1
      ToolTipText     =   "Change the environment and optionally the audio file"
      Top             =   720
      Width           =   5535
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
      Left            =   3840
      TabIndex        =   0
      ToolTipText     =   "Most audio files are too loud for an environment"
      Top             =   3000
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Fertig:
'*Neuer Rahmen oben Win10
'*Farben allgemein ändern mit einstellungen
'*Design neu
'*Java Prüfung
'*Farbe in Elin Farbe ändern
'*Einstellungen farbe bei start übernehmen
'*Fehler wird angezeig in command line
'*save last window pos on exit
'*Option für auto(save) und button clear build path when builded = If Dir(BuildPath & "\*.*") <> "" Then Kill BuildPath & "\*.*"
'*Neues SpaceStation Env. aufnehmen unf ggf. Automatik bauen und auswahl bei switch aus zweien
'*Beim löschen der gltf dauerschleife install entfernen!!! (In IDE macht er dauerschleife!!! und prüft java nicht)
'*Startet auch wenn Build leer und text datei erzeugt wird!!!??????????
'*Audio switch name _new ändern
'*Prüfung glTF doppelt vorhanden
'*Delete old textures before copy bei textur_tmp
'*Build und " & apppath & "\files folder prüfung, vieleicht mit dateien
'*Prüfung audio file arten oder apk, sonst fehler oder nichts
'*Textur schutz bei modifizierten texturen (schreibschutz oder copy beim builden)
'*Prüfung ob Install erfolgreich war (Kritischer Fehler sound und "Done!" mit Rot "Error!" ersetzen
'*Neuer audio codec
'*Remote Wireless install build
'*save Quest IP
'*port in config and change Button
'*Button Kill ADB Server on Exit
'*change Port Button
'*Label deaktivieren von neuen funktionen (blink kurz auf bei start)
'*Onerror resume überall
'*Files Ordner aufräumen
'*Konsole ausgabe nachtragen bei neuen sachen
'*Msgbox durch neue Message funktion tauschen
'*Neuen audio codec Testen ob speed/Pitch auf bei -8db ok ist

'v1.7.1 fertig:
'*logo kaputt unter win10
'Msgbox in timer löschen und Wifi connect grün fertigstellen
'WiFi connect freeze problem lösen
'tooltips reparieren
'tcp einbauen und standart port auf 5555
'Connect als fehlermeldung (cannot connect to) aufnehmen
'Wifi Auto connect bei start option
'ADB WiFi active when adb is running at start
'New freeze free command window for adb connect
'default audio file wird silent bei release
'Create 6 releases without zip them (option)
'InputBox("Pure APK-Filename tauschen durch neue Box
'Create 6 releases ende ohne deaktivierung des tools
'delete gltf before open create
'Rahmen Form7 anpassen und farben Checkbox anpassen
'Rahmenposition anpassen
'Pano Builder mit möglicher Rotation
'Audio Spiegeln Panobuilder
'Neuer Font (Arial)
'Help Button (Form5)
'Uninstaller for environments
'apk unpack to Build folder
'restart after release creation
'Install funktion für env.apks
'texture_temp in " & apppath & "\files
'kommastellen im db decrease ist nicht funktional mit sox compiler (audio) also nich -16.5 db
'Pitch/Speed increase when encoding audio without decrase volume
'No more restart after Convert/Switch or create

'Todo:


'Next Release:

'JPG compression und png conversion mit modul in autinstall
'default audio je nach typ wählen (SpaceStation hat kein Fireplace sound)
'ADB vorher beenden/task kill wenn WiFi connect
'Delete untitled.apk beim build after install option
'Freeze wenn WiFi ADB connect lost (MSGBOX noch da?)
'ogg file encode bug fixed
'Delete dragged audio (text) Button
'prüfung ob bei build jpg und ktx gleichzeitig
'texture protect
'cyberhome
'loop install wenn schon ktx im verzeichniss bei autoinstall

'Fertig:

'Save last Quest Video

Private Type BrowseInfo
    lngHwnd        As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "Kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private WithEvents objDOS As DOSOutputs
Attribute objDOS.VB_VarHelpID = -1

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustomFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (lpofn As OPENFILENAME) As Long

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwflags As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
                         ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
                         
Private Declare Function GetRTTAndHopCount Lib "iphlpapi.dll" (ByVal lDestIPAddr As Long, ByRef lHopCount As Long, _
                         ByVal lMaxHops As Long, ByRef lRTT As Long) As Long
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long

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
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (lpChoosecolor As ChooseColorStruct) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
  ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private WithEvents Combo1 As ComboBox
Attribute Combo1.VB_VarHelpID = -1
                         
Private snd As Long
Private ogg As String
Private pack As String
Private patapk As String
Private J As String
Private java As String
Private za As Long
Private qw As Integer
Private sa As String
Private adb As Integer
Private fn As String
Private fsx As Object
Private oFile As Object
Private t1 As String
Private fin2 As String
Private s As Long
Private storedx As Integer
Private storedy As Integer
Private dra As Integer
Private tp As Boolean
Private qip As String
Private wcon As Boolean
Private za2 As Long
Private com3_Self As Boolean
Private out_text As Boolean
Private start_adb As Boolean
Private wifi_auto As Boolean
Private apppath As String
Private comp1 As Boolean
Private fowi As Long
Private sour As String
Private dest As String

Private Sub Form_Load()

On Error Resume Next

Dim i As Long
Dim ctrl As Control
Dim cl As String
Dim ch As String
Dim hbc As String
Dim t2 As String
Dim bt As Boolean

Form1.Width = 11520
'Picture1.ScaleMode = 3
'Picture1.AutoRedraw = True
'Picture1.PaintPicture Picture1.Picture, _
'0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, _
'0, 0, Picture1.Picture.Width / 21.2, _
'Picture1.Picture.Height / 21.2
'Picture1.Picture = Picture1.Image
'Form7.lvButtons_H.Enabled = False
wcon = False
za2 = 1
start_adb = True
start_pano = False
apppath = App.path
Call SetWindowPos(Form1.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
Set fsx = CreateObject("Scripting.FileSystemObject")
If fsx.FolderExists(apppath & "\files") = False Then
   Timer1.Enabled = False
   Timer2.Enabled = False
   MessageBeep (16)
   Message "No .\files Folder found, Error!"
   End
End If
If Dir(apppath & "\files\SpaceStation.zip") = "" Or Dir(apppath & "\files\WinterLodge.zip") = "" Or _
   Dir(apppath & "\files\ClassicHome.zip") = "" Or Dir(apppath & "\files\CyberCity.zip") = "" Or Dir(apppath & "\files\adb.exe") = "" Or _
   fsx.FolderExists(apppath & "\files\ClassicHome") = False Or fsx.FolderExists(apppath & "\files\SpaceStation") = False Or _
   fsx.FolderExists(apppath & "\files\WinterLodge") = False Or fsx.FolderExists(apppath & "\files\CyberCity") = False Then
   Timer1.Enabled = False
   Timer2.Enabled = False
   MessageBeep (16)
   Message "Missing files in ..\files Folder, Error!"
   End
End If
If Dir(apppath & "\files\config.ini") = "" Then
   PutINISetting "Color", "HoverBackColor", "&H00BF1675&", apppath & "\files\config.ini"
   PutINISetting "Color", "TitleBarColor", "&H00404040&", apppath & "\files\config.ini"
   PutINISetting "Color", "ConsoleColor", "&H00FF80FF&", apppath & "\files\config.ini"
   PutINISetting "Paths", "BuildPath", "", apppath & "\files\config.ini"
   PutINISetting "WindowPos", "Left", Form1.Left, apppath & "\files\config.ini"
   PutINISetting "WindowPos", "Top", Form1.Top, apppath & "\files\config.ini"
   PutINISetting "CheckValue", "Check", "0101100", apppath & "\files\config.ini"
   PutINISetting "Save", "WindowPos", "1", apppath & "\files\config.ini"
   PutINISetting "Save", "ButtonState", "1", apppath & "\files\config.ini"
   PutINISetting "Save", "TextureDelete", "1", apppath & "\files\config.ini"
   PutINISetting "Save", "AutoClear", "0", apppath & "\files\config.ini"
   PutINISetting "QuestIP", "Port", "5555", apppath & "\files\config.ini"
   PutINISetting "Save", "ADBKill", "1", apppath & "\files\config.ini"
   PutINISetting "Save", "Pack", "1", apppath & "\files\config.ini"
   PutINISetting "Save", "WiFiAuto", "0", apppath & "\files\config.ini"
   PutINISetting "Save", "AudioReset", "0", apppath & "\files\config.ini"
End If
If GetINISetting("Save", "AudioReset", apppath & "\files\config.ini") = "" Then PutINISetting "Save", "AudioReset", "0", apppath & "\files\config.ini"
If GetINISetting("Save", "WiFiAuto", apppath & "\files\config.ini") = "" Then PutINISetting "Save", "WiFiAuto", "0", apppath & "\files\config.ini"
If GetINISetting("Save", "ADBKill", apppath & "\files\config.ini") = "" Then PutINISetting "Save", "ADBKill", "1", apppath & "\files\config.ini"
If GetINISetting("QuestIP", "Port", apppath & "\files\config.ini") = "" Then PutINISetting "QuestIP", "Port", "5555", apppath & "\files\config.ini"
If GetINISetting("Save", "Pack", apppath & "\files\config.ini") = "" Then PutINISetting "Save", "Pack", "1", apppath & "\files\config.ini"
If GetINISetting("Save", "WiFiAuto", apppath & "\files\config.ini") = "1" Then
   Check18.Value = True
   wifi_auto = True
Else
   Check18.Value = False
End If
If GetINISetting("Save", "Pack", apppath & "\files\config.ini") = "1" Then
   Check17.Value = True
Else
   Check17.Value = False
End If
If GetINISetting("Save", "AutoClear", apppath & "\files\config.ini") = "1" Then
   Check15.Value = True
Else
   Check15.Value = False
End If
If GetINISetting("Save", "AudioReset", apppath & "\files\config.ini") = "1" Then
   lvButtons_H11.Value = True
Else
   lvButtons_H11.Value = False
End If
If GetINISetting("Save", "WindowPos", apppath & "\files\config.ini") = "1" Then
   Form1.Left = GetINISetting("WindowPos", "Left", apppath & "\files\config.ini")
   Form1.Top = GetINISetting("WindowPos", "Top", apppath & "\files\config.ini")
Else
   Check11.Value = False
End If
If GetINISetting("Save", "ButtonState", apppath & "\files\config.ini") = "1" Then
   ch = GetINISetting("CheckValue", "Check", apppath & "\files\config.ini")
   Check0.Value = Mid(ch, 1, 1): Check6.Value = Mid(ch, 2, 1): Check7.Value = Mid(ch, 3, 1)
   Check8.Value = Mid(ch, 4, 1): Check9.Value = Mid(ch, 5, 1): Check10.Value = Mid(ch, 6, 1): Check19.Value = Mid(ch, 7, 1)
Else
   Check12.Value = False
End If
cl = GetINISetting("Color", "TitleBarColor", apppath & "\files\config.ini")
If GetINISetting("Paths", "BuildPath", apppath & "\files\config.ini") <> "" Then
   BuildPath = GetINISetting("Paths", "BuildPath", apppath & "\files\config.ini")
Else
   BuildPath = apppath & "\Build"
End If
Boarder1.BackColor = HTC(cl): Label21.BackColor = HTC(cl): Picture2.BackColor = HTC(cl)
Picture3.BackColor = HTC(cl): Picture4.BackColor = HTC(cl): Picture5.BackColor = HTC(cl): Picture6.BackColor = HTC(cl)
If Lux(HTC(GetINISetting("Color", "HoverBackColor", apppath & "\files\config.ini"))) > 120 Then
   Command1.HoverForeColor = vbBlack: Command2.HoverForeColor = vbBlack: Command3.HoverForeColor = vbBlack
   Command4.HoverForeColor = vbBlack: Command5.HoverForeColor = vbBlack: Command6.HoverForeColor = vbBlack
   Command7.HoverForeColor = vbBlack: Command8.HoverForeColor = vbBlack: Command9.HoverForeColor = vbBlack
   Command10.HoverForeColor = vbBlack: Command11.HoverForeColor = vbBlack
End If
hbc = HTC(GetINISetting("Color", "ConsoleColor", apppath & "\files\config.ini"))
txtOutputs.ForeColor = hbc
hbc = HTC(GetINISetting("Color", "HoverBackColor", apppath & "\files\config.ini"))
For Each ctrl In Form1
    If ctrl.HoverBackColor <> "" Then
       ctrl.HoverBackColor = hbc
       ctrl.CheckDownColor = hbc
       bt = ctrl.Enabled
       ctrl.Enabled = False
       ctrl.Enabled = True
       ctrl.Enabled = bt
       If Lux(hbc) > 120 Then
          ctrl.HoverForeColor = vbBlack
       Else
          ctrl.HoverForeColor = vbWhite
       End If
    End If
Next
For Each ctrl In Form7
    If ctrl.HoverBackColor <> "" Then
       ctrl.HoverBackColor = hbc
       ctrl.CheckDownColor = hbc
       bt = ctrl.Enabled
       ctrl.Enabled = False
       ctrl.Enabled = True
       ctrl.Enabled = bt
       If Lux(hbc) > 120 Then
          ctrl.HoverForeColor = vbBlack
       Else
          ctrl.HoverForeColor = vbWhite
       End If
    End If
Next
lvButtons_H12.CheckDownColor = HTC(GetINISetting("Color", "TitleBarColor", apppath & "\files\config.ini"))

If Command1.BackColor = HTC(GetINISetting("Color", "TitleBarColor", apppath & "\files\config.ini")) Then
   lvButtons_H12.BackColor = RGB(110, 110, 110)
Else
   lvButtons_H12.BackColor = Command1.BackColor
End If
lvButtons_H12.Enabled = False
lvButtons_H12.Enabled = True
Form7.Check4.Enabled = False
Form7.Check1.Enabled = False
Form7.lvButtons_H.Enabled = False
'Picture2.Picture = Form1.Icon
Timer2.Enabled = True
Timer2.Interval = 200
Me.BorderStyle = 0
Me.Caption = Me.Caption
s = SetSysColors(1, COLOR_CAPTIONTEXT, vbBlack)
For i = 1 To 32
Set Combo1 = DropList1.Combo
lvButtons_H.Enabled = False
Combo1.AddItem (i)
Next i
Combo1.Text = "19"
Set objDOS = New DOSOutputs
za = 0
qw = 1
adb = 0
If Dir$(BuildPath & "\*.*") = vbNullString Then
   Check6.Enabled = False
   Check0.Enabled = False
   Check7.Enabled = False
   Check8.Enabled = False
   Check9.Enabled = False
   Check10.Enabled = False
   Check13.Enabled = False
   Command3.Enabled = False
   Command5.Enabled = False
   Label13.Enabled = False
   Label14.Enabled = False
   Label15.Enabled = False
   Label16.Enabled = False
   Label17.Enabled = False
   Label22.Enabled = False
   Label25.Enabled = False
   Label31.Enabled = False
   Check19.Enabled = False
Else
   Check10.Enabled = True
   Check13.Enabled = True
   qw = 0
   txtOutputs.Text = txtOutputs.Text & vbNewLine & "Found gltf-model in .\Build" & vbNewLine & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
End If
lvButtons_H1.Visible = False
lvButtons_H2.Visible = False
lvButtons_H3.Visible = False
Label18.Visible = False
Label19.Visible = False
Label20.Visible = False
'Combo1.Enabled = False
Check1.Enabled = False
Label1.Enabled = False
Option2.Value = True
Check4.Enabled = False
Label10.Enabled = False
Command1.Enabled = False
fowi = Form1.Width
If Dir(apppath & "\files\Java64\bin\java.exe") <> "" Then
   java = Chr$(34) & apppath & "\files\Java64\bin\java.exe" & Chr$(34)
Else
   java = "java"
   If App.LogMode = 1 And Shell("java", vbHide) = 0 And Dir(apppath & "\files\Java64\bin\java.exe") = "" Then
      Timer1.Enabled = False
      Timer2.Enabled = False
      MessageBeep (16)
      Message " Error, Java not found! " & vbNewLine & " Please Install Java or use Portable Java Converter Version"
      End
   End If
End If
If GetINISetting("Save", "Splash", apppath & "\files\config.ini") = "" Then
   PutINISetting "Save", "Splash", "1", apppath & "\files\config.ini"
   Call Form9.Show(vbModal)
End If

If Command = "" Then Exit Sub

Open apppath & "\batch.txt" For Input As #1
  Line Input #1, dest
  Line Input #1, dest
Close #1
If Right$(dest, 1) <> "\" Then dest = dest & "\"
Timer3.Enabled = True


End Sub

Private Sub lvButtons_H10_Click()

On Error Resume Next

Dim fso As New FileSystemObject
Dim fil As File
Dim fpat As String
Dim tkt As String
Dim f As Integer

For Each fil In fso.GetFolder(BuildPath).Files
  tkt = fso.GetExtensionName(BuildPath & "\" & fil.Name)
  If tkt = "ktx" Then
     txtOutputs.Text = txtOutputs.Text & vbNewLine & "Convert KTX-Texture File to JPG: " & BuildPath & "\" & fil.Name & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
     'PVRTexToolCLI -i 1.ktx -f r8g8b8a8 -d 2.png
     objDOS.CommandLine = (apppath & "\files\PVRTexToolCLI.exe -i " & J & BuildPath & "\" & fil.Name & J & " -f r8g8b8a8 -d " & J & BuildPath & "\" & fso.GetBaseName(BuildPath & "\" & fil.Name) & ".png" & J)
     objDOS.ExecuteCommand
     Kill BuildPath & "\" & fso.GetBaseName(BuildPath & "\" & fil.Name) & ".pvr"
     Kill BuildPath & "\" & fil.Name
  End If
  If tkt = "gltf" Then
     f = FreeFile
     Open BuildPath & "\" & fil.Name For Binary As #f
         tkt = Space$(LOF(f))
         Get #f, , tkt
     Close #f
     tkt = Replace(tkt, ".ktx", ".png")
     tkt = Replace(tkt, "/ktx", ".png")
     f = FreeFile
     Open BuildPath & "\" & fil.Name For Output As #f
          Print #f, tkt
     Close #f
  End If
Next
Beep
txtOutputs.Text = txtOutputs.Text & vbNewLine & "Done Convert!" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)

End Sub

Private Sub lvButtons_H12_Click()

If lvButtons_H12.Value = False Then
   Call SetWindowPos(Form1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
Else
   Call SetWindowPos(Form1.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End If

End Sub

Private Sub Timer3_Timer()

Drag (Command)
'Pause 0.3
Call lvButtons_H4_Click
Mak = True
Call Command5_Click
Tod
Timer3.Enabled = False

End Sub

Private Sub Command15_Click()

On Error Resume Next

If Dir(BuildPath & "\*.*") <> "" Then
   If Question(apppath & "\files found in Build-Folder!" & vbNewLine & "Delete " & apppath & "\files?") = True Then
      Kill BuildPath & "\*.*"
      txtOutputs.Text = txtOutputs.Text & vbNewLine & vbNewLine & "Build folder cleaned!" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
      Beep
   End If
End If
Call Form7.Show(vbModal)
If start_pano = True Then
   start_pano = False
   t1 = ""
   Call Command3_Click
End If

End Sub

Private Sub Check8_Click()

Form7.lvButtons_H4.Value = Check8.Value

End Sub

Private Sub Check10_Click()

On Error Resume Next

'Dim erme As String
'Dim ipget As Boolean
'
'On Error Resume Next
'
'J = Chr$(34)
'ipget = False
'qip = ""
'
'If adb = 1 Then
'   If Check10.Value = False And wcon = True Then
'      adb = 0
'      Label11.Caption = "ADB Stop!"
'      Label22.ForeColor = vbWhite
'      txtOutputs.Text = txtOutputs.Text & vbNewLine & vbNewLine & "Kill ADB Server!" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
'      Pause (0.2)
'      objDOS.CommandLine = (apppath & "\files\adb.exe kill-server")
'      objDOS.ExecuteCommand
'      txtOutputs.Text = txtOutputs.Text & vbNewLine & vbNewLine & "Wireless ADB Disconnected!" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
'      wcon = False
'      Message "Wireless ADB Disconnected!"
'      Exit Sub
'   Else
'      txtOutputs.Text = txtOutputs.Text & vbNewLine & vbNewLine & "Kill ADB Server!" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
'      Pause (0.2)
'      wcon = False
'      adb = 0
'      objDOS.CommandLine = (apppath & "\files\adb.exe kill-server")
'      objDOS.ExecuteCommand
'      txtOutputs.Text = txtOutputs.Text & vbNewLine & vbNewLine & "ADB Disconnected!" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
'   End If
'End If
'If adb = 1 And wcon = False Then
'   txtOutputs.Text = txtOutputs.Text & vbNewLine & vbNewLine & "Kill old ADB Server!" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
'   Pause (0.2)
'   objDOS.CommandLine = (apppath & "\files\adb.exe kill-server")
'   objDOS.ExecuteCommand
'End If
'If GetINISetting("QuestIP", "Adress", apppath & "\files\config.ini") <> "" Then
'   qip = GetINISetting("QuestIP", "Adress", apppath & "\files\config.ini")
'   If (GetRTTAndHopCount(inet_addr(qip), 0, 20, 200) = 1) = True Then
'       ipget = True
'       GoTo conip
'   End If
'End If
'Message "Connect USB-Cable! " & qip & vbNewLine & "(Wake up Quest from StandBy)"
'txtOutputs.Text = txtOutputs.Text & vbNewLine & vbNewLine & "Try to obtain Quest IP... Please Wait or Exit DOS Window when adb is freezed!" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
'Label11.Caption = "Wait!"
'Pause (0.2)
''qip = ShellExecuteCapture(apppath & "\files\adb.exe shell ip route")
'qip = ShellRun(apppath & "\files\adb.exe shell ip route", 6)
'
'conip:
'
'If InStr(1, qip, "src", 0) <> 0 Or ipget = True Then
'   If ipget = False Then qip = Trim(Left$(Mid$(qip, InStr(1, qip, "src", 0) + 4, Len(qip)), Len(Mid$(qip, InStr(1, qip, "src", 0) + 4, Len(qip))) - 2))
'   txtOutputs.Text = txtOutputs.Text & vbNewLine & vbNewLine & "Try to connect to Quest with IP: " & qip & "  Please Wait or Exit DOS Window when adb is freezed!" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
'   Label11.Caption = "Wait!"
'   Pause (0.2)
'   erme = ShellRun(apppath & "\files\adb.exe tcpip " & GetINISetting("QuestIP", "Port", apppath & "\files\config.ini"), 4)
'   'erme = ShellExecuteCapture(apppath & "\files\adb.exe tcpip " & GetINISetting("QuestIP", "Port", apppath & "\files\config.ini"))
'   erme = ShellRun(apppath & "\files\adb.exe connect " & qip & ":" & GetINISetting("QuestIP", "Port", apppath & "\files\config.ini"), 5)
'   'erme = ShellExecuteCapture(apppath & "\files\adb.exe connect " & qip & ":" & GetINISetting("QuestIP", "Port", apppath & "\files\config.ini"))
'   If InStr(1, erme, "connected", 0) <> 0 Then
'      If ipget = False Then
'         Beep
'         Message "Connected to " & qip & vbNewLine & "Remove USB-Cable Please!"
'         adb = 1
'         PutINISetting "QuestIP", "Adress", qip, apppath & "\files\config.ini"
'      Else
'         Beep
'         Message "Connected to " & qip
'         adb = 1
'      End If
'      wcon = True
'      Check10.Value = True
'      Label11.Caption = "Connected!"
'      Pause 3
'      Label22.ForeColor = RGB(130, 255, 130)
'   Else
'      MessageBeep (16)
'      'Message "Can´t Connect!" & vbNewLine & "Quest in Deep-StandBy?"
'      Message "Can´t Connect! Quest in Deep-StandBy?" & vbNewLine & "Maybe try to connect with USB-Cable!", True
'      Check10.Value = False
'      Label11.Caption = "Error!"
'   End If
'Else
'   MessageBeep (16)
'   Message "Can´t obtain Quest IP!" & vbNewLine & "Quest in Deep-StandBy?"
'   Check10.Value = False
'   Label11.Caption = "Error!"
'End If

End Sub

Private Sub Check16_Click()

On Error Resume Next

If Check16.Value = True Then
   PutINISetting "Save", "ADBKill", "1", apppath & "\files\config.ini"
Else
   PutINISetting "Save", "ADBKill", "0", apppath & "\files\config.ini"
End If
Pause (0.5)
Command2_Click

End Sub

Private Sub Check18_Click()

On Error Resume Next

If Check18.Value = True Then
   PutINISetting "Save", "WiFiAuto", "1", apppath & "\files\config.ini"
Else
   PutINISetting "Save", "WiFiAuto", "0", apppath & "\files\config.ini"
End If
Pause (0.5)
Command2_Click

End Sub

Private Sub Command13_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error Resume Next

Form4.Show (vbModal)
Command2_Click

End Sub

Private Sub Check11_Click()

On Error Resume Next

If Check11.Value = True Then
   PutINISetting "Save", "WindowPos", "1", apppath & "\files\config.ini"
Else
   PutINISetting "Save", "WindowPos", "0", apppath & "\files\config.ini"
End If
Pause (0.5)
Command2_Click

End Sub

Private Sub Command12_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error Resume Next

If Dir(BuildPath & "\*.*") <> "" Then
   Kill BuildPath & "\*.*"
   txtOutputs.Text = txtOutputs.Text & vbNewLine & vbNewLine & "Build folder cleaned!" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
   Beep
End If

End Sub

Private Sub Check14_Click()

On Error Resume Next

If Check11.Value = True Then
   PutINISetting "Save", "TextureDelete", "1", apppath & "\files\config.ini"
Else
   PutINISetting "Save", "TextureDelete", "0", apppath & "\files\config.ini"
End If
Pause (0.5)
Command2_Click

End Sub

Private Sub Check12_Click()

On Error Resume Next

If Check12.Value = True Then
   PutINISetting "Save", "ButtonState", "1", apppath & "\files\config.ini"
Else
   PutINISetting "Save", "ButtonState", "0", apppath & "\files\config.ini"
   PutINISetting "CheckValue", "Check", "010110", apppath & "\files\config.ini"
End If
Pause (0.5)
Command2_Click

End Sub

Private Sub Command10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim cl As Long
Dim ctrl As Control

On Error Resume Next

cl = ShowColorDialog(Me.hwnd, True, Form1.Check1.ForeColor)
If cl = -1 Then Command2_Click: Exit Sub
PutINISetting "Color", "ConsoleColor", HexIt(cl), apppath & "\files\config.ini"
txtOutputs.ForeColor = cl
Command2_Click
Me.Refresh

End Sub

Private Sub Command11_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim cl As Long
Dim ctrl As Control

On Error Resume Next

PutINISetting "Color", "HoverBackColor", "&H00BF1675&", apppath & "\files\config.ini"
PutINISetting "Color", "TitleBarColor", "&H00404040&", apppath & "\files\config.ini"
PutINISetting "Color", "ConsoleColor", "&H00FF80FF&", apppath & "\files\config.ini"
cl = HTC(GetINISetting("Color", "HoverBackColor", apppath & "\files\config.ini"))
For Each ctrl In Form1
    If ctrl.HoverBackColor <> "" Then
       ctrl.HoverBackColor = cl
       ctrl.CheckDownColor = cl
       ctrl.Enabled = False
       ctrl.Enabled = True
    End If
Next
If Lux(cl) > 120 Then
   Command1.HoverForeColor = vbBlack: Command2.HoverForeColor = vbBlack: Command3.HoverForeColor = vbBlack
   Command4.HoverForeColor = vbBlack: Command5.HoverForeColor = vbBlack: Command6.HoverForeColor = vbBlack
   Command7.HoverForeColor = vbBlack: Command8.HoverForeColor = vbBlack: Command9.HoverForeColor = vbBlack
   Command10.HoverForeColor = vbBlack: Command11.HoverForeColor = vbBlack
Else
   Command1.HoverForeColor = vbWhite: Command2.HoverForeColor = vbWhite: Command3.HoverForeColor = vbWhite
   Command4.HoverForeColor = vbWhite: Command5.HoverForeColor = vbWhite: Command6.HoverForeColor = vbWhite
   Command7.HoverForeColor = vbWhite: Command8.HoverForeColor = vbWhite: Command9.HoverForeColor = vbWhite
   Command10.HoverForeColor = vbWhite: Command11.HoverForeColor = vbWhite
End If
cl = HTC(GetINISetting("Color", "TitleBarColor", apppath & "\files\config.ini"))
Boarder1.BackColor = cl
Label21.BackColor = cl
Picture2.BackColor = cl
Picture3.BackColor = cl
Picture4.BackColor = cl
Picture5.BackColor = cl
Picture6.BackColor = cl
cl = HTC(GetINISetting("Color", "ConsoleColor", apppath & "\files\config.ini"))
txtOutputs.ForeColor = cl

Command2_Click
Me.Refresh
txtOutputs.Text = txtOutputs.Text & vbNewLine & vbNewLine & "Set default Settings!" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)

End Sub

Private Sub DropList1_Closed()

On Error Resume Next

If Combo1.Text = "" Then Exit Sub
lvButtons_H.Caption = Combo1.Text: Form7.lvButtons_H.Caption = Combo1.Text

End Sub

Function check_build() As Boolean ' true is everything is ok

Dim cb1 As Long
Dim cb2 As Long
Dim fso As New FileSystemObject
Dim fil As File
Dim tkt As String

On Error Resume Next

For Each fil In fso.GetFolder(BuildPath).Files
  tkt = fso.GetExtensionName(BuildPath & "\" & fil.Name)
  If tkt = "jpg" Or tkt = "jpeg" Or tkt = "png" Or tkt = "bmp" Then cb1 = 1
  If tkt = "ktx" Then cb2 = 1
Next

If (cb1 + cb2) = 2 Then
    check_build = False
Else
    check_build = True
End If

End Function

Private Sub Command5_Click()

On Error Resume Next

Dim es As String
Dim sa As String
Dim i As Integer
Dim fin(8) As String
Dim MyPath As String
Dim pathstr As String

If Label9.Caption = "" And Check2.Value = False Then
   MessageBeep (16)
   Message "No Audio file!"
   Exit Sub
End If
MyPath = Dir(BuildPath & "\")
Do Until MyPath = vbNullString
        If Mid(MyPath, Len(MyPath) - 4) = ".gltf" Then
            idr2 = Left$(MyPath, Len(MyPath) - 5)
        End If
    MyPath = Dir
Loop
Call Form6.Show(vbModal)
If create(0, 1) = "0" Then Exit Sub
fin(1) = fin2
If create(0, 0) = "0" Then Exit Sub
fin(2) = fin2
If create(1, 1) = "0" Then Exit Sub
fin(3) = fin2
If create(1, 0) = "0" Then Exit Sub
fin(4) = fin2
If create(0, 2) = "0" Then Exit Sub
fin(5) = fin2
If create(1, 2) = "0" Then Exit Sub
fin(6) = fin2
If create(0, 3) = "0" Then Exit Sub
fin(7) = fin2
If create(1, 3) = "0" Then Exit Sub
fin(8) = fin2
comp1 = False

If GetINISetting("Save", "Pack", apppath & "\files\config.ini") = "1" Then
      If Command = "" Then
      sa = save2(idr2 & ".zip", apppath)
   Else
      sa = dest & idr2 & ".zip"
   End If
   If sa = "" Then GoTo ende
   If LCase(Right$(sa, 3)) <> "zip" Then sa = sa & ".zip"
   objDOS.CommandLine = (apppath & "\files\7za.exe a " & J & sa & J & " " & J & fin(1) & J & " " & J & fin(2) & J & " " & J & fin(3) & J & " " & J & fin(4) & J & " " & J & fin(5) & J & " " & J & fin(6) & J & " " & J & fin(7) & J & " " & J & fin(8) & J)
   objDOS.ExecuteCommand
   If InStrRev(sa, "\") > 0 Then
      pathstr = Left$(sa, InStrRev(sa, "\"))
   Else
      pathstr = ""
   End If
    If Dir(fin(1)) <> "" Then Kill fin(1)
    If Dir(fin(2)) <> "" Then Kill fin(2)
    If Dir(fin(3)) <> "" Then Kill fin(3)
    If Dir(fin(4)) <> "" Then Kill fin(4)
    If Dir(fin(5)) <> "" Then Kill fin(5)
    If Dir(fin(6)) <> "" Then Kill fin(6)
    If Dir(fin(7)) <> "" Then Kill fin(7)
    If Dir(fin(8)) <> "" Then Kill fin(8)
Else
   If Command = "" Then
      sa = BrowseForFolder(Me.hwnd, "Select destination Folder" & vbNewLine & "Or choose Cancel for Main App folder")
   Else
      sa = dest
   End If
   If sa <> "" Then
      If sa <> apppath Then
         FileCopy fin(1), sa & "\" & ExtractFile(fin(1)): FileCopy fin(2), sa & "\" & ExtractFile(fin(2)): FileCopy fin(3), sa & "\" & ExtractFile(fin(3))
         FileCopy fin(4), sa & "\" & ExtractFile(fin(4)): FileCopy fin(5), sa & "\" & ExtractFile(fin(5)): FileCopy fin(6), sa & "\" & ExtractFile(fin(6))
         FileCopy fin(7), sa & "\" & ExtractFile(fin(7)): FileCopy fin(8), sa & "\" & ExtractFile(fin(8))
         If Dir(fin(1)) <> "" Then Kill fin(1)
         If Dir(fin(2)) <> "" Then Kill fin(2)
         If Dir(fin(3)) <> "" Then Kill fin(3)
         If Dir(fin(4)) <> "" Then Kill fin(4)
         If Dir(fin(5)) <> "" Then Kill fin(5)
         If Dir(fin(6)) <> "" Then Kill fin(6)
         If Dir(fin(7)) <> "" Then Kill fin(7)
         If Dir(fin(8)) <> "" Then Kill fin(8)
      End If
      pathstr = sa
   Else
      pathstr = apppath
   End If
End If
txtOutputs.Text = txtOutputs.Text & vbNewLine & "Create Release finished! " & Time & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
Label11.Caption = "DONE!": Beep
snd = PlaySound(apppath & "\files\gong.wav", ByVal 0&, &H20000 Or &H1)
If Command = "" Then ShellExecute hwnd, "open", pathstr, vbNullString, vbNullString, 1
'Timer1.Enabled = False
'Command1.Enabled = False
'Command3.Enabled = False
'Command5.Enabled = False

On Error Resume Next

Kill apppath & "\files\tmp\scene.zip"
Kill apppath & "\files\tmp\tmp.apk"
Kill apppath & "\files\ClassicHome\assets\scene.zip"
Kill apppath & "\files\WinterLodge\assets\scene.zip"
Kill apppath & "\files\SpaceStation\assets\scene.zip"
Kill apppath & "\files\CyberCity\assets\scene.zip"
If GetINISetting("Save", "TextureDelete", apppath & "\files\config.ini") = "1" Then
   If Dir(apppath & "\files\texture_tmp\*.*") <> "" Then Kill apppath & "\files\texture_tmp\*.*"
   RmDir apppath & "\files\texture_tmp"
End If
PutINISetting "WindowPos", "Left", Form1.Left, apppath & "\files\config.ini"
PutINISetting "WindowPos", "Top", Form1.Top, apppath & "\files\config.ini"
PutINISetting "CheckValue", "Check", GetCheck, apppath & "\files\config.ini"
If adb <> 0 Then
If Check16.Value = False Then End
   objDOS.CommandLine = (apppath & "\files\adb.exe kill-server")
   objDOS.ExecuteCommand
End If
If renunp = True Then
   If GetINISetting("Save", "AutoClear", apppath & "\files\config.ini") = "1" Then
      If Dir(BuildPath & "\*.*") <> "" Then Kill BuildPath & "\*.*"
      If Dir(apppath & "\_BACKGROUND_LOOP.ogg") <> "" Then Kill apppath & "\_BACKGROUND_LOOP.ogg"
   End If
End If
'Message "Done, restarting now!", True
'Shell apppath & "\" & App.EXEName
'End

Exit Sub

ende:


If Dir(fin(1)) <> "" Then Kill fin(1)
If Dir(fin(2)) <> "" Then Kill fin(2)
If Dir(fin(3)) <> "" Then Kill fin(3)
If Dir(fin(4)) <> "" Then Kill fin(4)
If Dir(fin(5)) <> "" Then Kill fin(5)
If Dir(fin(6)) <> "" Then Kill fin(6)
If Dir(fin(7)) <> "" Then Kill fin(7)
If Dir(fin(8)) <> "" Then Kill fin(8)

End Sub

Private Sub Check1_Click()

On Error Resume Next

If Check1.Value = True Then
   'Combo1.Enabled = True
   lvButtons_H.Enabled = True: Form7.lvButtons_H.Enabled = True
   If Check4.Enabled = True Then Check4.Value = True
Else
   'Combo1.Enabled = False
   lvButtons_H.Enabled = False: Form7.lvButtons_H.Enabled = False
End If
Form7.Check1.Value = Check1.Value
If Check1.Value = True Then

Else

End If

End Sub

Private Sub Check6_Click()

On Error Resume Next

If Check6.Value = False Then Check6.Value = True: Form7.lvButtons_H1.Value = True: Exit Sub
Check7.Value = False: Check0.Value = False: Check19.Value = False
Form7.lvButtons_H2.Value = False: Form7.lvButtons_H3.Value = False
Form7.lvButtons_H1.Value = True

End Sub

Private Sub Check7_Click()

On Error Resume Next

If Check7.Value = False Then Check7.Value = True: Form7.lvButtons_H2.Value = True: Exit Sub
Check6.Value = False: Check0.Value = False: Check19.Value = False
Form7.lvButtons_H1.Value = False: Form7.lvButtons_H3.Value = False
Form7.lvButtons_H2.Value = True

End Sub

Private Sub Check0_Click()

On Error Resume Next

If Check0.Value = False Then Check0.Value = True: Form7.lvButtons_H3.Value = True: Exit Sub
Check6.Value = False: Check7.Value = False: Check19.Value = False
Form7.lvButtons_H1.Value = False: Form7.lvButtons_H2.Value = False
Form7.lvButtons_H3.Value = True

End Sub

Private Sub Check19_Click()

If Check19.Value = False Then Check19.Value = True: Form7.lvButtons_H3.Value = True: Exit Sub
Check6.Value = False: Check7.Value = False: Check0.Value = False
Form7.lvButtons_H1.Value = False: Form7.lvButtons_H2.Value = False
Form7.lvButtons_H3.Value = True

'
End Sub

Private Function CompressKTX()

On Error Resume Next

Dim fso As New FileSystemObject
Dim fil As File
Dim fpat As String
Dim tkt As String
Dim f As Integer

For Each fil In fso.GetFolder(BuildPath).Files
  'MsgBox fil.Name
  'MsgBox fso.GetExtensionName(BuildPath & "\" & fil.Name)
  'MsgBox fso.GetBaseName(BuildPath & "\" & fil.Name)
  tkt = fso.GetExtensionName(BuildPath & "\" & fil.Name)
  If tkt = "jpg" Or tkt = "jpeg" Or tkt = "png" Or tkt = "bmp" Then
     txtOutputs.Text = txtOutputs.Text & vbNewLine & "Compress Texture File: " & BuildPath & "\" & fil.Name & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
     'MsgBox apppath & "\files\astcenc.exe -c " & BuildPath & "\" & fil.Name & " " & fso.GetBaseName(BuildPath & "\" & fil.Name) & ".astc" & " 8x8 -veryfast"
     'objDOS.CommandLine = (apppath & "\files\astcenc.exe -c " & BuildPath & "\" & fil.Name & " " & BuildPath & "\" & fso.GetBaseName(BuildPath & "\" & fil.Name) & ".astc" & " 8x8 -veryfast")
     'objDOS.ExecuteCommand
     'MsgBox apppath & "\files\img2ktx.exe -o " & BuildPath & "\" & fso.GetBaseName(BuildPath & "\" & fil.Name) & ".ktx" & " -m -f ASTC8x8 " & BuildPath & "\" & fso.GetBaseName(BuildPath & "\" & fil.Name) & ".astc"
     
     If lvButtons_H9.Value = True Then
        objDOS.CommandLine = (apppath & "\files\img2ktx.exe -o " & BuildPath & "\" & fso.GetBaseName(BuildPath & "\" & fil.Name) & ".ktx" & " -m -f ASTC8x8 " & BuildPath & "\" & fil.Name)
     Else
        objDOS.CommandLine = (apppath & "\files\img2ktx.exe -o " & BuildPath & "\" & fso.GetBaseName(BuildPath & "\" & fil.Name) & ".ktx" & " -f ASTC8x8 " & BuildPath & "\" & fil.Name)
     End If
     objDOS.ExecuteCommand
     'Kill BuildPath & "\" & fso.GetBaseName(BuildPath & "\" & fil.Name & ".astc")
     'Kill BuildPath & "\" & fso.GetBaseName(BuildPath & "\" & fso.GetExtensionName(BuildPath & "\" & fil.Name))
     Kill BuildPath & "\" & fil.Name
  End If
  If tkt = "gltf" Then
     f = FreeFile
     Open BuildPath & "\" & fil.Name For Binary As #f
         tkt = Space$(LOF(f))
         Get #f, , tkt
     Close #f
     tkt = Replace(tkt, ".jpg", ".ktx")
     tkt = Replace(tkt, "/jpg", ".ktx")
     tkt = Replace(tkt, ".jpeg", ".ktx")
     tkt = Replace(tkt, "/jpg", ".ktx")
     tkt = Replace(tkt, ".png", ".ktx")
     tkt = Replace(tkt, "/png", ".ktx")
     tkt = Replace(tkt, "/bmp", ".ktx")
     f = FreeFile
     Open BuildPath & "\" & fil.Name For Output As #f
          Print #f, tkt
     Close #f
  End If
Next

End Function

Private Sub Command3_Click()

On Error Resume Next

Dim ap As String
Dim ap1 As String
Dim t As String
Dim MyPath As String
Dim idr As String
Dim an As String
Dim ie As Boolean

com3_Self = True
' apk prüfung build ordner!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
ie = False
J = Chr$(34)
If Dir$(BuildPath & "\*.*") = vbNullString Then
   MessageBeep (16)
   Message "Put your .gltf model-files in dir .\Build"
   Exit Sub
End If
If Dir$(BuildPath & "\*.glb") <> "" Then
   MessageBeep (16)
   Message "GLB file found in .\Build, Error!" & vbNewLine & "Choose glTF separate in Blender!"
   Exit Sub
End If
If Dir$(BuildPath & "\*.apk") <> "" Then
   MessageBeep (16)
   Message "APK file found in .\Build, Error!"
   Exit Sub
End If
If Dir$(BuildPath & "\*.zip") <> "" Then
   MessageBeep (16)
   Message "ZIP file found in .\Build, Error!"
   Exit Sub
End If
If Dir$(BuildPath & "\*.mp3") <> "" Then
   MessageBeep (16)
   Message "mp3 file found in .\Build, Error!"
   Exit Sub
End If
If Dir$(BuildPath & "\*.ogg") <> "" Then
   MessageBeep (16)
   Message "ogg file found in .\Build, Error!"
   Exit Sub
End If
If Dir$(BuildPath & "\*.wav") <> "" Then
   MessageBeep (16)
   Message "wav file found in .\Build, Error!"
   Exit Sub
End If
If Dir$(BuildPath & "\*.wma") <> "" Then
   MessageBeep (16)
   Message "wma file found in .\Build, Error!"
   Exit Sub
End If
If Dir$(BuildPath & "\*.zip") <> "" Then
   MessageBeep (16)
   Message "ZIP file found in .\Build, Error!"
   Exit Sub
End If
If Dir$(BuildPath & "\*.fla") <> "" Then
   MessageBeep (16)
   Message "fla file found in .\Build, Error!"
   Exit Sub
End If
If Dir$(BuildPath & "\*.pcm") <> "" Then
   MessageBeep (16)
   Message "pcm file found in .\Build, Error!"
   Exit Sub
End If
If Dir$(BuildPath & "\*.ovrscene") <> "" Then
   MessageBeep (16)
   Message ".ovrscene file found in .\Build, Error!"
   Exit Sub
End If
If Dir$(BuildPath & "\*.glTF") = "" Then
   Exit Sub
End If
If CountFiles(BuildPath & "\*.gltf") > 1 Then
   MessageBeep (16)
   Message "More than one .glTF file found in .\Build, Error!"
   Exit Sub
End If

If Dir(apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene") <> "" Then Kill apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene"
If Dir(apppath & "\files\tmp\scene.zip") <> "" Then Kill apppath & "\files\tmp\scene.zip"
If Dir(apppath & "\files\tmp\_BACKGROUND_LOOP.ogg") <> "" Then Kill apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"

If Check10.Value = True Then CompressKTX
If check_build = False Then Message ".JPG/.PNG and .KTX present in same folder!"

weit3:

If Check2.Value = True Then
   FileCopy apppath & "\files\default.ogg", apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"
   GoTo tell
End If
If Check3.Value = True Then
   FileCopy apppath & "\files\silent.ogg", apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"
   GoTo tell
End If
If LCase(Right$(Label9.Caption, 3)) = "ogg" Then
   If Check4.Value = False And Check1.Value = False Then
      FileCopy aud, apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"
      GoTo tell
   End If
End If

If Label9.Caption = "" And Check2.Value = False And Check3.Value = False Then GoTo killer
If Check1.Value = True Then
   txtOutputs.Text = txtOutputs.Text & vbNewLine & vbNewLine & "Encode Audio File..." & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
   Pause (0.2)
   objDOS.CommandLine = (apppath & "\files\sox.exe -S " & J & aud & J & " -C 3 " & J & apppath & "\files\tmp\_BACKGROUND_LOOP.ogg" & J & " vol -" & lvButtons_H.Caption & " dB")
   objDOS.ExecuteCommand
Else
   txtOutputs.Text = txtOutputs.Text & vbNewLine & vbNewLine & "Encode Audio File..." & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
   Pause (0.2)
   objDOS.CommandLine = (apppath & "\files\sox.exe -S " & J & aud & J & " -C 3 " & J & apppath & "\files\tmp\_BACKGROUND_LOOP.ogg" & J)
   objDOS.ExecuteCommand
End If

killer:

If Dir(apppath & "\files\tmpz.apk") <> "" Then Kill apppath & "\files\tmpz.apk"
If Dir(apppath & "\files\tmp.apk") <> "" Then Kill apppath & "\files\tmp.apk"
If Dir(apppath & "\files\tmp.zip") <> "" Then Kill apppath & "\files\tmp.apk"
If Dir(apppath & "\files\scene.zip") <> "" Then Kill apppath & "\files\scene.zip"

If Dir(apppath & "\files\tmp\tmpz.apk") <> "" Then Kill apppath & "\files\tmp\tmpz.apk"
If Dir(apppath & "\files\tmp\tmp.apk") <> "" Then Kill apppath & "\files\tmp\tmp.apk"
If Dir(apppath & "\files\tmp\tmp.zip") <> "" Then Kill apppath & "\files\tmp\tmp.apk"
If Dir(apppath & "\files\tmp\scene.zip") <> "" Then Kill apppath & "\files\tmp\scene.zip"
If Dir(apppath & "\files\tmp\temp_ec.wav") <> "" Then Kill apppath & "\files\tmp\temp_ec.wav"

tell:
'--------------------------------------------------------------
objDOS.CommandLine = (apppath & "\files\7za.exe a " & J & apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene.zip" & J & " " & J & BuildPath & "\*" & J)
objDOS.ExecuteCommand
Name apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene.zip" As apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene"
If Dir(apppath & "\files\tmp\_BACKGROUND_LOOP.ogg") = "" Then FileCopy apppath & "\files\silent.ogg", apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"
objDOS.CommandLine = (apppath & "\files\7za.exe a " & J & apppath & "\files\tmp\scene.zip" & J & " " & J & apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene" & J & " " & J & apppath & "\files\tmp\_BACKGROUND_LOOP.ogg" & J)
objDOS.ExecuteCommand
If Dir(apppath & "\files\tmp\_BACKGROUND_LOOP.ogg") <> "" Then Kill apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"
Kill apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene"
If Check6.Value = True Then
   ap = apppath & "\files\WinterLodge\assets\"
   ap1 = apppath & "\files\WinterLodge"
   an = "WinterLodge"
End If
If Check7.Value = True Then
   ap = apppath & "\files\ClassicHome\assets\"
   ap1 = apppath & "\files\ClassicHome"
   an = "ClassicHome"
End If
If Check0.Value = True Then
   ap = apppath & "\files\SpaceStation\assets\"
   ap1 = apppath & "\files\SpaceStation"
   an = "SpaceStation"
End If
If Check19.Value = True Then
   ap = apppath & "\files\CyberCity\assets\"
   ap1 = apppath & "\files\CyberCity"
   an = "CyberCity"
End If

GoTo fastbuild

'----------------------------------------------------------------------------
FileCopy apppath & "\files\tmp\scene.zip", ap & "scene.zip"
objDOS.CommandLine = (java & " -Xmx1024m -jar " & J & apppath & "\files\apktool_2.3.4.jar" & J & " b -f -o " & J & apppath & "\files\tmp\tmp.apk" & J & " " & J & ap1 & J)
objDOS.ExecuteCommand
txtOutputs.Text = txtOutputs.Text & vbNewLine & "Zipalign APK-file" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
objDOS.CommandLine = (apppath & "\files\zipalign.exe -f 4 " & J & apppath & "\files\tmp\tmp.apk" & J & " " & J & apppath & "\files\tmp\tmpz.apk" & J)
objDOS.ExecuteCommand
txtOutputs.Text = txtOutputs.Text & vbNewLine & "Sign APK-file" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
objDOS.CommandLine = (java & " -Xmx1024m -jar " & J & apppath & "\files\ApkSigner.jar" & J & " sign  --key " & J & apppath & "\files\apkeasytool.pk8" & J & " --cert " & J & _
     apppath & "\files\apkeasytool.pem" & J & " --out " & J & apppath & "\files\tmp\tmpz.apk" & J & " " & J & apppath & "\files\tmp\tmpz.apk" & J)
objDOS.ExecuteCommand

MyPath = Dir(BuildPath & "\")
Do Until MyPath = vbNullString
        If Mid(MyPath, Len(MyPath) - 4) = ".gltf" Then
            idr = Left$(MyPath, Len(MyPath) - 5)
        End If
    MyPath = Dir
Loop

FileCopy apppath & "\files\tmp\tmpz.apk", apppath & "\" & idr & "." & an & ".apk"
GoTo insta
'-----------------------------------------------------

fastbuild:

FileCopy apppath & "\files\" & an & ".zip", apppath & "\files\tmp.zip"
FileCopy apppath & "\files\tmp\scene.zip", apppath & "\files\scene.zip"
ChDir apppath
ChDrive Left$(apppath, 2)
objDOS.CommandLine = (apppath & "\files\7za.exe a " & J & apppath & "\files\tmp.zip" & J & " files\scene.zip" & J)
objDOS.ExecuteCommand
objDOS.CommandLine = (apppath & "\files\7za.exe rn " & J & apppath & "\files\tmp.zip" & J & " files\ assets\")
objDOS.ExecuteCommand
Name apppath & "\files\tmp.zip" As apppath & "\files\tmp.apk"
txtOutputs.Text = txtOutputs.Text & vbNewLine & "Zipalign APK-file" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
objDOS.CommandLine = (apppath & "\files\zipalign.exe -f 4 " & J & apppath & "\files\tmp.apk" & J & " " & J & apppath & "\files\tmpz.apk" & J)
objDOS.ExecuteCommand
txtOutputs.Text = txtOutputs.Text & vbNewLine & "Sign APK-file" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
objDOS.CommandLine = (java & " -Xmx1024m -jar " & J & apppath & "\files\ApkSigner.jar" & J & " sign  --key " & J & apppath & "\files\apkeasytool.pk8" & J & " --cert " & J & _
     apppath & "\files\apkeasytool.pem" & J & " --out " & J & apppath & "\files\tmpz.apk" & J & " " & J & apppath & "\files\tmpz.apk" & J)
objDOS.ExecuteCommand
MyPath = Dir(BuildPath & "\")
Do Until MyPath = vbNullString
        If Mid(MyPath, Len(MyPath) - 4) = ".gltf" Then
            idr = Left$(MyPath, Len(MyPath) - 5)
        End If
    MyPath = Dir
Loop

FileCopy apppath & "\files\tmpz.apk", apppath & "\" & idr & "." & an & ".apk"



insta:

If Check8.Value = False Then GoTo nex
txtOutputs.Text = txtOutputs.Text & vbNewLine & "Try connecting to Quest for APK-install" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
objDOS.CommandLine = (apppath & "\files\adb.exe install -r " & J & apppath & "\" & idr & "." & an & ".apk" & J)
If InStr(1, objDOS.ExecuteCommand, "connect error", 0) <> 0 Then ie = True
adb = 1
'objDOS.CommandLine = (apppath & "\files\adb.exe kill-server")
'objDOS.ExecuteCommand

nex:

If Dir(apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene") <> "" Then Kill apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene"
If Dir(apppath & "\files\tmp\scene.zip") <> "" Then Kill apppath & "\files\tmp\scene.zip"
If Dir(apppath & "\files\scene.zip") <> "" Then Kill apppath & "\files\scene.zip"
If Dir(apppath & "\files\tmp\_BACKGROUND_LOOP.ogg") <> "" Then Kill apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"
If Dir(apppath & "\files\tmp\tmpz.apk") <> "" Then Kill apppath & "\files\tmp\tmpz.apk"
If Dir(apppath & "\files\tmp\tmp.apk") <> "" Then Kill apppath & "\files\tmp\tmp.apk"
If Dir(apppath & "\files\tmpz.apk") <> "" Then Kill apppath & "\files\tmpz.apk"
If Dir(apppath & "\files\tmp.apk") <> "" Then Kill apppath & "\files\tmp.apk"
If Dir(apppath & "\files\tmp.zip") <> "" Then Kill apppath & "\files\tmp.apk"
If Dir(apppath & "\files\tmp\temp_ec.wav") <> "" Then Kill apppath & "\files\tmp\temp_ec.wav"

If GetINISetting("Save", "AutoClear", apppath & "\files\config.ini") = "1" Then If Dir(BuildPath & "\*.*") <> "" Then Kill BuildPath & "\*.*"
txtOutputs.Text = txtOutputs.Text & vbNewLine & "Build APK finished! " & Time & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)

If ie = True Then
   MessageBeep (16)
   Label11.Caption = "Error!"
   ie = False
   If InStr(1, objDOS.ExecuteCommand, "more", vbTextCompare) > 0 Then
      Message "Error: USB-Cable Connected!" & vbNewLine & "Please Remove USB-Cable!"
   End If
Else
   If lvButtons_H11.Value = True Then lvButtons_H8_Click
   Label11.Caption = "Done!"
   snd = PlaySound(apppath & "\files\gong.wav", ByVal 0&, &H20000 Or &H1)
End If

End Sub

Private Sub Command4_Click()

On Error Resume Next

ShellExecute hwnd, "open", BuildPath & "", vbNullString, vbNullString, 1

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error Resume Next

Picture5.Enabled = True
Picture5.Visible = True
Picture6.Enabled = False
Picture6.Visible = False
Draw_Cross

End Sub

Private Sub Form_Resize()

On Error Resume Next

Draw_Cross
dra = 0

End Sub

Private Sub lvButtons_H_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error Resume Next

DropList1.DropDown

End Sub

Private Sub lvButtons_H1_Click()

lvButtons_H1.Value = True
lvButtons_H2.Value = False
lvButtons_H3.Value = False

End Sub

Private Sub lvButtons_H2_Click()


lvButtons_H1.Value = False
lvButtons_H2.Value = True
lvButtons_H3.Value = False


End Sub

Private Sub lvButtons_H3_Click()

lvButtons_H1.Value = False
lvButtons_H2.Value = False
lvButtons_H3.Value = True

End Sub

Private Function Minimize()

On Error Resume Next

Me.BorderStyle = 1
Me.Caption = Me.Caption

Me.WindowState = vbMinimized

End Function


Private Sub Command6_Click()

Dim cl As Long
Dim ctrl As Control
Dim col1 As Boolean

On Error Resume Next

cl = ShowColorDialog(Me.hwnd, True, Form1.Check1.HoverBackColor)
If cl = -1 Then Command2_Click: Exit Sub
PutINISetting "Color", "HoverBackColor", HexIt(cl), apppath & "\files\config.ini"
'col1 = Command1.Enabled
For Each ctrl In Form1
    If ctrl.HoverBackColor <> "" Then
       ctrl.HoverBackColor = cl
       ctrl.CheckDownColor = cl
       col1 = ctrl.Enabled
       ctrl.Enabled = False
       ctrl.Enabled = True
       ctrl.Enabled = col1
       If Lux(cl) > 120 Then
          ctrl.HoverForeColor = vbBlack
       Else
          ctrl.HoverForeColor = vbWhite
       End If
    End If
Next
For Each ctrl In Form7
    If ctrl.HoverBackColor <> "" Then
       ctrl.HoverBackColor = cl
       ctrl.CheckDownColor = cl
       col1 = ctrl.Enabled
       ctrl.Enabled = False
       ctrl.Enabled = True
       ctrl.Enabled = col1
       If Lux(cl) > 120 Then
          ctrl.HoverForeColor = vbBlack
       Else
          ctrl.HoverForeColor = vbWhite
       End If
    End If
Next
'Command1.Enabled = col1
'If Lux(cl) > 120 Then
'   Command1.HoverForeColor = vbBlack: Command2.HoverForeColor = vbBlack: Command3.HoverForeColor = vbBlack
'   Command4.HoverForeColor = vbBlack: Command5.HoverForeColor = vbBlack: Command6.HoverForeColor = vbBlack
'   Command7.HoverForeColor = vbBlack: Command8.HoverForeColor = vbBlack: Command9.HoverForeColor = vbBlack
'   Command10.HoverForeColor = vbBlack: Command11.HoverForeColor = vbBlack
'Else
'   Command1.HoverForeColor = vbWhite: Command2.HoverForeColor = vbWhite: Command3.HoverForeColor = vbWhite
'   Command4.HoverForeColor = vbWhite: Command5.HoverForeColor = vbWhite: Command6.HoverForeColor = vbWhite
'   Command7.HoverForeColor = vbWhite: Command8.HoverForeColor = vbWhite: Command9.HoverForeColor = vbWhite
'   Command10.HoverForeColor = vbWhite: Command11.HoverForeColor = vbWhite
'End If
Command2_Click
Me.Refresh

End Sub

Private Sub Command7_Click()

Dim cl As Long

On Error Resume Next

cl = ShowColorDialog(Me.hwnd, True, Boarder1.BackColor)
If cl = -1 Then Command2_Click: Exit Sub
PutINISetting "Color", "TitleBarColor", HexIt(cl), apppath & "\files\config.ini"
Boarder1.BackColor = cl
Label21.BackColor = cl
Picture2.BackColor = cl
Picture3.BackColor = cl
Picture4.BackColor = cl
Picture5.BackColor = cl
Picture6.BackColor = cl
Command2_Click

End Sub

Private Sub Command8_Click()

Dim fo6 As String

On Error Resume Next

fo6 = BrowseForFolder(Me.hwnd, "Select new Build Folder")
If fo6 = "" Then Command2_Click: Exit Sub
BuildPath = fo6
PutINISetting "Paths", "BuildPath", BuildPath, apppath & "\files\config.ini"
Command2_Click

End Sub

Private Sub Command9_Click()

On Error Resume Next

Call Form9.Show(vbModal)
Command2_Click

End Sub

Private Sub lvButtons_H4_Click()

On Error Resume Next

J = Chr$(34)
renunp = True

If Dir(BuildPath & "\*.*") <> "" Then
   Kill BuildPath & "\*.*"
End If

txtOutputs.Text = txtOutputs.Text & vbNewLine & "Extracting glTF-Files to ..\Build" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
Pause 0.2
objDOS.CommandLine = (apppath & "\files\7za.exe e " & J & patapk & J & " -o" & J & BuildPath & J & " scene.zip -r -aoa > nul")
objDOS.ExecuteCommand
objDOS.CommandLine = (apppath & "\files\7za.exe e " & J & BuildPath & "\scene.zip" & J & " -o" & J & BuildPath & J & " _WORLD_MODEL.gltf.ovrscene -r -aoa > nul")
objDOS.ExecuteCommand
objDOS.CommandLine = (apppath & "\files\7za.exe e " & J & BuildPath & "\scene.zip" & J & " -o" & J & apppath & J & " _BACKGROUND_LOOP.ogg -r -aoa > nul")
objDOS.ExecuteCommand
objDOS.CommandLine = (apppath & "\files\7za.exe e " & J & BuildPath & "\_WORLD_MODEL.gltf.ovrscene" & J & " -o" & J & BuildPath & J & " -aoa")
objDOS.ExecuteCommand

Kill BuildPath & "\scene.zip"
Kill BuildPath & "\_WORLD_MODEL.gltf.ovrscene"

Drag apppath & "\_BACKGROUND_LOOP.ogg"

Label11.Caption = "Done!"
snd = PlaySound(apppath & "\files\gong.wav", ByVal 0&, &H20000 Or &H1)
   
End Sub

Private Sub lvButtons_H5_Click()

Dim ie As Boolean

On Error Resume Next

txtOutputs.Text = txtOutputs.Text & vbNewLine & "Try connecting to Quest for APK-install" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
Pause 0.2
objDOS.CommandLine = (apppath & "\files\adb.exe install -r " & J & patapk & J)
If InStr(1, objDOS.ExecuteCommand, "connect error", 0) <> 0 Then ie = True

If ie = True Then
   MessageBeep (16)
   Label11.Caption = "Error!"
   ie = False
   If InStr(1, objDOS.ExecuteCommand, "more", vbTextCompare) > 0 Then
      Message "Error: USB-Cable Connected!" & vbNewLine & "Please Remove USB-Cable!"
   End If
Else
   Label11.Caption = "Done!"
   snd = PlaySound(apppath & "\files\gong.wav", ByVal 0&, &H20000 Or &H1)
End If

adb = 1

End Sub

Private Sub lvButtons_H6_Click()

Dim ie As Boolean
Dim tme As String
Dim ff As String
Dim pack2 As String
Dim fields() As String
Dim i As Long

On Error Resume Next

J = Chr$(34)

out_text = True
objDOS.CommandLine = (J & apppath & "\files\aapt.exe" & J & " d badging " & J & patapk & J)
ff = objDOS.ExecuteCommand
fields() = Split(ff, "'")
For i = 0 To UBound(fields)
    If InStr(Trim$(fields(i)), "package") <> 0 Then pack2 = Trim$(fields(i + 1))
Next i
txtOutputs.Text = txtOutputs.Text & vbNewLine & "Try connecting to Quest for APK-uninstall (" & pack2 & ")" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
Pause 0.2
objDOS.CommandLine = (apppath & "\files\adb.exe uninstall " & pack2)
ff = objDOS.ExecuteCommand
If InStr(1, ff, "connect error", 0) <> 0 Then ie = True
If InStr(1, ff, "cannot Connect", 0) <> 0 Then ie = True
If InStr(1, ff, "Unknown package", 0) <> 0 Then MessageBeep (16): Message "Error: Unknown package!" & vbNewLine & "APK already uninstalled?!": ie = True
If InStr(1, ff, "more", 0) <> 0 Then MessageBeep (16): Message "Error: USB-Cable Connected!" & vbNewLine & "Please Remove USB-Cable!": ie = True
If ie = True Then
   MessageBeep (16)
   Label11.Caption = "Error!"
   ie = False
Else
   Label11.Caption = "Done!"
   snd = PlaySound(apppath & "\files\gong.wav", ByVal 0&, &H20000 Or &H1)
End If

adb = 1

End Sub

Private Sub lvButtons_H7_Click()

Dim fold As String
Dim entry() As String
Dim sa As String
Dim path4 As String

On Error Resume Next

ChDir apppath
ChDrive Left$(apppath, 2)
adb = 1
fold = ShellRun(apppath & "\files\adb.exe shell ls -lt /sdcard/Oculus/VideoShots", 4, True)
entry = Split(fold, vbCrLf, , vbTextCompare)
If entry(1) = "" Then
   MessageBeep (16)
   Label11.Caption = "Error!"
   Exit Sub
End If
fold = Mid$(entry(1), InStr(1, entry(1), "com.oculus.vrshell", 1), Len(entry(1)) - InStr(1, entry(1), "com.oculus.vrshell", 1) + 1)
sa = ShellRun(apppath & "\files\adb.exe pull /sdcard/Oculus/VideoShots/" & fold & " " & Chr$(34) & apppath & "\files\tmp" & Chr$(34), 4)
path4 = save3(fold, apppath)
If path4 = "" Then path4 = apppath
FileCopy apppath & "\files\tmp\" & fold, path4
Kill apppath & "\files\tmp\" & fold
Beep
ShellExecute hwnd, "open", path4, vbNullString, vbNullString, 1
Reset

End Sub

Private Sub lvButtons_H8_Click()

txtOutputs.Text = txtOutputs.Text & vbNewLine & "Delete cache and initalize sound on new environments" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
objDOS.CommandLine = (apppath & "\files\adb.exe shell pm clear com.oculus.vrshell")
objDOS.ExecuteCommand
txtOutputs.Text = txtOutputs.Text & vbNewLine & "Restart Oculus Shell" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
objDOS.CommandLine = (apppath & "\files\adb.exe shell pm clear com.oculus.vrshell.desktop")
objDOS.ExecuteCommand
txtOutputs.Text = txtOutputs.Text & vbNewLine & "Done!" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
Label11.Caption = "Done!"
snd = PlaySound(apppath & "\files\gong.wav", ByVal 0&, &H20000 Or &H1)
adb = 1

End Sub

Private Sub Picture2_DblClick()

On Error Resume Next

Kill apppath & "\files\tmp\scene.zip"
Kill apppath & "\files\tmp\tmp.apk"
Kill apppath & "\files\ClassicHome\assets\scene.zip"
Kill apppath & "\files\WinterLodge\assets\scene.zip"
Kill apppath & "\files\SpaceStation\assets\scene.zip"
If GetINISetting("Save", "TextureDelete", apppath & "\files\config.ini") = "1" Then
   If Dir(apppath & "\files\texture_tmp\*.*") <> "" Then Kill apppath & "\files\texture_tmp\*.*"
   RmDir apppath & "\files\texture_tmp"
End If
PutINISetting "WindowPos", "Left", Form1.Left, apppath & "\files\config.ini"
PutINISetting "WindowPos", "Top", Form1.Top, apppath & "\files\config.ini"
PutINISetting "CheckValue", "Check", GetCheck, apppath & "\files\config.ini"
If adb = 0 Then End
If Check16.Value = False Then End
objDOS.CommandLine = (apppath & "\files\adb.exe kill-server")
objDOS.ExecuteCommand
End

End Sub

Private Sub Picture3_Click()

On Error Resume Next

Kill apppath & "\files\tmp\scene.zip"
Kill apppath & "\files\tmp\tmp.apk"
Kill apppath & "\files\ClassicHome\assets\scene.zip"
Kill apppath & "\files\WinterLodge\assets\scene.zip"
Kill apppath & "\files\SpaceStation\assets\scene.zip"
If GetINISetting("Save", "TextureDelete", apppath & "\files\config.ini") = "1" Then
   If Dir(apppath & "\files\texture_tmp\*.*") <> "" Then Kill apppath & "\files\texture_tmp\*.*"
   RmDir apppath & "\files\texture_tmp"
End If
PutINISetting "WindowPos", "Left", Form1.Left, apppath & "\files\config.ini"
PutINISetting "WindowPos", "Top", Form1.Top, apppath & "\files\config.ini"
PutINISetting "CheckValue", "Check", GetCheck, apppath & "\files\config.ini"
If adb = 0 Then End
If Check16.Value = False Then End
objDOS.CommandLine = (apppath & "\files\adb.exe kill-server")
objDOS.ExecuteCommand
End

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error Resume Next

Kill apppath & "\files\tmp\scene.zip"
Kill apppath & "\files\tmp\tmp.apk"
Kill apppath & "\files\ClassicHome\assets\scene.zip"
Kill apppath & "\files\WinterLodge\assets\scene.zip"
Kill apppath & "\files\SpaceStation\assets\scene.zip"
If GetINISetting("Save", "TextureDelete", apppath & "\files\config.ini") = "1" Then
   If Dir(apppath & "\files\texture_tmp\*.*") <> "" Then Kill apppath & "\files\texture_tmp\*.*"
   RmDir apppath & "\files\texture_tmp"
End If
PutINISetting "WindowPos", "Left", Form1.Left, apppath & "\files\config.ini"
PutINISetting "WindowPos", "Top", Form1.Top, apppath & "\files\config.ini"
PutINISetting "CheckValue", "Check", GetCheck, apppath & "\files\config.ini"
If adb = 0 Then End
If Check16.Value = False Then End
objDOS.CommandLine = (apppath & "\files\adb.exe kill-server")
objDOS.ExecuteCommand

End Sub

Private Function Tod()

On Error Resume Next

Kill apppath & "\files\tmp\scene.zip"
Kill apppath & "\files\tmp\tmp.apk"
Kill apppath & "\files\ClassicHome\assets\scene.zip"
Kill apppath & "\files\WinterLodge\assets\scene.zip"
Kill apppath & "\files\SpaceStation\assets\scene.zip"
If GetINISetting("Save", "TextureDelete", apppath & "\files\config.ini") = "1" Then
   If Dir(apppath & "\files\texture_tmp\*.*") <> "" Then Kill apppath & "\files\texture_tmp\*.*"
   RmDir apppath & "\files\texture_tmp"
End If
PutINISetting "WindowPos", "Left", Form1.Left, apppath & "\files\config.ini"
PutINISetting "WindowPos", "Top", Form1.Top, apppath & "\files\config.ini"
PutINISetting "CheckValue", "Check", GetCheck, apppath & "\files\config.ini"
If adb = 0 Then End
If Check16.Value = False Then End
objDOS.CommandLine = (apppath & "\files\adb.exe kill-server")
objDOS.ExecuteCommand

End Function

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error Resume Next

Picture3.BackColor = Command2.HoverBackColor
Draw_Cross True, 3

End Sub

Private Sub Picture4_Click()

On Error Resume Next

Minimize

End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error Resume Next

Picture4.BackColor = Command2.HoverBackColor
Draw_Cross True, 4

End Sub

Private Function Draw_Cross(Optional hover As Boolean, Optional pic As Integer)

On Error Resume Next

Dim clx As String

clx = Boarder1.BackColor
If Lux(clx) > 120 Then
   clx = vbBlack
   Label21.ForeColor = vbBlack
Else
   clx = Command2.ForeColor
End If
Picture5.DrawWidth = 2
Picture5.Line (70, 70)-Step(160, 160), clx, B
Picture6.DrawWidth = 2
Picture6.Line (70, 70)-Step(160, 160), clx, B
If hover = False Then
   Picture3.DrawWidth = 2
   Picture3.Line (60, 60)-(210, 210), clx
   Picture3.Line (210, 60)-(60, 210), clx
   Picture4.DrawWidth = 2
   Picture4.Line (60, 210)-(230, 210), clx, B
Else
   If pic = 3 Then
      Picture3.DrawWidth = 2
      Picture3.Line (60, 60)-(210, 210), Command2.HoverForeColor
      Picture3.Line (210, 60)-(60, 210), Command2.HoverForeColor
   End If
   If pic = 4 Then
      Picture4.DrawWidth = 2
      Picture4.Line (60, 210)-(230, 210), Command2.HoverForeColor, B
   End If
End If

End Function

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error Resume Next

Picture6.Enabled = True
Picture6.Visible = True
Picture5.Enabled = False
Picture5.Visible = False

End Sub


Private Sub Picture6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error Resume Next

Picture5.Enabled = True
Picture5.Visible = True
Picture6.Enabled = False
Picture6.Visible = False

End Sub

Private Sub Timer2_Timer()

On Error Resume Next

If dra < 5 Then
   dra = dra + 1
   Draw_Cross
End If
'If dra = False Then dra = True: Draw_Cross
If Me.WindowState = 0 Then
   Me.BorderStyle = 0
   Me.Caption = Me.Caption
End If

End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error Resume Next

Draw_Cross
storedx = x
storedy = y

End Sub

Private Sub Label21_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error Resume Next

Picture3.BackColor = Boarder1.BackColor
Picture4.BackColor = Boarder1.BackColor
Picture5.BackColor = Boarder1.BackColor
Picture6.BackColor = Boarder1.BackColor
Draw_Cross
If Button = 1 Then
    Me.Left = x - storedx + Me.Left
    Me.Top = y - storedy + Me.Top
End If

End Sub


Private Sub Boarder1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error Resume Next

Draw_Cross
storedx = x
storedy = y

End Sub

Private Sub Boarder1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error Resume Next

Picture3.BackColor = Boarder1.BackColor
Picture4.BackColor = Boarder1.BackColor
Picture5.BackColor = Boarder1.BackColor
Picture6.BackColor = Boarder1.BackColor
Draw_Cross
If Button = 1 Then
    Me.Left = x - storedx + Me.Left
    Me.Top = y - storedy + Me.Top
End If

End Sub

Private Sub objDOS_ReceiveOutputs(CommandOutputs As String)

On Error Resume Next

If out_text = False Then txtOutputs.Text = txtOutputs.Text & CommandOutputs
out_text = False

End Sub

Private Sub Option2_Click()

On Error Resume Next

If Option2.Value = True Then
   Option1.Value = False
   lvButtons_H1.Visible = False
   lvButtons_H2.Visible = False
   lvButtons_H3.Visible = False
   Label18.Visible = False
   Label19.Visible = False
   Label20.Visible = False
Else
   lvButtons_H1.Visible = True
   lvButtons_H2.Visible = True
   lvButtons_H3.Visible = True
   Label18.Visible = True
   Label19.Visible = True
   Label20.Visible = True
   Option1.Value = True
End If

End Sub

Private Sub Option1_Click()

On Error Resume Next

If Option1.Value = True Then
   If Label8.Caption <> "" Then Command1.Enabled = True
   lvButtons_H1.Visible = True
   lvButtons_H2.Visible = True
   lvButtons_H3.Visible = True
   Label18.Visible = True
   Label19.Visible = True
   Label20.Visible = True
Else
   lvButtons_H1.Visible = False
   lvButtons_H2.Visible = False
   lvButtons_H3.Visible = False
   Label18.Visible = False
   Label19.Visible = False
   Label20.Visible = False
End If
If Option1.Value = True Then
   Option2.Value = False
Else
   Option2.Value = True
End If

End Sub

Private Sub Picture1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

On Error Resume Next

'Achtung!!!!!!  Gleiche routine als Drag funktion als spiegelbild vorhanden für unpack apk routine!

Dim intFile As Integer
Dim i As Long
Dim fields() As String
Dim k As String
Dim pat As String
Dim tme As String
Dim ff As String
Dim tw As Boolean

With Data
     For intFile = 1 To .Files.Count
         pat = Data.Files.Item(intFile)
     Next intFile
End With

pat = Replace(pat, Chr$(34), "")
tw = False
u = ExtractFile(pat)
k = LCase(Right$(u, 3))
If k = "apk" Then
   lvButtons_H6.Enabled = True
   lvButtons_H5.Enabled = True
   lvButtons_H4.Enabled = True
   lvButtons_H10.Enabled = True
   If Option1.Value = True Then
      Command1.Enabled = True
   Else
      If pataud <> "" Then Command1.Enabled = True
   End If
   J = Chr$(34)
   tme = txtOutputs.Text
   objDOS.CommandLine = (J & apppath & "\files\aapt.exe" & J & " d badging " & J & pat & J)
   ff = objDOS.ExecuteCommand
   txtOutputs.Text = tme & "Added APK: " & pat & vbNewLine & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
   fields() = Split(ff, "'")
   For i = 0 To UBound(fields)
       If InStr(Trim$(fields(i)), "package") <> 0 Then pack = Trim$(fields(i + 1))
   Next i
   Label8.Caption = u
   patapk = pat
   If pack = "com.oculus.environment.prod.winterlodge" Then
      lvButtons_H1.Value = False: lvButtons_H2.Value = True: lvButtons_H3.Value = False
      lvButtons_H1.Enabled = False: lvButtons_H2.Enabled = True: lvButtons_H3.Enabled = True
   End If
   If pack = "com.oculus.environment.prod.rifthome" Then
      lvButtons_H1.Value = True: lvButtons_H2.Value = False: lvButtons_H3.Value = False
      lvButtons_H1.Enabled = True: lvButtons_H2.Enabled = False: lvButtons_H3.Enabled = True
   End If
   If pack = "com.oculus.environment.prod.spacestation" Then
      lvButtons_H1.Value = True: lvButtons_H2.Value = False: lvButtons_H3.Value = False
      lvButtons_H1.Enabled = True: lvButtons_H2.Enabled = True: lvButtons_H3.Enabled = False
   End If
   Exit Sub
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
      Message "File or Audio-file type '" & k & "' not supported, sorry!"
      Exit Sub
   End If
End If
aud = pat
If k = "ogg" Then
   Check4.Enabled = True: Form7.Check4.Enabled = True
   Label10.Enabled = True: Form7.Label10.Enabled = True
End If
Check1.Enabled = True: Form7.Check1.Enabled = True
Label1.Enabled = True: Form7.Label1.Enabled = True
Label9.Caption = u
Form7.Label9.Caption = u
txtOutputs.Text = txtOutputs.Text & "Added Audio-file: " & aud & vbNewLine & vbNewLine
txtOutputs.SelStart = Len(txtOutputs.Text)
pataud = u
Check4.Enabled = False: Form7.Check4.Enabled = False
Label10.Enabled = False: Form7.Label10.Enabled = False
If k = "ogg" Then
   Check4.Enabled = True: Form7.Check4.Enabled = True
   Label10.Enabled = True: Form7.Label10.Enabled = True
End If
If patapk <> "" Then Command1.Enabled = True
Check2.Enabled = False: Form7.Check2.Enabled = False
Label4.Enabled = False: Form7.Label4.Enabled = False
Check3.Enabled = False: Form7.Check3.Enabled = False
Label5.Enabled = False: Form7.Label5.Enabled = False

End Sub

Private Function Drag(pat As String)

On Error Resume Next

'Achtung!!!!!!  Gleiche routine als Drag funktion als spiegelbild vorhanden für unpack apk routine!

Dim intFile As Integer
Dim i As Long
Dim fields() As String
Dim k As String
Dim tme As String
Dim ff As String
Dim tw As Boolean

pat = Replace(pat, Chr$(34), "")
tw = False
u = ExtractFile(pat)
k = LCase(Right$(u, 3))
If k = "apk" Then
   lvButtons_H5.Enabled = True
   lvButtons_H4.Enabled = True
   lvButtons_H10.Enabled = True
   If Option1.Value = True Then
      Command1.Enabled = True
   Else
      If pataud <> "" Then Command1.Enabled = True
   End If
   J = Chr$(34)
   tme = txtOutputs.Text
   objDOS.CommandLine = (J & apppath & "\files\aapt.exe" & J & " d badging " & J & pat & J)
   ff = objDOS.ExecuteCommand
   txtOutputs.Text = tme & "Added APK: " & pat & vbNewLine & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
   fields() = Split(ff, "'")
   For i = 0 To UBound(fields)
       If InStr(Trim$(fields(i)), "package") <> 0 Then pack = Trim$(fields(i + 1))
   Next i
   Label8.Caption = u
   patapk = pat
   If pack = "com.oculus.environment.prod.winterlodge" Then
      lvButtons_H1.Value = False: lvButtons_H2.Value = True: lvButtons_H3.Value = False
      lvButtons_H1.Enabled = False: lvButtons_H2.Enabled = True: lvButtons_H3.Enabled = True
   End If
   If pack = "com.oculus.environment.prod.rifthome" Then
      lvButtons_H1.Value = True: lvButtons_H2.Value = False: lvButtons_H3.Value = False
      lvButtons_H1.Enabled = True: lvButtons_H2.Enabled = False: lvButtons_H3.Enabled = True
   End If
   If pack = "com.oculus.environment.prod.spacestation" Then
      lvButtons_H1.Value = True: lvButtons_H2.Value = False: lvButtons_H3.Value = False
      lvButtons_H1.Enabled = True: lvButtons_H2.Enabled = True: lvButtons_H3.Enabled = False
   End If
   Exit Function
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
      Message "File or Audio-file type '" & k & "' not supported, sorry!"
      Exit Function
   End If
End If
aud = pat
If k = "ogg" Then
   Check4.Enabled = True: Form7.Check4.Enabled = True
   Label10.Enabled = True: Form7.Label10.Enabled = True
End If
Check1.Enabled = True: Form7.Check1.Enabled = True
Label1.Enabled = True: Form7.Label1.Enabled = True
Label9.Caption = u
Form7.Label9.Caption = u
txtOutputs.Text = txtOutputs.Text & "Added Audio-file: " & aud & vbNewLine & vbNewLine
txtOutputs.SelStart = Len(txtOutputs.Text)
pataud = u
Check4.Enabled = False: Form7.Check4.Enabled = False
Label10.Enabled = False: Form7.Label10.Enabled = False
If k = "ogg" Then
   Check4.Enabled = True: Form7.Check4.Enabled = True
   Label10.Enabled = True: Form7.Label10.Enabled = True
End If
If patapk <> "" Then Command1.Enabled = True
Check2.Enabled = False: Form7.Check2.Enabled = False
Label4.Enabled = False: Form7.Label4.Enabled = False
Check3.Enabled = False: Form7.Check3.Enabled = False
Label5.Enabled = False: Form7.Label5.Enabled = False

End Function

Private Sub Timer1_Timer()

On Error Resume Next

Dim t2 As String

If start_adb = True Then
   start_adb = False
    If GetINISetting("Save", "ADBKill", apppath & "\files\config.ini") = "1" Then
       Check16.Value = True
    Else
       Check16.Value = False
       If IsEXERunning("adb.exe") = True Then
          objDOS.CommandLine = (apppath & "\files\adb.exe devices")
          out_text = True
          qip = objDOS.ExecuteCommand
          If InStr(1, qip, GetINISetting("QuestIP", "Adress", apppath & "\files\config.ini"), vbTextCompare) > 0 Then
             Label22.ForeColor = RGB(130, 255, 130)
             Check10.Value = True
             adb = 1
             wcon = True
             qip = GetINISetting("QuestIP", "Adress", apppath & "\files\config.ini")
             txtOutputs.Text = txtOutputs.Text & vbNewLine & "Found active ADB-Connection (WiFi): " & qip & ":" & GetINISetting("QuestIP", "Port", apppath & "\files\config.ini") & vbNewLine & vbNewLine
             txtOutputs.SelStart = Len(txtOutputs.Text)
          Else
             qip = GetINISetting("QuestIP", "Adress", apppath & "\files\config.ini")
             adb = 1
             txtOutputs.Text = txtOutputs.Text & vbNewLine & "Found active ADB-Connection(USB) " & vbNewLine & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
          End If
       End If
    End If
End If

If GetINISetting("Save", "WiFiAuto", apppath & "\files\config.ini") = "1" And wifi_auto = True Then
   wifi_auto = False
   Call Check10_Click
End If

If Label11.Caption <> "" Then
   If Command1.Enabled = False And tp = False Then GoTo cont5
   za = za + 1
   If za > 15 Then
      Label11.Caption = ""
      za = 0
      tp = False
   End If
End If

cont5:

If wcon = True Then
   za2 = za2 + 1
   If za2 > 5 Then
      za2 = 1
      If (GetRTTAndHopCount(inet_addr(qip), 0, 20, 200) = 1) = True Then
         Label22.ForeColor = RGB(130, 255, 130)
         objDOS.CommandLine = (apppath & "\files\adb.exe shell settings get global wifi_on")
         out_text = True
         t2 = objDOS.ExecuteCommand
         If t2 = 1 Then
            Label22.ForeColor = RGB(130, 255, 130)
         Else
            Label22.ForeColor = vbWhite
            Check10.Value = False
            MessageBeep (16)
            Message "Error: Lost WiFi Connection!"
            adb = 0
            wcon = False
            objDOS.CommandLine = (apppath & "\files\adb.exe kill-server")
            objDOS.ExecuteCommand
            Exit Sub
         End If
         If InStr(1, t1, "more", vbTextCompare) > 0 Then
            MessageBeep (16)
            Message "Error: USB-Cable Connected!" & vbNewLine & "Please Remove USB-Cable!"
            Exit Sub
         End If
         'error more than one
         '1 wenn gut
         '0 wenn aus
      Else
         Label22.ForeColor = vbWhite
         Check10.Value = False
         MessageBeep (16)
         Message "Error: Lost WiFi Connection/IP!"
         adb = 0
         wcon = False
         objDOS.CommandLine = (apppath & "\files\adb.exe kill-server")
         objDOS.ExecuteCommand
         Exit Sub
      End If
   End If
End If

If Dir$(BuildPath & "\*.*") <> vbNullString Then
    Set fsx = CreateObject("Scripting.FileSystemObject")
    For Each oFile In fsx.GetFolder(BuildPath & "").Files
        If LCase(fsx.GetExtensionName(oFile.path)) = "gltf" Then
            fn = fsx.GetFileName(oFile.path)
            Exit For
        End If
    Next
    If t1 = "" Then t1 = FileDateTime(BuildPath & "\" & fn)
    'If FileDateTime(BuildPath & "\" & fn) <> t1 Then
    If DateDiff("s", t1, FileDateTime(BuildPath & "\" & fn)) > 8 Then
       t1 = FileDateTime(BuildPath & "\" & fn)
       If Check9.Value = True Then
          If Dir$(BuildPath & "\*.glTF") = "" Then
             Exit Sub
          Else
             If com3_Self = True Then Call Command3_Click
          End If
       End If
    Else
       t1 = FileDateTime(BuildPath & "\" & fn)
    End If
End If
If Dir$(BuildPath & "\*.*") = vbNullString Then
   Check6.Enabled = False
   Check0.Enabled = False
   Check7.Enabled = False
   Check8.Enabled = False
   Check9.Enabled = False
   Check10.Enabled = False
   Check13.Enabled = False
   Check19.Enabled = False
   Command3.Enabled = False
   Command5.Enabled = False
   Label13.Enabled = False
   Label14.Enabled = False
   Label15.Enabled = False
   Label16.Enabled = False
   Label17.Enabled = False
   Label22.Enabled = False
   Label25.Enabled = False
   Label31.Enabled = False
Else
   If Dir$(BuildPath & "\*.gltf") <> "" Then
      Check6.Enabled = True
      Check0.Enabled = True
      Check7.Enabled = True
      Check8.Enabled = True
      Check9.Enabled = True
      Check10.Enabled = True
      Check13.Enabled = True
      Check19.Enabled = True
      Command3.Enabled = True
      Command5.Enabled = True
      Label13.Enabled = True
      Label14.Enabled = True
      Label15.Enabled = True
      Label16.Enabled = True
      Label17.Enabled = True
      Label22.Enabled = True
      Label25.Enabled = True
      Label31.Enabled = True
      If qw = 1 Then
         qw = 0
         txtOutputs.Text = txtOutputs.Text & vbNewLine & "Found gltf-model in .\Build" & vbNewLine & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
      End If
   End If
End If

End Sub

Private Sub txtOutputs_Change()
    
On Error Resume Next

txtOutputs.SelStart = Len(txtOutputs.Text)

End Sub

Private Sub Check2_Click()

On Error Resume Next

If Check2.Value = True Then Check3.Value = False: Form7.Check3.Value = False
If Label8.Caption <> "" Then Command1.Enabled = True
Form7.Check2.Value = Check2.Value

End Sub

Private Sub Check3_Click()

On Error Resume Next

If Check3.Value = True Then Check2.Value = False: Form7.Check2.Value = False
If Label8.Caption <> "" Then Command1.Enabled = True
Form7.Check3.Value = Check3.Value

End Sub

Private Sub Check4_Click()

On Error Resume Next

If Check4.Value = False Then Check1.Value = False
Form7.Check4.Value = Check4.Value

End Sub

Private Sub Command1_Click()
'& " > " & apppath & "\files\log.txt 2>&1"
'On Error Resume Next

On Error Resume Next

Dim ap As String
Dim ap1 As String
Dim t As String

J = Chr$(34)

objDOS.CommandLine = (apppath & "\files\7za.exe e " & J & patapk & J & " -o" & J & apppath & "\files\tmp" & J & " assets\scene.zip -aoa > nul")
objDOS.ExecuteCommand

'If pataud = "" Or Option1.Value = True And Check3.Value = "0" Then GoTo switch
If pataud = "" And Check2.Value = False And Check3.Value = False Then GoTo switch


If Option2.Value = False Then GoTo switchaud
' Nur Audio tauschen:
objDOS.CommandLine = (apppath & "\files\7za.exe e " & J & apppath & "\files\tmp\scene.zip" & J & " -o" & J & apppath & "\files\tmp" & J & " -aoa")
objDOS.ExecuteCommand
If Dir(apppath & "\files\tmp\scene.zip") <> "" Then Kill apppath & "\files\tmp\scene.zip"
If Dir(apppath & "\files\tmp\_BACKGROUND_LOOP.ogg") <> "" Then Kill apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"

If Check2.Value = True Then FileCopy apppath & "\files\default.ogg", apppath & "\files\tmp\_BACKGROUND_LOOP.ogg": GoTo tell
If Check3.Value = True Then FileCopy apppath & "\files\silent.ogg", apppath & "\files\tmp\_BACKGROUND_LOOP.ogg": GoTo tell
If LCase(Right$(Label9.Caption, 3)) = "ogg" Then
   If Check4.Value = False And Check1.Value = False Then
      FileCopy aud, apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"
      GoTo tell
   End If
End If
If Check1.Value = True Then
   txtOutputs.Text = txtOutputs.Text & vbNewLine & vbNewLine & "Encode Audio File..." & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
   Pause (0.2)
   objDOS.CommandLine = (apppath & "\files\sox.exe -S " & J & aud & J & " -C 3 " & J & apppath & "\files\tmp\_BACKGROUND_LOOP.ogg" & J & " vol -" & lvButtons_H.Caption & " dB")
   objDOS.ExecuteCommand
Else
   txtOutputs.Text = txtOutputs.Text & vbNewLine & vbNewLine & "Encode Audio File..." & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
   Pause (0.2)
   objDOS.CommandLine = (apppath & "\files\sox.exe -S " & J & aud & J & " -C 3 " & J & apppath & "\files\tmp\_BACKGROUND_LOOP.ogg" & J)
   objDOS.ExecuteCommand
End If

tell: 'ohne switch
'--------------------------------------------------------------

objDOS.CommandLine = (apppath & "\files\7za.exe a " & J & apppath & "\files\tmp\scene.zip" & J & " " & J & apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene" & J & " " & J & apppath & "\files\tmp\_BACKGROUND_LOOP.ogg" & J)
objDOS.ExecuteCommand
If Dir(apppath & "\files\tmp\_BACKGROUND_LOOP.ogg") <> "" Then Kill apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"
Kill apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene"
If pack = "com.oculus.environment.prod.winterlodge" Then
   ap = apppath & "\files\WinterLodge\assets\"
   ap1 = apppath & "\files\WinterLodge"
End If
If pack = "com.oculus.environment.prod.rifthome" Then
   ap = apppath & "\files\ClassicHome\assets\"
   ap1 = apppath & "\files\ClassicHome"
End If
If pack = "com.oculus.environment.prod.spacestation" Then
   ap = apppath & "\files\SpaceStation\assets\"
   ap1 = apppath & "\files\SpaceStation"
End If
If pack = "com.oculus.environment.prod.cyberhome" Then
   ap = apppath & "\files\CyberCity\assets\"
   ap1 = apppath & "\files\CyberCity"
End If

FileCopy apppath & "\files\tmp\scene.zip", ap & "scene.zip"
objDOS.CommandLine = (java & " -Xmx1024m -jar " & J & apppath & "\files\apktool_2.3.4.jar" & J & " b -f -o " & J & apppath & "\files\tmp\tmp.apk" & J & " " & J & ap1 & J)
objDOS.ExecuteCommand
txtOutputs.Text = txtOutputs.Text & vbNewLine & "Zipalign APK-file" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
objDOS.CommandLine = (apppath & "\files\zipalign.exe -f 4 " & J & apppath & "\files\tmp\tmp.apk" & J & " " & J & apppath & "\files\tmp\tmpz.apk" & J)
objDOS.ExecuteCommand
txtOutputs.Text = txtOutputs.Text & vbNewLine & "Sign APK-file" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
objDOS.CommandLine = (java & " -Xmx1024m -jar " & J & apppath & "\files\ApkSigner.jar" & J & " sign  --key " & J & apppath & "\files\apkeasytool.pk8" & J & " --cert " & J & _
     apppath & "\files\apkeasytool.pem" & J & " --out " & J & apppath & "\files\tmp\tmpz.apk" & J & " " & J & apppath & "\files\tmp\tmpz.apk" & J)
objDOS.ExecuteCommand

On Error Resume Next

sa = save(Rename(ap1), patapk)
If sa = "" Then GoTo ende2
FileCopy apppath & "\files\tmp\tmpz.apk", sa
'" & j & "
If Check5.Value = True Then
   txtOutputs.Text = txtOutputs.Text & vbNewLine & "Try connecting to Quest for APK-install" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
   Pause (0.1)
   objDOS.CommandLine = (apppath & "\files\adb.exe install -r " & J & sa & J)
   objDOS.ExecuteCommand
   adb = 1
   'objDOS.CommandLine = (apppath & "\files\adb.exe kill-server")
   'objDOS.ExecuteCommand
End If

ende2:
GoTo ende
'---------------------------------------------------------------------

switchaud:

If Label9.Caption = "" Then
   If Check2.Value = False And Check3.Value = False Then GoTo switch
End If
objDOS.CommandLine = (apppath & "\files\7za.exe e " & J & apppath & "\files\tmp\scene.zip" & J & " -o" & J & apppath & "\files\tmp" & J & " -aoa")
objDOS.ExecuteCommand
Kill apppath & "\files\tmp\scene.zip"
Kill apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"
If Check2.Value = True Then FileCopy apppath & "\files\default.ogg", apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"
If Check3.Value = True Then FileCopy apppath & "\files\silent.ogg", apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"
If LCase(Right$(Label9.Caption, 3)) = "ogg" Then
   If Check4.Value = False And Check1.Value = False Then
      FileCopy aud, apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"
      GoTo tell
   End If
End If
If LCase(Right$(Label9.Caption, 3)) = "ogg" Then
   If Check4.Value = False And Check4.Value = False Then
      FileCopy aud, apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"
      GoTo switch
   End If
End If
If Check1.Value = True Then
   txtOutputs.Text = txtOutputs.Text & vbNewLine & vbNewLine & "Encode Audio File..." & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
   Pause (0.2)
   objDOS.CommandLine = (apppath & "\files\sox.exe -S " & J & aud & J & " -C 3 " & J & apppath & "\files\tmp\_BACKGROUND_LOOP.ogg" & J & " vol -" & lvButtons_H.Caption & " dB")
   objDOS.ExecuteCommand
Else
   txtOutputs.Text = txtOutputs.Text & vbNewLine & vbNewLine & "Encode Audio File..." & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
   Pause (0.2)
   objDOS.CommandLine = (apppath & "\files\sox.exe -S " & J & aud & J & " -C 3 " & J & apppath & "\files\tmp\_BACKGROUND_LOOP.ogg" & J)
   objDOS.ExecuteCommand
End If

'--------------------------------------------------------------------------
switch:

If Label8.Caption <> "" Then
   objDOS.CommandLine = (apppath & "\files\7za.exe a " & J & apppath & "\files\tmp\scene.zip" & J & " " & J & apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene" & J & " " & J & apppath & "\files\tmp\_BACKGROUND_LOOP.ogg" & J)
   objDOS.ExecuteCommand
End If
If lvButtons_H1.Value = True Then
   ap = apppath & "\files\WinterLodge\assets\"
   ap1 = apppath & "\files\WinterLodge"
End If
If lvButtons_H2.Value = True Then
   ap = apppath & "\files\ClassicHome\assets\"
   ap1 = apppath & "\files\ClassicHome"
End If
If lvButtons_H3.Value = True Then
   ap = apppath & "\files\SpaceStation\assets\"
   ap1 = apppath & "\files\SpaceStation"
End If
FileCopy apppath & "\files\tmp\scene.zip", ap & "scene.zip"
If Dir(apppath & "\files\tmp\_BACKGROUND_LOOP.ogg") <> "" Then Kill apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"
If Dir(apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene") <> "" Then Kill apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene"
objDOS.CommandLine = (java & " -Xmx1024m -jar " & J & apppath & "\files\apktool_2.3.4.jar" & J & " b -f -o " & J & apppath & "\files\tmp\tmp.apk" & J & " " & J & ap1 & J)
objDOS.ExecuteCommand
txtOutputs.Text = txtOutputs.Text & vbNewLine & "Zipalign APK-file" & vbNewLine & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
objDOS.CommandLine = (apppath & "\files\zipalign.exe -f 4 " & J & apppath & "\files\tmp\tmp.apk" & J & " " & J & apppath & "\files\tmp\tmpz.apk" & J)
objDOS.ExecuteCommand
txtOutputs.Text = txtOutputs.Text & vbNewLine & "Sign APK-file" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
objDOS.CommandLine = (java & " -Xmx1024m -jar " & J & apppath & "\files\ApkSigner.jar" & J & " sign  --key " & J & apppath & "\files\apkeasytool.pk8" & J & " --cert " & J & _
     apppath & "\files\apkeasytool.pem" & J & " --out " & J & apppath & "\files\tmp\tmpz.apk" & J & " " & J & apppath & "\files\tmp\tmpz.apk" & J)
objDOS.ExecuteCommand

On Error Resume Next

sa = save(Rename(ap1), patapk)
If sa = "" Then GoTo ende
FileCopy apppath & "\files\tmp\tmpz.apk", sa
'" & j & "
If Check5.Value = True Then
   txtOutputs.Text = txtOutputs.Text & vbNewLine & "Try connecting to Quest for APK-install" & vbNewLine & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
   Pause (0.1)
   objDOS.CommandLine = (apppath & "\files\adb.exe install -r " & J & sa & J)
   objDOS.ExecuteCommand
   adb = 1
   'objDOS.CommandLine = (apppath & "\files\adb.exe kill-server")
   'objDOS.ExecuteCommand
End If

ende:

If Dir(apppath & "\files\tmp\tmpz.apk") <> "" Then Kill apppath & "\files\tmp\tmpz.apk"
If Dir(apppath & "\files\tmp\tmp.apk") <> "" Then Kill apppath & "\files\tmp\tmp.apk"
If Dir(apppath & "\files\tmp\tmp.zip") <> "" Then Kill apppath & "\files\tmp\tmp.apk"
If Dir(apppath & "\files\tmp\scene.zip") <> "" Then Kill apppath & "\files\tmp\scene.zip"
If Dir(apppath & "\files\tmp\temp_ec.wav") <> "" Then Kill apppath & "\files\tmp\temp_ec.wav"
Deaktiv

End Sub

Private Function Deaktiv()


On Error Resume Next

txtOutputs.Text = txtOutputs.Text & vbNewLine & "All operations finished!" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
Label11.Caption = "DONE!"
snd = PlaySound(apppath & "\files\gong.wav", ByVal 0&, &H20000 Or &H1)
'Command1.Enabled = False
'Option1.Enabled = False
'Option2.Enabled = False
'Check1.Enabled = False
'Check2.Enabled = False
'Check3.Enabled = False
'Check4.Enabled = False
'Check5.Enabled = False
'Check6.Enabled = False
'Check0.Enabled = False
'Check7.Enabled = False
'Check8.Enabled = False
'Check9.Enabled = False
'Check10.Enabled = False

End Function

Private Sub Command2_Click()

On Error Resume Next

If Form1.Width = 11520 Or Form1.Width = 11430 Or Form1.Width = 11448 Or Form1.Width = fowi Then
   Form1.Width = 14280
Else
   Form1.Width = 11430
End If

End Sub

Private Function save(filename As String, path As String) As String

On Error Resume Next

Dim filebox As OPENFILENAME
Dim fname As String
Dim result As Long

With filebox
    .lStructSize = Len(filebox)
    .hwndOwner = Me.hwnd
    .hInstance = 0
    .lpstrFilter = "Android apk (*.apk)" & vbNullChar & "*.apk" & vbNullChar
    .nMaxCustomFilter = 0
    .nFilterIndex = 1
    .lpstrFileTitle = "454544" & vbNullChar
    .lpstrFile = filename & Space(257 - Len(filename)) & vbNullChar
    .nMaxFile = Len(.lpstrFile)
    .lpstrFileTitle = Space(256) & vbNullChar
    .nMaxFileTitle = Len(.lpstrFileTitle)
    .lpstrInitialDir = path & vbNullChar
    .lpstrTitle = "Save APK as" & vbNullChar
    .flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
    .nFileOffset = 0
    .nFileExtension = 0
    .lCustData = 0
    .lpfnHook = 0
End With
result = GetSaveFileName(filebox)
If result <> 0 Then
    fname = Left(filebox.lpstrFile, InStr(filebox.lpstrFile, vbNullChar) - 1)
    save = fname
Else
    save = ""
End If

End Function

Private Function save2(filename As String, path As String) As String

On Error Resume Next

Dim filebox As OPENFILENAME
Dim fname As String
Dim result As Long

With filebox
    .lStructSize = Len(filebox)
    .hwndOwner = Me.hwnd
    .hInstance = 0
    .lpstrFilter = "Zip File (*.zip)" & vbNullChar & "*.zip" & vbNullChar
    .nMaxCustomFilter = 0
    .nFilterIndex = 1
    .lpstrFileTitle = "454544" & vbNullChar
    .lpstrFile = filename & Space(257 - Len(filename)) & vbNullChar
    .nMaxFile = Len(.lpstrFile)
    .lpstrFileTitle = Space(256) & vbNullChar
    .nMaxFileTitle = Len(.lpstrFileTitle)
    .lpstrInitialDir = path & vbNullChar
    .lpstrTitle = "Save release Zip as" & vbNullChar
    .flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
    .nFileOffset = 0
    .nFileExtension = 0
    .lCustData = 0
    .lpfnHook = 0
End With
result = GetSaveFileName(filebox)
If result <> 0 Then
    fname = Left(filebox.lpstrFile, InStr(filebox.lpstrFile, vbNullChar) - 1)
    save2 = fname
Else
    save2 = ""
End If

End Function

Private Function save3(filename As String, path As String) As String

On Error Resume Next

Dim filebox As OPENFILENAME
Dim fname As String
Dim result As Long

With filebox
    .lStructSize = Len(filebox)
    .hwndOwner = Me.hwnd
    .hInstance = 0
    .lpstrFilter = "mp4 File (*.mp4)" & vbNullChar & "*.mp4" & vbNullChar
    .nMaxCustomFilter = 0
    .nFilterIndex = 1
    .lpstrFileTitle = "454544" & vbNullChar
    .lpstrFile = filename & Space(257 - Len(filename)) & vbNullChar
    .nMaxFile = Len(.lpstrFile)
    .lpstrFileTitle = Space(256) & vbNullChar
    .nMaxFileTitle = Len(.lpstrFileTitle)
    .lpstrInitialDir = path & vbNullChar
    .lpstrTitle = "Save Video as" & vbNullChar
    .flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
    .nFileOffset = 0
    .nFileExtension = 0
    .lCustData = 0
    .lpfnHook = 0
End With
result = GetSaveFileName(filebox)
If result <> 0 Then
    fname = Left(filebox.lpstrFile, InStr(filebox.lpstrFile, vbNullChar) - 1)
    save3 = fname
Else
    save3 = ""
End If

End Function

Private Function create(fu As Integer, fu2 As Integer) As String

On Error Resume Next

Dim ap As String
Dim ap1 As String
Dim t As String
Dim MyPath As String
Dim idr As String
Dim an As String

create = "0"
' apk prüfung build ordner!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
J = Chr$(34)
If Dir$(BuildPath & "\*.*") = vbNullString Then
   MessageBeep (16)
   Message "Put your .gltf model-files in dir .\Build"
   Exit Function
End If
If Dir$(BuildPath & "\*.glb") <> "" Then
   MessageBeep (16)
   Message "GLB file found in .\Build, Error!" & vbNewLine & "Choose glTF separate in Blender!"
   Exit Function
End If
If Dir$(BuildPath & "\*.apk") <> "" Then
   MessageBeep (16)
   Message "APK file found in .\Build, Error!"
   Exit Function
End If
If Dir$(BuildPath & "\*.zip") <> "" Then
   MessageBeep (16)
   Message "ZIP file found in .\Build, Error!"
   Exit Function
End If
If Dir(apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene") <> "" Then Kill apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene"
If Dir(apppath & "\files\tmp\scene.zip") <> "" Then Kill apppath & "\files\tmp\scene.zip"
If Dir(apppath & "\files\tmp\_BACKGROUND_LOOP.ogg") <> "" Then Kill apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"

If Check10.Value = True And comp1 = False Then
   CompressKTX
   If check_build = False Then Message ".JPG/.PNG and .KTX present in same folder!"
   comp1 = True
End If

'If Check2.Value = "1" Then
'   FileCopy apppath & "\files\default.ogg", apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"
'   GoTo tell
'End If
If fu = "1" Then
   FileCopy apppath & "\files\silent.ogg", apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"
   GoTo tell
End If
If LCase(Right$(Label9.Caption, 3)) = "ogg" Then
   If Check4.Value = False And Check1.Value = False Then
      FileCopy aud, apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"
      GoTo tell
   End If
End If
If fu = 0 And Check2.Value = True Then
   FileCopy apppath & "\files\default.ogg", apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"
   GoTo tell
End If
If Label9.Caption = "" And Check2.Value = False And Check3.Value = False Then GoTo killer
If Check1.Value = True Then
   txtOutputs.Text = txtOutputs.Text & vbNewLine & vbNewLine & "Encode Audio File..." & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
   Pause (0.2)
   objDOS.CommandLine = (apppath & "\files\sox.exe -S " & J & aud & J & " -C 3 " & J & apppath & "\files\tmp\_BACKGROUND_LOOP.ogg" & J & " vol -" & lvButtons_H.Caption & " dB")
   objDOS.ExecuteCommand
Else
   txtOutputs.Text = txtOutputs.Text & vbNewLine & vbNewLine & "Encode Audio File..." & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
   Pause (0.2)
   objDOS.CommandLine = (apppath & "\files\sox.exe -S " & J & aud & J & " -C 3 " & J & apppath & "\files\tmp\_BACKGROUND_LOOP.ogg" & J)
   objDOS.ExecuteCommand
End If

killer:

If Dir(apppath & "\files\tmpz.apk") <> "" Then Kill apppath & "\files\tmpz.apk"
If Dir(apppath & "\files\tmp.apk") <> "" Then Kill apppath & "\files\tmp.apk"
If Dir(apppath & "\files\tmp.zip") <> "" Then Kill apppath & "\files\tmp.apk"
If Dir(apppath & "\files\scene.zip") <> "" Then Kill apppath & "\files\scene.zip"
If Dir(apppath & "\files\tmp\temp_ec.wav") <> "" Then Kill apppath & "\files\tmp\temp_ec.wav"

tell:
'--------------------------------------------------------------
objDOS.CommandLine = (apppath & "\files\7za.exe a " & J & apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene.zip" & J & " " & J & BuildPath & "\*" & J)
objDOS.ExecuteCommand
Name apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene.zip" As apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene"
If Dir(apppath & "\files\tmp\_BACKGROUND_LOOP.ogg") = "" Then FileCopy apppath & "\files\silent.ogg", apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"
objDOS.CommandLine = (apppath & "\files\7za.exe a " & J & apppath & "\files\tmp\scene.zip" & J & " " & J & apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene" & J & " " & J & apppath & "\files\tmp\_BACKGROUND_LOOP.ogg" & J)
objDOS.ExecuteCommand
If Dir(apppath & "\files\tmp\_BACKGROUND_LOOP.ogg") <> "" Then Kill apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"
Kill apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene"
If fu2 = "1" Then
   ap = apppath & "\files\WinterLodge\assets\"
   ap1 = apppath & "\files\WinterLodge"
   an = "WinterLodge"
End If
If fu2 = "0" Then
   ap = apppath & "\files\ClassicHome\assets\"
   ap1 = apppath & "\files\ClassicHome"
   an = "ClassicHome"
End If
If fu2 = "2" Then
   ap = apppath & "\files\SpaceStation\assets\"
   ap1 = apppath & "\files\SpaceStation"
   an = "SpaceStation"
End If
If fu2 = "3" Then
   ap = apppath & "\files\CyberCity\assets\"
   ap1 = apppath & "\files\CyberCity"
   an = "CyberCity"
End If

GoTo fastbuild

'----------------------------------------------------------------------------
FileCopy apppath & "\files\tmp\scene.zip", ap & "scene.zip"
objDOS.CommandLine = (java & " -Xmx1024m -jar " & J & apppath & "\files\apktool_2.3.4.jar" & J & " b -f -o " & J & apppath & "\files\tmp\tmp.apk" & J & " " & J & ap1 & J)
objDOS.ExecuteCommand
txtOutputs.Text = txtOutputs.Text & vbNewLine & "Zipalign APK-file" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
objDOS.CommandLine = (apppath & "\files\zipalign.exe -f 4 " & J & apppath & "\files\tmp\tmp.apk" & J & " " & J & apppath & "\files\tmp\tmpz.apk" & J)
objDOS.ExecuteCommand
txtOutputs.Text = txtOutputs.Text & vbNewLine & "Sign APK-file" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
objDOS.CommandLine = (java & " -Xmx1024m -jar " & J & apppath & "\files\ApkSigner.jar" & J & " sign  --key " & J & apppath & "\files\apkeasytool.pk8" & J & " --cert " & J & _
     apppath & "\files\apkeasytool.pem" & J & " --out " & J & apppath & "\files\tmp\tmpz.apk" & J & " " & J & apppath & "\files\tmp\tmpz.apk" & J)
objDOS.ExecuteCommand

MyPath = Dir(BuildPath & "\")
Do Until MyPath = vbNullString
        If Mid(MyPath, Len(MyPath) - 4) = ".gltf" Then
            idr = Left$(MyPath, Len(MyPath) - 5)
        End If
    MyPath = Dir
Loop

FileCopy apppath & "\files\tmp\tmpz.apk", apppath & "\" & idr & "." & an & ".apk"
GoTo insta
'-----------------------------------------------------

fastbuild:

ChDir apppath
ChDrive Left$(apppath, 2)
FileCopy apppath & "\files\" & an & ".zip", apppath & "\files\tmp.zip"
FileCopy apppath & "\files\tmp\scene.zip", apppath & "\files\scene.zip"
objDOS.CommandLine = (apppath & "\files\7za.exe a " & J & apppath & "\files\tmp.zip" & J & " files\scene.zip")
objDOS.ExecuteCommand
objDOS.CommandLine = (apppath & "\files\7za.exe rn " & J & apppath & "\files\tmp.zip" & J & " files\ assets\")
objDOS.ExecuteCommand
Name apppath & "\files\tmp.zip" As apppath & "\files\tmp.apk"
txtOutputs.Text = txtOutputs.Text & vbNewLine & "Zipalign APK-file" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
objDOS.CommandLine = (apppath & "\files\zipalign.exe -f 4 " & J & apppath & "\files\tmp.apk" & J & " " & J & apppath & "\files\tmpz.apk" & J)
objDOS.ExecuteCommand
txtOutputs.Text = txtOutputs.Text & vbNewLine & "Sign APK-file" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
objDOS.CommandLine = (java & " -Xmx1024m -jar " & J & apppath & "\files\ApkSigner.jar" & J & " sign  --key " & J & apppath & "\files\apkeasytool.pk8" & J & " --cert " & J & _
     apppath & "\files\apkeasytool.pem" & J & " --out " & J & apppath & "\files\tmpz.apk" & J & " " & J & apppath & "\files\tmpz.apk" & J)
objDOS.ExecuteCommand
MyPath = Dir(BuildPath & "\")
Do Until MyPath = vbNullString
        If Mid(MyPath, Len(MyPath) - 4) = ".gltf" Then
            idr = Left$(MyPath, Len(MyPath) - 5)
        End If
    MyPath = Dir
Loop

If fu = 1 Then an = an & "_Silent"
FileCopy apppath & "\files\tmpz.apk", apppath & "\" & idr2 & "." & an & ".apk"
fin2 = apppath & "\" & idr2 & "." & an & ".apk"
insta:

'objDOS.CommandLine = (apppath & "\files\adb.exe kill-server")
'objDOS.ExecuteCommand

nex:

If Dir(apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene") <> "" Then Kill apppath & "\files\tmp\_WORLD_MODEL.gltf.ovrscene"
If Dir(apppath & "\files\tmp\scene.zip") <> "" Then Kill apppath & "\files\tmp\scene.zip"
If Dir(apppath & "\files\scene.zip") <> "" Then Kill apppath & "\files\scene.zip"
If Dir(apppath & "\files\tmp\_BACKGROUND_LOOP.ogg") <> "" Then Kill apppath & "\files\tmp\_BACKGROUND_LOOP.ogg"
If Dir(apppath & "\files\tmp\tmpz.apk") <> "" Then Kill apppath & "\files\tmp\tmpz.apk"
If Dir(apppath & "\files\tmp\tmp.apk") <> "" Then Kill apppath & "\files\tmp\tmp.apk"
If Dir(apppath & "\files\tmpz.apk") <> "" Then Kill apppath & "\files\tmpz.apk"
If Dir(apppath & "\files\tmp.apk") <> "" Then Kill apppath & "\files\tmp.apk"
If Dir(apppath & "\files\tmp.zip") <> "" Then Kill apppath & "\files\tmp.apk"

txtOutputs.Text = txtOutputs.Text & vbNewLine & "Build APK " & fu & "finished! " & Time & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
create = "1"

End Function

Private Sub Pause(Seconds As Single)


On Error Resume Next

Dim Timerx1 As Single, Timerx2 As Single, currentDate As Date

currentDate = Date
Timerx1 = Timer + Seconds
Timerx2 = Timerx1 - 86400
While ((Timer() < Timerx1) And (currentDate = Date)) Or ((Timer() < Timerx2) And (currentDate + 1 = Date))
  DoEvents
Wend

End Sub

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

Function ShowColorDialog(Optional ByVal hParent As Long, Optional ByVal bFullOpen As Boolean, Optional ByVal InitColor As OLE_COLOR) As Long

Dim CC As ChooseColorStruct
Dim aColorRef(15) As Long
Dim lInitColor As Long

On Error Resume Next

If InitColor <> 0 Then
   If OleTranslateColor(InitColor, 0, lInitColor) Then
      lInitColor = &HFFFF
   End If
End If
With CC
    .lStructSize = Len(CC)
    .hwndOwner = hParent
    .lpCustColors = VarPtr(aColorRef(0))
    .rgbResult = lInitColor
    .flags = &H80& Or &H100& Or &H1& Or IIf(bFullOpen, &H2&, 0)
End With
If ChooseColor(CC) Then
   ShowColorDialog = CC.rgbResult
Else
   ShowColorDialog = -1
End If
    
End Function

Private Sub Form_Unload(Cancel As Integer)

'To set default colors...... (Don't Change or remove this)
  s = SetSysColors(1, COLOR_CAPTIONTEXT, vbWhite)
'Removing this will make all app's Titlebar text to be that color!!

End Sub

Private Function BrowseForFolder(ByVal lngHwnd As Long, ByVal strPrompt As String) As String

On Error GoTo ehBrowseForFolder

Dim intNull As Integer
Dim lngIDList As Long, lngResult As Long
Dim strPath As String
Dim udtBI As BrowseInfo

With udtBI
    .lngHwnd = lngHwnd
    .lpszTitle = lstrcat(strPrompt, "")
    .ulFlags = BIF_RETURNONLYFSDIRS
End With
lngIDList = SHBrowseForFolder(udtBI)
If lngIDList <> 0 Then
   strPath = String(MAX_PATH, 0)
   lngResult = SHGetPathFromIDList(lngIDList, strPath)
   Call CoTaskMemFree(lngIDList)
   intNull = InStr(strPath, vbNullChar)
   If intNull > 0 Then
      strPath = Left(strPath, intNull - 1)
   End If
End If
BrowseForFolder = strPath
Exit Function

ehBrowseForFolder:
BrowseForFolder = Empty

End Function

Private Function GetCheck() As String

On Error Resume Next

GetCheck = Abs(CInt(Check0.Value)) & Abs(CInt(Check6.Value)) & Abs(CInt(Check7.Value)) & Abs(CInt(Check8.Value)) & Abs(CInt(Check9.Value)) & Abs(CInt(Check10.Value)) & Abs(CInt(Check19.Value))

End Function

Private Function Rename(ap9 As String) As String

Dim tu As String

On Error Resume Next

tu = Mid$(Label8.Caption, 1, Len(Label8.Caption) - 4) & "_new.apk"
If ap9 = apppath & "\files\WinterLodge" Then
   If InStr(1, LCase(tu), "classic", 0) <> 0 Then
      tu = Replace(tu, "classichome", "WinterLodge", , , vbTextCompare)
      tu = Replace(tu, "classic home", "WinterLodge", , , vbTextCompare)
      tu = Replace(tu, "classic_home", "WinterLodge", , , vbTextCompare)
      tu = Replace(tu, "classic.home", "WinterLodge", , , vbTextCompare)
      GoTo ren_end
   End If
   If InStr(1, LCase(tu), "space", 0) <> 0 Then
      tu = Replace(tu, "spacestation", "WinterLodge", , , vbTextCompare)
      tu = Replace(tu, "space station", "WinterLodge", , , vbTextCompare)
      tu = Replace(tu, "space_station", "WinterLodge", , , vbTextCompare)
      tu = Replace(tu, "space.station", "WinterLodge", , , vbTextCompare)
      GoTo ren_end
   End If
End If
If ap9 = apppath & "\files\ClassicHome" Then
   If InStr(1, LCase(tu), "winter", 0) <> 0 Then
      tu = Replace(tu, "WinterLodge", "ClassicHome", , , vbTextCompare)
      tu = Replace(tu, "Winter Lodge", "ClassicHome", , , vbTextCompare)
      tu = Replace(tu, "Winter_Lodge", "ClassicHome", , , vbTextCompare)
      tu = Replace(tu, "Winter.Lodge", "ClassicHome", , , vbTextCompare)
      GoTo ren_end
   End If
   If InStr(1, LCase(tu), "space", 0) <> 0 Then
      tu = Replace(tu, "spacestation", "ClassicHome", , , vbTextCompare)
      tu = Replace(tu, "space station", "ClassicHome", , , vbTextCompare)
      tu = Replace(tu, "space_station", "ClassicHome", , , vbTextCompare)
      tu = Replace(tu, "space.station", "ClassicHome", , , vbTextCompare)
      GoTo ren_end
   End If
End If
If ap9 = apppath & "\files\SpaceStation" Then
   If InStr(1, LCase(tu), "winter", 0) <> 0 Then
      tu = Replace(tu, "WinterLodge", "SpaceStation", , , vbTextCompare)
      tu = Replace(tu, "Winter Lodge", "SpaceStation", , , vbTextCompare)
      tu = Replace(tu, "Winter_Lodge", "SpaceStation", , , vbTextCompare)
      tu = Replace(tu, "Winter.Lodge", "SpaceStation", , , vbTextCompare)
      GoTo ren_end
   End If
   If InStr(1, LCase(tu), "classic", 0) <> 0 Then
      tu = Replace(tu, "classichome", "SpaceStation", , , vbTextCompare)
      tu = Replace(tu, "classic home", "SpaceStation", , , vbTextCompare)
      tu = Replace(tu, "classic_home", "SpaceStation", , , vbTextCompare)
      tu = Replace(tu, "classic.home", "SpaceStation", , , vbTextCompare)
      GoTo ren_end
   End If
End If

ren_end:

Rename = tu

End Function

Private Function CountFiles(StrFileName As String) As Long

On Error Resume Next

StrFileName = Dir$(StrFileName)
Do While Len(StrFileName) <> 0
    CountFiles = CountFiles + 1
    StrFileName = Dir$
Loop

End Function

Private Sub Check13_Click()

On Error Resume Next

If Check13.Value = False Then
   RmDir apppath & "\files\texture_tmp"
   If Dir(apppath & "\files\texture_tmp" & "\*.*") <> "" Then Kill apppath & "\files\texture_tmp" & "\*.*"
   Label11.Caption = "ERASED!": Beep
   txtOutputs.Text = txtOutputs.Text & vbNewLine & vbNewLine & "Textures in .\Build deleted!" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
   tp = True
   Exit Sub
End If
If Dir(apppath & "\files\texture_tmp" & "\*.*") <> "" Then Kill apppath & "\files\texture_tmp" & "\*.*"

Set fsx = CreateObject("Scripting.FileSystemObject")
If fsx.FolderExists(apppath & "\files\texture_tmp") = False Then MkDir (apppath & "\files\texture_tmp")
If Dir$(apppath & "\Build" & "\*.*") <> vbNullString Then
    For Each oFile In fsx.GetFolder(apppath & "\Build" & "").Files
        If LCase(fsx.GetExtensionName(oFile.path)) <> "bin" And LCase(fsx.GetExtensionName(oFile.path)) <> "gltf" Then
            fn = fsx.GetFileName(oFile.path)
            FileCopy BuildPath & "\" & fn, apppath & "\files\texture_tmp\" & fn
        End If
    Next
    Label11.Caption = "SAVED!": Beep
    txtOutputs.Text = txtOutputs.Text & vbNewLine & vbNewLine & "Textures in .\Build saved!" & vbNewLine: txtOutputs.SelStart = Len(txtOutputs.Text)
    tp = True
End If

End Sub

Private Sub Check15_Click()

On Error Resume Next

If Check15.Value = True Then
   PutINISetting "Save", "AutoClear", "1", apppath & "\files\config.ini"
Else
   PutINISetting "Save", "AutoClear", "0", apppath & "\files\config.ini"
End If
Pause (0.5)
Command2_Click

End Sub

Private Sub Command14_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error Resume Next

Dim r As Long

'Call Form5.Show
If Question("Yes = Open tutorial in Browser" & vbNewLine & "No = Open with Adobe Reader", True) = True Then
   r = ShellExecute(0, "open", "https://documentcloud.adobe.com/link/track?uri=urn:aaid:scds:US:378deebf-9e73-4100-bdb1-40b816baef58", 0, 0, 1)
Else
   r = ShellExecute(0, "open", apppath & "\EnviromentConverterBuilder_HowTo.pdf", 0, 0, 1)
End If

End Sub

Private Sub Check17_Click()

If Check17.Value = True Then
   PutINISetting "Save", "Pack", "1", apppath & "\files\config.ini"
Else
   PutINISetting "Save", "Pack", "0", apppath & "\files\config.ini"
End If
Pause (0.5)
Command2_Click

End Sub

Private Sub lvButtons_H11_Click()

If lvButtons_H11.Value = True Then
   PutINISetting "Save", "AudioReset", "1", apppath & "\files\config.ini"
Else
   PutINISetting "Save", "AudioReset", "0", apppath & "\files\config.ini"
End If
Pause (0.5)
Command2_Click

End Sub

