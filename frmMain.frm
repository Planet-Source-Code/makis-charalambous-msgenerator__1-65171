VERSION 5.00
Begin VB.Form hscrolPercent 
   BackColor       =   &H00545249&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rife Generator 1.0"
   ClientHeight    =   5790
   ClientLeft      =   2415
   ClientTop       =   4560
   ClientWidth     =   12120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   386
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   808
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmbEndSweep 
      Caption         =   "End Sweep"
      Height          =   945
      Left            =   4290
      Picture         =   "frmMain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   4320
      Width           =   1455
   End
   Begin MsGen.msDial ctDial1 
      Height          =   555
      Left            =   720
      TabIndex        =   72
      Top             =   870
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   979
   End
   Begin VB.CommandButton bttnProg 
      Caption         =   "Exit"
      Height          =   645
      Index           =   5
      Left            =   10710
      Picture         =   "frmMain.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   5100
      Width           =   1095
   End
   Begin VB.CommandButton cmbSweep 
      Caption         =   "Start Sweep"
      Height          =   945
      Left            =   2820
      Picture         =   "frmMain.frx":0784
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00545249&
      Caption         =   "Shape"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1605
      Left            =   300
      TabIndex        =   47
      Top             =   2640
      Width           =   2205
      Begin MsGen.msSwitch bttnShape 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         value           =   -1  'True
         Caption         =   "Sine"
      End
      Begin MsGen.msSwitch bttnShape 
         Height          =   360
         Index           =   1
         Left            =   120
         TabIndex        =   64
         Top             =   540
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Square"
      End
      Begin MsGen.msSwitch bttnShape 
         Height          =   360
         Index           =   2
         Left            =   120
         TabIndex        =   65
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Sawtooth"
      End
      Begin MsGen.msSwitch bttnShape 
         Height          =   360
         Index           =   3
         Left            =   120
         TabIndex        =   66
         Top             =   1140
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ramp"
      End
   End
   Begin VB.Timer TmrProgram 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7710
      Top             =   4350
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00545249&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1605
      Left            =   9870
      TabIndex        =   34
      Top             =   2640
      Width           =   1935
      Begin MsGen.msDial HScroll2 
         Height          =   555
         Left            =   840
         TabIndex        =   76
         Top             =   330
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   979
      End
      Begin MsGen.msDial hscrChop 
         Height          =   555
         Left            =   120
         TabIndex        =   75
         Top             =   330
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   979
      End
      Begin MsGen.LED chkChop 
         Height          =   165
         Left            =   150
         TabIndex        =   35
         Top             =   60
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   291
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Pulses Second"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   40
         Top             =   1260
         Width           =   555
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Mark Space"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   345
         Index           =   4
         Left            =   210
         TabIndex        =   39
         Top             =   1260
         Width           =   465
      End
      Begin VB.Label lblChopFrequency 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   840
         TabIndex        =   38
         Top             =   1020
         Width           =   585
      End
      Begin VB.Label lblChop 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   90
         TabIndex        =   37
         Top             =   1020
         Width           =   585
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00545249&
         Caption         =   " Chop "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   6
         Left            =   465
         TabIndex        =   36
         Top             =   0
         Width           =   555
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00545249&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1605
      Left            =   7680
      TabIndex        =   25
      Top             =   2640
      Width           =   2175
      Begin MsGen.msDial hscrolPercent 
         Height          =   555
         Left            =   120
         TabIndex        =   74
         Top             =   300
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   979
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   1830
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1320
         Width           =   285
      End
      Begin VB.VScrollBar HScroll1 
         Height          =   1125
         LargeChange     =   10
         Left            =   1830
         Max             =   5000
         Min             =   100
         TabIndex        =   42
         Top             =   210
         Value           =   100
         Width           =   285
      End
      Begin MsGen.LED LED1 
         Height          =   165
         Left            =   120
         TabIndex        =   28
         Top             =   60
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   291
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00545249&
         Caption         =   " Wobbulation "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   3
         Left            =   390
         TabIndex        =   29
         Top             =   0
         Width           =   1185
      End
      Begin VB.Label lblFreq 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   870
         TabIndex        =   33
         Top             =   1260
         Width           =   795
      End
      Begin VB.Label lblFreq 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   32
         Top             =   660
         Width           =   795
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Low"
         ForeColor       =   &H0080FFFF&
         Height          =   225
         Index           =   1
         Left            =   870
         TabIndex        =   31
         Top             =   390
         Width           =   525
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "High"
         ForeColor       =   &H0080FFFF&
         Height          =   225
         Index           =   2
         Left            =   870
         TabIndex        =   30
         Top             =   990
         Width           =   525
      End
      Begin VB.Label lblWob 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   90
         TabIndex        =   27
         Top             =   960
         Width           =   585
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00545249&
      Caption         =   "Gain"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1605
      Left            =   5550
      TabIndex        =   24
      Top             =   2640
      Width           =   2115
      Begin MsGen.msSwitch bttnDB 
         Height          =   360
         Index           =   0
         Left            =   810
         TabIndex        =   51
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         value           =   -1  'True
      End
      Begin MsGen.msSwitch bttnDB 
         Height          =   360
         Index           =   1
         Left            =   810
         TabIndex        =   52
         Top             =   540
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "-10 db"
      End
      Begin MsGen.msSwitch bttnDB 
         Height          =   360
         Index           =   2
         Left            =   810
         TabIndex        =   53
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "-20 db"
      End
      Begin MsGen.msSwitch bttnDB 
         Height          =   360
         Index           =   3
         Left            =   810
         TabIndex        =   54
         Top             =   1140
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "-30db"
      End
      Begin MsGen.msDial ctDial2 
         Height          =   555
         Left            =   150
         TabIndex        =   73
         Top             =   330
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   979
      End
      Begin VB.Label lblVol 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   90
         TabIndex        =   26
         Top             =   1020
         Width           =   585
      End
   End
   Begin VB.Timer TimerWob 
      Enabled         =   0   'False
      Left            =   6480
      Top             =   4350
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00545249&
      BorderStyle     =   0  'None
      Caption         =   "Digital Selector"
      Height          =   765
      Left            =   330
      TabIndex        =   11
      Top             =   1860
      Width           =   1605
      Begin VB.VScrollBar VScroll4 
         Height          =   30
         Left            =   390
         TabIndex        =   19
         Top             =   60
         Width           =   255
      End
      Begin VB.TextBox txtDecimal 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   180
         Width           =   250
      End
      Begin VB.TextBox txtUnits 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   180
         Width           =   250
      End
      Begin VB.TextBox txtTens 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   570
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   180
         Width           =   250
      End
      Begin VB.TextBox txtThousands 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   180
         Width           =   250
      End
      Begin VB.TextBox txtHundreds 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   180
         Width           =   250
      End
      Begin VB.VScrollBar scrHundreds 
         Height          =   585
         Left            =   300
         Max             =   0
         Min             =   9
         TabIndex        =   13
         Top             =   60
         Width           =   285
      End
      Begin VB.VScrollBar scrThousands 
         Height          =   585
         Left            =   30
         Max             =   0
         Min             =   9
         TabIndex        =   18
         Top             =   60
         Width           =   285
      End
      Begin VB.VScrollBar scrTens 
         Height          =   585
         Left            =   570
         Max             =   0
         Min             =   9
         TabIndex        =   20
         Top             =   60
         Width           =   285
      End
      Begin VB.VScrollBar scrUnits 
         Height          =   585
         Left            =   840
         Max             =   0
         Min             =   9
         TabIndex        =   21
         Top             =   60
         Width           =   285
      End
      Begin VB.VScrollBar scrDecimal 
         Height          =   585
         Left            =   1200
         Max             =   0
         Min             =   9
         TabIndex        =   22
         Top             =   60
         Width           =   300
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1140
         TabIndex        =   23
         Top             =   270
         Width           =   45
      End
   End
   Begin MsGen.LCDDisplay LCDDisplay1 
      Height          =   795
      Left            =   2490
      TabIndex        =   10
      Top             =   720
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   1402
      DCount          =   5
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00545249&
      Caption         =   "Buffer"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1605
      Left            =   4020
      TabIndex        =   9
      Top             =   2640
      Width           =   1515
      Begin MsGen.msSwitch bttnBuffer 
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   55
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "4096"
      End
      Begin MsGen.msSwitch bttnBuffer 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   56
         Top             =   540
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "8192"
      End
      Begin MsGen.msSwitch bttnBuffer 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   57
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         value           =   -1  'True
         Caption         =   "16384"
      End
      Begin MsGen.msSwitch bttnBuffer 
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   58
         Top             =   1140
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "32768"
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00545249&
      Caption         =   "Sample Rate"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1605
      Left            =   2520
      TabIndex        =   8
      Top             =   2640
      Width           =   1515
      Begin MsGen.msSwitch bttnSamples 
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   59
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "11025"
      End
      Begin MsGen.msSwitch bttnSamples 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   60
         Top             =   540
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "22050"
      End
      Begin MsGen.msSwitch bttnSamples 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   61
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "24000"
      End
      Begin MsGen.msSwitch bttnSamples 
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   62
         Top             =   1140
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         value           =   -1  'True
         Caption         =   "48000"
      End
   End
   Begin VB.Timer tmrSweep 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7320
      Top             =   4350
   End
   Begin VB.TextBox txtSweep 
      Height          =   285
      Index           =   2
      Left            =   1470
      TabIndex        =   7
      Text            =   "100"
      Top             =   4950
      Width           =   1305
   End
   Begin VB.TextBox txtSweep 
      Height          =   285
      Index           =   1
      Left            =   1470
      TabIndex        =   6
      Text            =   "1000"
      Top             =   4620
      Width           =   1305
   End
   Begin VB.TextBox txtSweep 
      Height          =   285
      Index           =   0
      Left            =   1470
      TabIndex        =   5
      Text            =   "100"
      Top             =   4320
      Width           =   1305
   End
   Begin VB.Timer tmrChop 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   6060
      Top             =   4350
   End
   Begin VB.Frame fraMod 
      BackColor       =   &H00545249&
      Caption         =   "Amplitude Modulation"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   9870
      TabIndex        =   0
      Top             =   4260
      Width           =   1935
      Begin VB.HScrollBar hscMod 
         Height          =   255
         LargeChange     =   10
         Left            =   450
         Max             =   100
         TabIndex        =   1
         Top             =   270
         Width           =   1395
      End
      Begin VB.Label lblAmpMod 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   60
         TabIndex        =   69
         Top             =   300
         Width           =   435
      End
   End
   Begin VB.Frame fraOsc 
      Caption         =   "Oscilloscope"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1875
      Left            =   5550
      TabIndex        =   2
      Top             =   690
      Width           =   6285
      Begin VB.PictureBox picOsc 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00004000&
         Enabled         =   0   'False
         Height          =   1215
         Left            =   120
         ScaleHeight     =   77
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   397
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   6015
      End
      Begin VB.Line lin2 
         BorderColor     =   &H000000FF&
         X1              =   6090
         X2              =   6090
         Y1              =   1440
         Y2              =   1800
      End
      Begin VB.Line lin1 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   120
         Y1              =   1440
         Y2              =   1800
      End
      Begin VB.Line linGraph 
         BorderColor     =   &H00FF0000&
         X1              =   120
         X2              =   6090
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label lblGraph 
         Alignment       =   2  'Center
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   5895
      End
   End
   Begin VB.Timer tmrMod 
      Left            =   6900
      Top             =   4350
   End
   Begin MsGen.LCDDisplay LCDDisplay2 
      Height          =   795
      Left            =   2490
      TabIndex        =   44
      Top             =   1770
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   1402
      DCount          =   5
   End
   Begin MsGen.msSwitch bttnOnOff 
      Height          =   360
      Index           =   0
      Left            =   780
      TabIndex        =   67
      Top             =   5400
      Width           =   1035
      _ExtentX        =   2037
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      value           =   -1  'True
      Caption         =   "OFF"
   End
   Begin MsGen.msSwitch bttnOnOff 
      Height          =   360
      Index           =   1
      Left            =   1920
      TabIndex        =   68
      Top             =   5400
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ON"
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Duration Sec."
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   2
      Left            =   330
      TabIndex        =   50
      Top             =   4980
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "End Freq."
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   1
      Left            =   330
      TabIndex        =   49
      Top             =   4650
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Freq."
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   0
      Left            =   330
      TabIndex        =   48
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "ms"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   435
      Index           =   1
      Left            =   2010
      TabIndex        =   45
      Top             =   2190
      Width           =   435
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hz"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   435
      Index           =   0
      Left            =   2010
      TabIndex        =   43
      Top             =   1080
      Width           =   405
   End
   Begin VB.Label lblhead 
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed by Makis Charalambous"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   540
      Left            =   660
      TabIndex        =   41
      Top             =   90
      Width           =   1875
   End
   Begin VB.Image Image4 
      Height          =   645
      Left            =   240
      Picture         =   "frmMain.frx":0A03
      Top             =   0
      Width           =   11640
   End
   Begin VB.Image Image3 
      Height          =   8775
      Index           =   1
      Left            =   11880
      Picture         =   "frmMain.frx":336B
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   8775
      Index           =   0
      Left            =   0
      Picture         =   "frmMain.frx":4501
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   1
      Left            =   330
      Picture         =   "frmMain.frx":5697
      Top             =   5430
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   285
      Index           =   0
      Left            =   330
      Picture         =   "frmMain.frx":5CED
      Top             =   5430
      Visible         =   0   'False
      Width           =   390
   End
End
Attribute VB_Name = "hscrolPercent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sep As String
Dim iProgramMode As Integer
Dim lSeconds As Long

Dim iShape As Integer
Dim iChopFrequency As Integer
Dim wobValue As Integer
Dim wobpercent As Single
Dim iWobulator As Long
Dim wobProgress As Long
Dim iWobulatorMax As Long
Dim iWobulatorMin As Long
Dim iWobulatorStep As Integer

Dim iPrevVolume As Integer
Dim iChopMarkSpace As Integer
Dim iCounter As Integer
Dim DontChangeFlag As Boolean
Dim tmp As Single

Dim bSweepInProgress As Boolean
Dim iSweep0 As Single
Dim iSweep1 As Single
Dim iSweepSteps As Single
Dim iSweepStep As Single
Dim iSweepValue As Long
Dim iSweepMaxValue As Long
Dim iSweeNormValue As Long
Dim iSign As Integer
Dim MinDSBFrequency As Long

Dim nVolume As Single
Const MaxVolume = 0
Const MinVolume = -10000
Const MaxFrequency = 100000
Const MinFrequency = 100

Dim nSamples As Long
Dim nBasicBufferSize As Long
Const pi = 3.14159265358979
Dim c As Double
Dim DX7         As New DirectX7
Dim DS          As DirectSound
Dim DSB         As DirectSoundBuffer
Dim PCM         As WAVEFORMATEX
Dim DSBD        As DSBUFFERDESC
Dim i As Long

Dim nFreq As Single, nMod!, nModDir%

Private Sub SinBuffer(ByVal nFrequency As Single, ByVal nVolume!)
    
    Dim lpBuffer() As Byte, c#, nBuffer&
    Dim sValue As Double

    If nFrequency <= 0 Then Exit Sub
    
    SetDigitalSelector
    
    LCDDisplay1.Value = FormatNumber(nFreq, 1, , , vbFalse)
    LCDDisplay2.Value = FormatNumber(1000 / nFrequency, 2, , , vbFalse)
        
    lblVol = FormatPercent(nVolume, 0)
    
    c = nSamples / nFrequency
    
    nBuffer = (nBasicBufferSize \ c) * c
    If nBuffer = 0 Then nBuffer = c

    ReDim lpBuffer(nBuffer - 1)
    
    sValue = (2 * pi * nFrequency / nSamples)
    
    For i = 0 To nBuffer - 1
        c = Sin(i * sValue)
        If iShape = 1 Then
            c = Sgn(c)
            If c = 0 Then c = 1
        End If
        lpBuffer(i) = (c * nMod * nVolume + 1) * 127.5!
    ' lpBuffer(i) = (127.5! * Sin(2 * pi * (i * nSamples/ nFrequency ))) + 127.5!
    Next
    
    If DSBD.lBufferBytes <> nBuffer Then
        DSBD.lBufferBytes = nBuffer
        DSBD.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
        Set DSB = DS.CreateSoundBuffer(DSBD, PCM)
    End If

    DSB.WriteBuffer 0, 0, lpBuffer(0), DSBLOCK_ENTIREBUFFER
    If bttnOnOff(1).Value Then
       DSB.Play DSBPLAY_LOOPING
    End If
    CalculateWobbulation
    
    c = 1000

    Do While nFrequency * 20 > picOsc.ScaleWidth
        nFrequency = nFrequency \ 2
        c = c / 2
    Loop

    lblGraph = FormatNumber(c, 1) & " ms"

    picOsc.Cls
    picOsc.Line (0, picOsc.ScaleHeight \ 2)-(picOsc.ScaleWidth, picOsc.ScaleHeight \ 2), &H8000&
    picOsc.Line (0, (picOsc.ScaleHeight \ 2) * (1 - nVolume))-(picOsc.ScaleWidth, (picOsc.ScaleHeight \ 2) * (1 - nVolume)), &H6000&
    picOsc.Line (0, (picOsc.ScaleHeight \ 2) * (1 + nVolume))-(picOsc.ScaleWidth, (picOsc.ScaleHeight \ 2) * (1 + nVolume)), &H6000&

    If iShape = 1 Then
        For i = 0 To picOsc.ScaleWidth
            c = Sgn(Sin(i / picOsc.ScaleWidth * pi * 2 * nFrequency))
            If c = 0 Then c = 1
            picOsc.PSet (i, ((picOsc.ScaleHeight - 1) \ 2) * (1 - c * nMod * nVolume)), vbGreen
        Next
    Else
        tmp = 255 / picOsc.ScaleHeight
        picOsc.Line (0, picOsc.ScaleHeight \ 2)-(0, picOsc.ScaleHeight \ 2)
        For i = 0 To picOsc.ScaleWidth
            picOsc.Line -(i, (picOsc.ScaleHeight \ 2) * (1 - Sin(i / picOsc.ScaleWidth * pi * 2 * nFrequency) * nMod * nVolume)), vbGreen
            'picOsc.Line -(i, ((picOsc.ScaleHeight \ 2) * (1 - ((lpBuffer(i)) / picOsc.ScaleWidth) * nMod * nVolume))), vbGreen
        Next
    End If

End Sub

Private Sub bttnBuffer_Click(Index As Integer)
     
    For i = 0 To 3
      If i <> Index Then
         bttnBuffer(i).Value = False
      End If
    Next
   
    Select Case Index
     Case 0: nBasicBufferSize = 4096
     Case 1: nBasicBufferSize = 8192
     Case 2: nBasicBufferSize = 16384
     Case 3: nBasicBufferSize = 32768
    End Select
    SinBuffer nFreq, nVolume

End Sub

Private Sub bttnDB_Click(Index As Integer)
  Dim iVol As Long
  
  For i = 0 To 3
    If i <> Index Then
       bttnDB(i).Value = False
    End If
  Next
  
  Select Case Index
    Case 0: iVol = DSBVOLUME_MAX
    Case 1: iVol = -1000 ' - 10db
    Case 2: iVol = -2000 ' etc
    Case 3: iVol = -3000
  End Select
  DSB.SetVolume iVol

End Sub

Private Sub bttnOnOff_Click(Index As Integer)
    
    For i = 0 To 1
      If i <> Index Then
         bttnOnOff(i).Value = False
      End If
    Next
   
   Select Case Index
      Case 0
          DSB.Stop
          Image2(0).Visible = False
          Image2(1).Visible = True
      Case 1
          DSB.Play DSBPLAY_LOOPING
          Image2(0).Visible = True
          Image2(1).Visible = False
   End Select

End Sub

Private Sub bttnProg_Click(Index As Integer)
    DSB.Stop
    End
End Sub

Private Sub bttnSamples_Click(Index As Integer)
  
    For i = 0 To 3
      If i <> Index Then
         bttnSamples(i).Value = False
      End If
    Next
    
    Select Case Index
      Case 0: nSamples = 11025
      Case 1: nSamples = 22050
      Case 2: nSamples = 24000
      Case 3: nSamples = 48000
    End Select
    
    PCM.nFormatTag = WAVE_FORMAT_PCM
    PCM.nChannels = 1
    PCM.lSamplesPerSec = nSamples
    PCM.nBitsPerSample = 8
    PCM.nBlockAlign = 1
    PCM.lAvgBytesPerSec = nSamples * PCM.nBlockAlign
    
    HScroll1.Max = (nSamples / 10) * 2
    HScroll1.Min = HScroll1.Max / 20
    HScroll1.Value = nSamples / 10
    SinBuffer nFreq, nVolume

End Sub

Private Sub bttnShape_Click(Index As Integer)
  
    For i = 0 To 3
      If i <> Index Then
         bttnShape(i).Value = False
      End If
    Next
    
    iShape = Index
    SinBuffer nFreq, nVolume

End Sub

Private Sub chkChop_Click()
  
  If chkChop.State Then
    iCounter = 0
    iPrevVolume = DSB.GetVolume
    
    tmrChop.Interval = (1000 / iChopFrequency) / 5
    tmrChop.Enabled = True
    
  Else
  
    tmrChop.Enabled = False
    nMod = 1
    nVolume = 1
    DSB.SetVolume 0
    SinBuffer nFreq, nVolume
    
  End If
  
  
End Sub

Private Sub cmbEndSweep_Click()
      
      nFreq = 1000
      txtSweep(0).Text = nFreq
      SinBuffer nFreq, nVolume
      DSB.SetFrequency DSBFREQUENCY_ORIGINAL
      tmrSweep.Enabled = False
      bttnOnOff_Click 0

End Sub

Private Sub cmbSweep_Click()
  
  Dim StepsPerSecond As Single
  Dim TotalSteps As Single
  Dim MaxDSBFrequency As Long
    
  LED1.State = False
  LED1_Click
  
  iSweep0 = Val(txtSweep(0).Text)
  iSweep1 = Val(txtSweep(1).Text)
  iSign = 1
  
  nFreq = iSweep1 \ 2 ' First set our frequency as half of maximum
      
  StepsPerSecond = 10
  TotalSteps = mVal(txtSweep(2).Text) * StepsPerSecond
  
  If TotalSteps <= 0 Then
     Exit Sub
  End If
  
  iSweeNormValue = DSB.GetFrequency
  
  MaxDSBFrequency = iSweeNormValue * 2
  MinDSBFrequency = iSweeNormValue * (iSweep0 / iSweep1) * 2
  
  iSweepValue = MinDSBFrequency
  iSweepStep = (MaxDSBFrequency - MinDSBFrequency) / TotalSteps
  iSweepMaxValue = MaxDSBFrequency
  
  bSweepInProgress = True
  tmrSweep.Interval = 1000 \ StepsPerSecond
  
  DSB.Stop
  DSB.SetFrequency iSweepValue
  
  SinBuffer nFreq, nVolume
  LCDDisplay1.Value = FormatNumber(nFreq * (iSweepValue / iSweeNormValue), 0, , , vbFalse)
  tmrSweep.Enabled = True
  
End Sub
Private Sub Command2_Click()
   DSB.SetFrequency DSBFREQUENCY_ORIGINAL
End Sub
Private Sub ctDial1_DialChange(nValue As Integer)
    nFreq = 1 + ctDial1.Value * 22.049! * Log(1 + ctDial1.Value / 100) / Log(2)
    txtSweep(0).Text = nFreq
    SinBuffer nFreq, nVolume
End Sub

Private Sub ctDial2_DialChange(nValue As Integer)
   nVolume = nValue / 100
   SinBuffer nFreq, nVolume
End Sub

Private Sub Form_Load()
    
    lblhead.Caption = "Programmed by" & vbCrLf & "Makis Charalambous"
    nMod = 1
    sep = Chr$(9)
    nSamples = 48000
    nBasicBufferSize = 16384
    
    Set DS = DX7.DirectSoundCreate(vbNullString)
    DS.SetCooperativeLevel hWnd, DSSCL_PRIORITY
    
    PCM.nFormatTag = WAVE_FORMAT_PCM
    PCM.nChannels = 1
    PCM.lSamplesPerSec = nSamples
    PCM.nBitsPerSample = 8
    PCM.nBlockAlign = 1
    PCM.lAvgBytesPerSec = PCM.lSamplesPerSec * PCM.nBlockAlign
    DSBD.lFlags = DSBCAPS_STATIC
    
    HScroll1.Value = nSamples \ 10
    hscrChop_DialChange 30
    HScroll2_DialChange 30
    bttnSamples_Click 3
    
    ctDial2_DialChange 50
    
    nFreq = 1 + ctDial1.Value * 22.049! * Log(1 + ctDial1.Value / 1000) / Log(2)
    txtSweep(0).Text = nFreq
    SinBuffer nFreq, nVolume
    InitializeDigitalSelector
    centerform Me
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
     
     Set DSB = Nothing
     TmrProgram.Enabled = False

End Sub

Private Sub hscMod_Change()
    hscMod_Scroll
    lblAmpMod = hscMod.Value
End Sub

Private Sub hscMod_Scroll()
    If hscMod.Value = 0 Then
        tmrMod.Interval = 0
        nMod = 1
    Else
        tmrMod.Interval = 1
    End If
End Sub

Private Sub hscrChop_DialChange(nValue As Integer)
   
   iChopMarkSpace = nValue * 1.666
   lblChop = Str$(iChopMarkSpace) & " %"
   
End Sub

Private Sub HScroll1_Change()
    HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()

    On Error Resume Next
    DSB.SetFrequency HScroll1.Value * 10&
    
End Sub
Private Sub HScroll2_DialChange(nValue As Integer)
   
   lblChopFrequency.Caption = FormatNumber(nValue / 6, 0) & " Hz"
   iChopFrequency = nValue / 6
   If iChopFrequency = 0 Then iChopFrequency = 1
   tmrChop.Interval = (1000 / iChopFrequency) / 5
   
End Sub

Private Sub hscrolPercent_DialChange(nValue As Integer)
  
  wobValue = nValue
  CalculateWobbulation

End Sub

Private Sub LED1_Click()
    
  Dim tmp As Integer
  
  If LED1.State = True Then
    iWobulator = DSB.GetFrequency
    tmp = (wobpercent * iWobulator) / 100
    iWobulatorMax = iWobulator + tmp
    iWobulatorMin = iWobulator - tmp
    iWobulatorStep = tmp / 10
    wobProgress = iWobulator
    TimerWob.Interval = 10
  
    TimerWob.Enabled = True
  
  Else
   
   TimerWob.Enabled = False
   DSB.SetFrequency DSBFREQUENCY_ORIGINAL
  End If

End Sub
Private Sub scrDecimal_Change()
    
    txtDecimal.Text = Str(scrDecimal.Value)
    conFreq
    
End Sub

Private Sub TimerWob_Timer()
  
  Static signflag As Boolean
  
  If Not signflag Then
   wobProgress = wobProgress + iWobulatorStep
   If wobProgress > iWobulatorMax Then
      signflag = True
   End If
  Else
   wobProgress = wobProgress - iWobulatorStep
   If wobProgress < iWobulatorMin Then
      signflag = False
   End If
  End If
  DSB.SetFrequency wobProgress
  
End Sub

Private Sub tmrChop_Timer()
  
  Static iState As Integer
    
  iState = nMod
  
  iCounter = iCounter + 20
  If iCounter > 100 Then
      iCounter = 0
  End If
 
  If iCounter >= iChopMarkSpace Then
    nMod = False
  Else
    nMod = True
    
  End If
  
  If iState <> nMod Then
     iState = nMod
       If nMod = 0 Then
        DSB.SetVolume -10000
       Else
        DSB.SetVolume iPrevVolume
       End If
       'SinBuffer nFreq, nVolume
  End If
  DoEvents
End Sub

Private Sub tmrMod_Timer()
    
    If nModDir >= 0 Then
        nMod = nMod + 0.2! / (101 - hscMod.Value)
        If nMod > 1 Then nMod = 1: nModDir = -1
    Else
        nMod = nMod - 0.2! / (101 - hscMod.Value)
        If nMod < -1 Then nMod = -1: nModDir = 1
    End If
    SinBuffer nFreq, nVolume
   ' DSB.SetVolume (nMod * 5000) - 5000
        
End Sub

'Select Units Value

Private Sub scrUnits_Change()

    txtUnits.Text = Str(scrUnits.Value)
    conFreq
    
End Sub

'Select Tens Value

Private Sub scrTens_Change()

    txtTens.Text = Str(scrTens.Value)
    conFreq
    
End Sub

'Select Hundreds and mask leading zeroes

Private Sub scrHundreds_Change()

    txtHundreds.Text = Str(scrHundreds.Value)
    If scrHundreds.Value = 0 And scrThousands.Value = 0 Then
        txtHundreds.Text = ""
    End If

    conFreq
    
End Sub

'Select Thousands and mask leading zero

Private Sub scrThousands_Change()

    txtThousands.Text = Str(scrThousands.Value)
    If scrThousands.Value = 0 Then
        txtThousands.Text = ""
    End If
    
    conFreq
    
End Sub

'Concatenate the frequency selector settings
Private Sub conFreq()

   If DontChangeFlag Then Exit Sub
    nFreq = (scrThousands.Value * 1000) + (scrHundreds.Value * 100) + (scrTens.Value * 10) + scrUnits.Value + (scrDecimal.Value / 10)
    txtSweep(0).Text = nFreq
   
   ' If nFreq < 20 Then
   '     MsgBox "Frequency cannot be lower than 20 Hz."
   '     nFreq = 20
   '     txtTens.Text = "2"
   '     txtUnits.Text = "0"
   '     txtDecimal.Text = "0"
   ' End If
    
    SinBuffer nFreq, nVolume

End Sub

Sub InitializeDigitalSelector()
    
    'Initialize the Freqency Selector buttons
    
    scrDecimal.Max = 0
    scrDecimal.Min = 9
    scrDecimal.Value = 0
    txtDecimal.Text = Str(scrDecimal.Value / 10)
    
    scrUnits.Max = 0
    scrUnits.Min = 9
    scrUnits.Value = 0
    txtUnits.Text = Str(scrUnits.Value)
    
    scrTens.Max = 0
    scrTens.Min = 9
    scrTens.Value = 0
    txtTens.Text = Str(scrTens.Value)
    
    scrHundreds.Max = 0
    scrHundreds.Min = 9
    scrHundreds.Value = 0
    txtHundreds.Text = Str(scrHundreds.Value)
    
    scrThousands.Max = 0
    scrThousands.Min = 9
    scrThousands.Value = 1
    txtThousands.Text = Str(scrThousands.Value)

End Sub
Sub SetDigitalSelector()
       
    Dim tt As Integer
    Dim sDum As String
    Dim cc As Integer
    sDum = Trim$(Format$(nFreq, "0.0"))
    
    DontChangeFlag = True
    
    scrDecimal.Value = 0
    scrUnits.Value = 0
    scrTens.Value = 0
    scrHundreds.Value = 0
    scrThousands.Value = 0

    For tt = Len(sDum) To 1 Step -1
      Select Case cc
       Case 0
          scrDecimal.Value = Val(Mid$(sDum, tt, 1))
       Case 2
          scrUnits.Value = Val(Mid$(sDum, tt, 1))
       Case 3
          scrTens.Value = Val(Mid$(sDum, tt, 1))
       Case 4
          scrHundreds.Value = Val(Mid$(sDum, tt, 1))
       Case 5
          scrThousands.Value = Val(Mid$(sDum, tt, 1))
      End Select
      cc = cc + 1
    Next
    DontChangeFlag = False
    
End Sub
Sub CalculateWobbulation()
  
  Dim tmp As Integer
  Dim sVariation As Single

  lblWob = Str$(wobValue / 100) & " %"
  wobpercent = wobValue / 100
  
  iWobulator = DSB.GetFrequency
  tmp = (wobpercent * iWobulator) / 100
  iWobulatorMax = iWobulator + tmp
  iWobulatorMin = iWobulator - tmp
  iWobulatorStep = tmp / 10
  wobProgress = iWobulator
  TimerWob.Interval = 10
  
  sVariation = (wobpercent * nFreq) / 100
  lblFreq(0).Caption = FormatNumber(nFreq - sVariation, 1, , , vbFalse)
  lblFreq(1).Caption = FormatNumber(nFreq + sVariation, 1, , , vbFalse)

End Sub

Function mVal(vVal As Variant) As Currency
   
   On Error GoTo WrongArguments
   
   If IsNull(vVal) Then
     mVal = 0
     Exit Function
   End If
   
   If Val(vVal) = 0 Then
     mVal = 0
     Exit Function
   End If
   
   If Len(vVal) = 0 Then
     mVal = 0
     Exit Function
   End If
   
   mVal = CCur(vVal)
  
ExitMval:
   
   Exit Function

WrongArguments:
    
    mVal = 0
    Resume ExitMval
    
End Function

Private Sub TmrProgram_Timer()
    lSeconds = lSeconds + 1
End Sub
Sub centerform(frm As Form)
    frm.Move (Screen.Width - frm.Width) \ 2, (Screen.Height - frm.Height) \ 2
    Screen.MousePointer = vbNormal
End Sub

Private Sub tmrSweep_Timer()
  
  Static istep As Integer
      
  DSB.SetFrequency iSweepValue
  iSweepValue = iSweepValue + (iSweepStep * iSign)
  
  istep = istep + 1
  If istep > 10 Then
     istep = 0
     LCDDisplay1.Value = FormatNumber(nFreq * (iSweepValue / iSweeNormValue), 0, , , vbFalse)
  End If
  
  If iSweepValue >= iSweepMaxValue Then
      iSign = -1
      nFreq = iSweep1
      SinBuffer nFreq, nVolume
      DSB.SetFrequency DSBFREQUENCY_ORIGINAL
      tmrSweep.Enabled = False
      bSweepInProgress = False
  End If
'  If iSweepValue <= MinDSBFrequency Then
'      iSign = 1
'  End If
  
End Sub

Private Sub txtSweep_KeyPress(Index As Integer, KeyAscii As Integer)
     Select Case Index
        Case 2
            IntKP txtSweep(Index), 9, KeyAscii
        Case Else
           PointKP txtSweep(Index), 9, 2, KeyAscii
     End Select
End Sub
