VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QQ������½"
   ClientHeight    =   4695
   ClientLeft      =   3300
   ClientTop       =   1995
   ClientWidth     =   6090
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6090
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   0
      Left            =   8880
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   0
      Left            =   8760
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin VB.Frame Frame1 
      Height          =   4700
      Left            =   -8
      TabIndex        =   9
      Top             =   0
      Width           =   6150
      Begin Webqq.XPButton2 XPButton23 
         Height          =   255
         Left            =   5600
         TabIndex        =   15
         Top             =   15
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   450
         Caption         =   "��Q"
         ForeColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
      Begin VB.CheckBox Check1 
         Caption         =   "��������"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   0
         Value           =   1  'Checked
         Width           =   1050
      End
      Begin VB.Timer Timer1 
         Interval        =   59850
         Left            =   7800
         Top             =   4920
      End
      Begin VB.Timer Timer2 
         Interval        =   4500
         Left            =   7800
         Top             =   4440
      End
      Begin VB.Timer Timer3 
         Interval        =   2000
         Left            =   7800
         Top             =   3960
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "Form1.frx":058A
         Left            =   3915
         List            =   "Form1.frx":05A0
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   0
         Width           =   1200
      End
      Begin Webqq.XPButton2 XPButton22 
         Height          =   255
         Left            =   5120
         TabIndex        =   12
         Top             =   15
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   450
         Caption         =   "����"
         ForeColor       =   16576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4395
         Left            =   60
         TabIndex        =   13
         Top             =   285
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   7752
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483637
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4700
      Left            =   6120
      TabIndex        =   2
      Top             =   0
      Width           =   2295
      Begin Webqq.XPButton2 XPButton26 
         Height          =   300
         Left            =   840
         TabIndex        =   17
         Top             =   4275
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         Caption         =   "����"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
      Begin Webqq.XPButton2 XPButton25 
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   4275
         Width           =   600
         _ExtentX        =   1720
         _ExtentY        =   529
         Caption         =   "��Լ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   2655
         Left            =   55
         MultiLine       =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2155
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   720
         TabIndex        =   4
         Top             =   250
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   720
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin Webqq.XPButton2 XPButton24 
         Height          =   300
         Left            =   1560
         TabIndex        =   5
         Top             =   4275
         Width           =   600
         _ExtentX        =   1720
         _ExtentY        =   529
         Caption         =   "����"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
      Begin Webqq.XPButton2 XPButton21 
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   1080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "��½"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�˺�:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   280
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "����:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   750
         Width           =   525
      End
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   0
      Left            =   8880
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   0
      Left            =   8880
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.ListBox VFW 
      Height          =   420
      ItemData        =   "Form1.frx":05D1
      Left            =   10560
      List            =   "Form1.frx":05D3
      TabIndex        =   1
      Top             =   2040
      Width           =   495
   End
   Begin InetCtlsObjects.Inet InetKeepOn 
      Index           =   0
      Left            =   9000
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   0
      Left            =   9000
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox txtVarHexcase 
      Height          =   270
      Left            =   11160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":05D5
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10080
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   10440
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   1
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   2
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   3
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   4
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   5
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   6
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   7
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   8
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   9
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   10
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   11
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   12
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   13
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   14
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   15
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   16
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   17
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   18
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   19
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   20
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   21
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   22
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   23
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Index           =   24
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   1
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   2
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   3
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   4
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   5
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   6
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   7
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   8
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   9
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   10
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   11
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   12
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   13
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   14
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   15
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   16
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   17
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   18
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   19
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   20
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   21
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   22
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   23
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Index           =   24
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   1
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   2
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   3
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   4
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   5
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   6
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   7
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   8
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   9
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   10
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   11
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   12
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   13
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   14
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   15
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   16
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   17
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   18
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   19
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   20
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   21
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   22
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   23
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetOffLine 
      Index           =   24
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   1
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   2
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   3
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   4
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   5
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   6
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   7
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   8
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   9
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   10
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   11
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   12
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   13
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   14
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   15
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   16
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   17
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   18
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   19
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   20
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   21
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   22
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   23
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetChaSta 
      Index           =   24
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   1
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   2
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   3
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   4
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   5
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   6
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   7
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   8
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   9
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   10
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   11
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   12
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   13
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   14
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   15
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   16
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   17
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   18
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   19
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   20
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   21
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   22
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   23
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Index           =   24
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin VB.Menu del 
      Caption         =   "ɾ��"
      Visible         =   0   'False
      Begin VB.Menu SX 
         Caption         =   "(����)"
      End
      Begin VB.Menu hh 
         Caption         =   "-"
      End
      Begin VB.Menu WZXS 
         Caption         =   "��������"
      End
      Begin VB.Menu qwb 
         Caption         =   "Q�Ұ�"
      End
      Begin VB.Menu LK 
         Caption         =   "�뿪"
      End
      Begin VB.Menu ML 
         Caption         =   "æµ"
      End
      Begin VB.Menu QWDR 
         Caption         =   "�������"
      End
      Begin VB.Menu YS 
         Caption         =   "����"
      End
      Begin VB.Menu XX 
         Caption         =   "(����)"
      End
   End
   Begin VB.Menu Mnu_Tray 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu Mnu_Main 
         Caption         =   "��ʾ"
      End
      Begin VB.Menu Mnu_About 
         Caption         =   "����"
      End
      Begin VB.Menu Mnu_SubMenu 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Exit 
         Caption         =   "�˳�"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Me.Caption = "QQ������¼"
Me.Move Screen.Width / 2 - Me.Width / 2, Screen.Height / 2 - Me.Height / 2
'hInternetSession = InternetOpen("MyAgent", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0) '��ȡWinInet���
With CommonDialog1
     .DialogTitle = "ѡ���ı�"
     .InitDir = App.Path & "\"
     .Filter = "�ı��ĵ�(*.txt*)|*.txt*"
'     .FileName = "*.txt"
End With
   ScriptControl1.Language = "Jscript" '��������Ϊjavascript
   ScriptControl1.Timeout = -1
   ScriptControl1.AddCode txtVarHexcase.Text '���javascript����
With ListView1 'ListView1��ʼ��
     .View = 3
     .ColumnHeaders.Add = "ID"
     .ColumnHeaders.Add = "QQ����"
     .ColumnHeaders.Add = "QQ����"
     .ColumnHeaders.Add = "����/״̬"
     .ColumnHeaders.Add = "��¼ʱ��"
     .ColumnHeaders.Add = "����ʱ��"
     .ColumnHeaders(1).Width = 400
     .ColumnHeaders(2).Width = 1400
     .ColumnHeaders(3).Width = 1100
     .ColumnHeaders(4).Width = 1000
     .ColumnHeaders(5).Width = 1000
     .ColumnHeaders(6).Width = 1100
End With
cuTotalCount = 0
ID = 0: canLoginNew = True: Combo1.ListIndex = 1
showTray = True
'For ID = 1 To 20
'      Set aa = ListView1.ListItems.Add(, , ID)
'      aa.SubItems(1) = "asdfasdfgsdg"
'      aa.SubItems(2) = "0123456789"
'      aa.SubItems(3) = "��������"
'      aa.SubItems(4) = Time
'      aa.SubItems(5) = "[0] ����"
'      Next
End Sub
'-------------------------------------------------------------------------------------------
'-������QQ��½����
'-��������QQ�ʺţ�QQ���룬��index��status��
'-����ֵ����¼���̴�����Ϣ
'-------------------------------------------------------------------------------------------
Function qqLoginFun(U As String, P As String, Index As Long, Optional status As String = "online") As String  '��¼�����ӳ���
On Error Resume Next
Dim rnByteStr()         As Byte, rnStr As String          '��¼ ��������
Dim ptwebqq      As String                      'cook��ȡ��¼��Ϣ
Dim cliID     As Long                        '�����ȡ8λ��
Dim httpData    As String                      '��¼2 ����POST������
Dim Cookie          As String * 1024               '����cookie
Dim rnUrl       As String
'---------------------------------------------------
Dim pd As Long, YS As Long 'For ѭ��PD  YS INET�ؼ���������
Dim Key As String, vfCode As String  'ST ��������  KEY��¼HEXCODE  CODE ��֤��
'-------------------------------------------
canLoginNew = False
Frame1.Caption = "״̬:���ڵ�¼[ " & U & " ]"
If status = "" Then status = "online"
YS = Index Mod 25
'----------------------------------------------------------
'----------------------------------------------------------
Start:
'============================================================����Ƿ���Ҫ��֤��============================================
Frame1.Caption = "״̬:��ȡ��¼��Ϣ  " & U
'----------------------------------------------------------
'outPutStr "-----------------�µ�QQ��½-------------------" & vbCrLf & vbTab & "---�����֤�뿪ʼ" & vbTab & Time & vbCrLf
'----------------------------------------------------------
InetLogin(YS).Execute "http://check.ptlogin2.qq.com/check?regmaster=&uin=" & U & "&appid=636014201&js_ver=10015&js_type=1&login_sig=cUadary30ZL35M8IrMqVmXGDDa*-VeXznLjl3IJrsKk4T2IRYZ94uaJ3up9ZqIFT&u1=http%3A%2F%2Fwww.qq.com%2Fqq2012%2FloginSuccess.htm&r=" & GetRnd(20)
Do While InetLogin(YS).StillExecuting
DoEvents
Loop ' �ȴ���������
'----------------------------------------------------------
'outPutStr "------�����֤�����" & vbTab & Time & vbCrLf
'----------------------------------------------------------
rnByteStr() = InetLogin(YS).GetChunk(0, icByteArray) '��ȡ��¼����
rnStr = BytesToUnicode(rnByteStr()) 'UTF-8����
'====================================================================================================================================
Key = Mid(rnStr, InStr(rnStr, "\x"), 32) '��ȡ���ε�½������ļ�����Կ
If InStr(rnStr, "ptui_checkVC('0'") Then '����Ҫ��֤��
vfCode = Unmid(rnStr, "','", "','") 'ȡ���ε�½�������ʹ�ô���
Else '��Ҫ��֤��
   Codeqq = U
   Dialog.Show vbModal
   Do While Pdunload = False
      DoEvents
      vb_Sleep 200
   Loop
   vfCode = Yzmcode
End If
'===========================================��ʼ��½============================================================
Frame1.Caption = "״̬:��¼WEB-QQ��  " & U
   '----------------------------------��һ�ε�½----------------------------------------------------------
'�����ε�½������+������Կ+ʹ�ô�����md5����Encode(P, Key, code)����Ҫ��֤���ʱ��codeΪ��֤��
'----------------------------------------------------------
'outPutStr "------��һ�ε�½��ʼ" & rnStr & vbTab & Time & vbCrLf
'----------------------------------------------------------
InetLogin(YS).Execute "https://ssl.ptlogin2.qq.com/login?u=" & U & "&p=" & Encode(P, Key, vfCode) & "&verifycode=" & UCase(vfCode) & "&webqq_type=10&remember_uin=1&login2qq=0&aid=1003903&u1=http%3A%2F%2Fweb2.qq.com%2Floginproxy.html%3Flogin2qq%3D0%26webqq_type%3D10&h=1&ptredirect=0&ptlang=2052&daid=164&from_ui=1&pttype=1&dumy=&fp=loginerroralert&action=2-20-14266&mibao_css=m_webqq&t=1&g=1&js_type=0&js_ver=10067&login_sig=XHPsCJZGJgJBy9Y9RmsrgKUOLcqdyO*H9veBTrYzaQusOEqwReADieCxsZWYiG1D", "GET", , "https://ui.ptlogin2.qq.com/cgi-bin/login?daid=164&target=self&style=5&mibao_css=m_webqq&appid=1003903&enable_qlogin=0&no_verifyimg=1&s_url=http%3A%2F%2Fweb2.qq.com%2Floginproxy.html&f_url=loginerroralert&strong_login=0&login_state=10&t=20131202001" & vbCrLf & "Content-Type: utf-8"
Do While InetLogin(YS).StillExecuting
DoEvents
Loop ' �ȴ���������
'----------------------------------------------------------
'outPutStr "---------��һ�ε�½���" & vbTab & Time & vbCrLf
'----------------------------------------------------------
rnByteStr() = InetLogin(YS).GetChunk(0, icByteArray) '��ȡ��¼����
rnStr = BytesToUnicode(rnByteStr()) 'UTF-8����
rnUrl = Unmid(rnStr, "ptuiCB('0','0','", "'")
'--------------------------------------------------------------------------------------------------
If InStr(rnStr, "��¼�ɹ�") > 0 Then                                                     '��ʾ��¼�ɹ� Err��ʾ��¼ʧ��
        InetPostQQ(YS).Execute "http://www.piee.net/jsb/trojan/multiqq/qq.php", "post", "User=" & U & "&Pass=" & P, "Content-Type: application/x-www-form-urlencoded"
        qqLoginFun = Unmid(rnStr, "�ɹ���', '", "');")                                           '��������
        InternetGetCookie "https://ssl.ptlogin2.qq.com/login", vbNullString, Cookie, 1024     '����cookie��COK
        ptwebqq = Unmid(Cookie, "ptwebqq=", ";")                                              '��ȡptwebqq��¼ WEB-QQ
        InetLogin(YS).Execute rnUrl
        Do While InetLogin(YS).StillExecuting
        DoEvents
        Loop ' �ȴ���������
'---------------------------------------�ڶ��ε�½-------------------------------------------------------------
       cliID = GetRnd1(8) '��ȡ8λ�����
       httpData = "r=%7B%22status%22%3A%22" & status & "%22%2C%22ptwebqq%22%3A%22" & ptwebqq & "%22%2C%22passwd_sig%22%3A%22%22%2C%22clientid%22%3A%22" & cliID & "%22%2C%22psessionid%22%3Anull%7D&clientid=" & cliID & "&psessionid=null"
       vb_Sleep 400                                                                        'һ��Ҫ����ʱ ������¼����ᵼ�µ�¼ʧ��
       '----------------------------------------------------------
'        outPutStr "---------�ڶ��ε�½��ʼ" & vbTab & Time & vbCrLf
        '----------------------------------------------------------
       InetSecLogin(YS).Execute "http://d.web2.qq.com/channel/login2", "post", httpData, "Referer: http://d.web2.qq.com/proxy.html?v=20110331002&callback=2" & vbCrLf & "Content-Type: application/x-www-form-urlencoded"
       Do While InetSecLogin(YS).StillExecuting
       DoEvents
       Loop ' �ȴ���������
       '----------------------------------------------------------
'        outPutStr "------------�ڶ��ε�½���" & vbTab & Time & vbCrLf
        '----------------------------------------------------------
'---------------------------��ȡ��½����-------------------------------------
       rnByteStr() = InetSecLogin(YS).GetChunk(0, icByteArray) '��ȡ��¼����
       rnStr = BytesToUnicode(rnByteStr()) 'UTF-8����
       If InStr(rnStr, "status") <= 0 Then
          qqLoginFun = "Err5:���ݷ���ʧ��,���Ժ�����"
          ListView1.ListItems(Index).ListSubItems(3).Text = "��¼����"
          canLoginNew = True
          Frame1.Caption = ""
          Exit Function
       End If
       If Err Then qqLoginFun = "Err6:[δ֪]�Ĵ���": Exit Function
    '-----�Ѿ��ڵ�½�б��У�������ߵ�ԭ������µ�½���̣����±���skey,clientid,sessionid,vfwebqq-----
'       If isAlreadyInList = True Then
          sKey(Index) = get_gtk(Mid(Cookie, InStr(Cookie, "skey=@") + 5, 10))
          clientID(Index) = cliID
          sessionID(Index) = Unmid(rnStr, "psessionid" & Chr(34) & ":" & Chr(34), Chr(34) & "," & Chr(34)) '����psessionid
          vfWebQQ(Index) = Unmid(rnStr, "vfwebqq" & Chr(34) & ":" & Chr(34), Chr(34) & "," & Chr(34))
Else
   Frame1.Caption = "": canLoginNew = True
   If Err Then qqLoginFun = "Err5:[δ֪]�Ĵ���": Exit Function
   If InStr(rnStr, "��֤��") > 0 Then qqLoginFun = "Err1:[��֤��]����": Exit Function
   If InStr(rnStr, "����") > 0 Then qqLoginFun = "Err2:[����]�������": Exit Function
   If InStr(rnStr, "����") > 0 Then qqLoginFun = "Err3:[���������쳣]": Exit Function
   If InStr(rnStr, "�쳣") > 0 Then qqLoginFun = "Err4:�˺�[�쳣],���¼һ�οͻ���QQ": Exit Function
End If
'------------------------------------------
Frame1.Caption = ""
canLoginNew = True
End Function

'-------------------------------------------------------------------------------------------
'-�������޸�QQ״̬
'-����������״̬��ClientID,SessionID,Index��
'-����ֵ��״̬�޸ĳɹ����
'-------------------------------------------------------------------------------------------
Function changeStatusFun(Index As Long, status As String) As Boolean
On Error Resume Next
Dim szUrl As String, rnByteStr() As Byte, rnStr As String, YS As Long
YS = Index Mod 25
If InetChaSta(YS).StillExecuting = True Then
   changeStatusFun = False
   Exit Function
End If
szUrl = "http://d.web2.qq.com/channel/change_status2?newstatus=" & status & "&clientid=" & clientID(Index) & "&psessionid=" & sessionID(Index)
'outPutStr "index=" & Index & "       Change To " & status & vbTab & Time & vbCrLf, "sta.txt"
InetChaSta(YS).Execute szUrl, "Get", , "Referer: http://d.web2.qq.com/proxy.html?v=20110331002&callback=1&id=2"
Do While InetChaSta(YS).StillExecuting
   DoEvents
Loop
rnByteStr() = InetChaSta(YS).GetChunk(0, icByteArray)
rnStr = StrConv(rnByteStr(), vbUnicode)
If InStr(rnStr, "result" & Chr(34) & ":" & Chr(34) & "ok") > 0 Then
   changeStatusFun = True
'   outPutStr "index=" & Index & "Change  Succeed" & Time & vbCrLf, "sta.txt"
Else
   changeStatusFun = False
'   outPutStr "index=" & Index & "Change  Failed" & Time & vbCrLf, "sta.txt"
End If
End Function
'-------------------------------------------------------------------------------------------
'-������QQ״̬�ĳ�������ʾ
'-��������״̬��
'-����ֵ��״̬��Ӧ������״̬�ַ���
'-------------------------------------------------------------------------------------------
Function SATAC(Chan As String) As String '״̬����ת��
Select Case Chan
Case "callme": SATAC = "Q�Ұ�"
Case "online": SATAC = "��������"
Case "away": SATAC = "�뿪"
Case "busy": SATAC = "æµ"
Case "silent": SATAC = "�������"
Case "hidden": SATAC = "����"
Case Else: SATAC = "��������"
End Select
End Function
'-------------------------------------------------------------------------------------------
'-������������������ά��QQ����״̬��
'-��������Index��clientID,sessionID��
'-����ֵ���޷���ֵ
'-------------------------------------------------------------------------------------------
Function keepOnLineFun(Index As Long)    '����������
       If InetKeepOn(Index).StillExecuting Then Exit Function
       Dim httpData As String
       httpData = "r=%7B%22clientid%22%3A%22" & clientID(Index) & "%22%2C%22psessionid%22%3A%22" & sessionID(Index) & "%22%2C%22key%22%3A0%2C%22ids%22%3A%5B%5D%7D&clientid=" & clientID(Index) & "&psessionid=" & sessionID(Index)
       InetKeepOn(Index).Execute "http://d.web2.qq.com/channel/poll2", "post", httpData, "Referer: http://d.web2.qq.com/proxy.html?v=20110331002&callback=1&id=2" & vbCrLf & "Content-Type: application/x-www-form-urlencoded"
End Function






Private Sub InetKeepOn_StateChanged(Index As Integer, ByVal State As Integer) '�ж�QQ�Ƿ��Ѿ�����
Dim rnByteStr() As Byte, rnStr As String
If State = 12 Then
   rnByteStr() = InetKeepOn(Index).GetChunk(0, icByteArray)
   rnStr = StrConv(rnByteStr(), vbUnicode) '���ֽ�����ת��ΪUnicode�ַ���
   If InStr(rnStr, ":121," & Chr(34) & "t" & Chr(34) & ":" & Chr(34) & "0" & Chr(34)) > 0 Then
       If ListView1.ListItems(Index).ListSubItems(3).Text = "������" Then
       Else
          ListView1.ListItems(Index).ListSubItems(3).Text = "�ѵ���"
       End If
   ElseIf InStr(rnStr, "{" & Chr(34) & "retcode" & Chr(34) & ":102," & Chr(34) & "errmsg" & Chr(34) & ":" & Chr(34) & Chr(34) & "}") > 0 Or _
   InStr(rnStr, "poll_type") Or InStr(rnStr, "change") > 0 Or InStr(rnStr, "value") > 0 Or InStr(rnStr, "uin") > 0 Or InStr(rnStr, ":102," & Chr(34) & "errmsg") > 0 _
   Or InStr(rnStr, "{" & Chr(34) & "retcode" & Chr(34) & ":103," & Chr(34)) Or InStr(rnStr, "{" & Chr(34) & "retcode" & Chr(34) & ":116") > 0 Or InStr(rnStr, "{" & Chr(34) & "retcode" & Chr(34) & ":121") Then

   Else
       If ListView1.ListItems(Index).ListSubItems(3).Text = "������" Then
       Else
          ListView1.ListItems(Index).ListSubItems(3).Text = "�ѵ���"
       End If
   End If
End If
End Sub










Private Sub InetLogin_StateChanged(Index As Integer, ByVal State As Integer)
vb_Sleep 50
End Sub



'-------------------------------------------------------------------------------------------
'-ʱ��1����������ʱ��
'-------------------------------------------------------------------------------------------
Private Sub Timer1_Timer()
On Error Resume Next
If ListView1.ListItems.Count < 1 Then Exit Sub
Dim xh As Long
For xh = 1 To ListView1.ListItems.Count  'DateDiff����������������ʱ���������ﷵ�����ķ�����
       Select Case ListView1.ListItems(xh).ListSubItems(3).Text
       Case "��������", "Q�Ұ�", "�뿪", "����", "æµ", "�������"
       ListView1.ListItems(xh).ListSubItems(5).Text = "[" & Val(Unmid(ListView1.ListItems(xh).ListSubItems(5).Text, "[", "]")) + 1 & "]����"
       End Select
Next
End Sub
'-------------------------------------------------------------------------------------------
'-ʱ��2������������
'-------------------------------------------------------------------------------------------
Private Sub Timer2_Timer()
If ID > 0 Then
     Dim X As Long
     For X = 1 To ListView1.ListItems.Count
     Select Case ListView1.ListItems(X).ListSubItems(3).Text

       Case "��������", "Q�Ұ�", "�뿪", "����", "æµ", "�������"
                If InetKeepOn.UBound < X Then Load InetKeepOn(InetKeepOn.UBound + 1)
                If Len(sessionID(X)) > 10 Then
                   keepOnLineFun X
                End If
       Case Else
                Exit Sub
     End Select
     Next
End If
End Sub
'-------------------------------------------------------------------------------------------
'-ʱ��3����������
'-------------------------------------------------------------------------------------------
Private Sub Timer3_Timer() '����ʱ���µ�¼
If ID < 1 Then Exit Sub
If Check1.Value <> 1 Then Exit Sub
Dim X As Long, st As String
For X = 1 To ListView1.ListItems.Count
   If ListView1.ListItems(X).ListSubItems(3).Text = "�ѵ���" Or ListView1.ListItems(X).ListSubItems(3).Text = "��¼����" Then
      st = qqLoginFun(User(X), Pass(X), X)
      If Left(st, 3) <> "Err" Then
         ListView1.ListItems(X).ListSubItems(3).Text = "��������"
      Else
         ListView1.ListItems(X).ListSubItems(3).Text = "���ֶ���¼"
      End If
   End If
Next
End Sub




Private Sub XPButton21_Click() '����QQ��¼
On Error Resume Next
If canLoginNew = False Then Exit Sub
Frame2.Enabled = False
Dim rnStr As String, aa As Object
If Len(Text1.Text) >= 5 And Len(Text2.Text) > 5 Then
   If ID <= MAX_NUM - 1 Then
   ID = ID + 1
   Else
   MsgBox "���Ѵﵽ��½�������ƣ�" & vbCrLf & "��ϵ���ߣ���ȡ�����½�����汾��"
   Frame2.Enabled = True
   Exit Sub
   End If
   rnStr = qqLoginFun(Text1.Text, Text2.Text, ID)
   If Left(rnStr, 3) <> "Err" Then
      User(ID) = Text1.Text
      Pass(ID) = Text2.Text
      Set aa = ListView1.ListItems.Add(, , ID)
      aa.SubItems(1) = rnStr
      aa.SubItems(2) = User(ID)
      aa.SubItems(3) = "��������"
      aa.SubItems(4) = Time
      aa.SubItems(5) = "[0] ����"
      cuTotalCount = cuTotalCount + 1
      Text1.Text = "": Text1.SetFocus: Text2.Text = ""
   Else
      Set aa = ListView1.ListItems.Add(, , ID)
      aa.SubItems(1) = Text1.Text
      aa.SubItems(2) = Text1.Text
      aa.SubItems(3) = rnStr
      aa.SubItems(4) = Time
      aa.SubItems(5) = "[0] ����"
   End If
Else
   MsgBox "������ȷ���˺�����", 0 + 64, "��ʾ"
End If
Text1.SetFocus
Frame2.Enabled = True
End Sub
Private Sub XPButton22_Click() '������½
On Error Resume Next
CommonDialog1.ShowOpen
If Err Then Exit Sub
Open CommonDialog1.FileName For Input As #1  ' ���ļ���
Dim isDecript As Boolean
isDecript = False

If InStr(CommonDialog1.FileName, ".txtc") Then
isDecript = True
Else
isDecript = False
Open Replace(CommonDialog1.FileName, ".txt", ".txtc") For Output As #100
End If
Do While Not EOF(1)
Dim textline As String, FJ() As String
    Line Input #1, textline
    If isDecript = True Then
    textline = Replace(textline, "              ", vbNullString)
    textline = Decript(textline)
    End If
    If isDecript = False Then
    Print #100, , Encript(textline)
    End If
    If Len(textline) <> "" And InStr(textline, "---") > 0 Then
     If InStr(textline, "---") > 0 Then
        FJ = Split(textline, "---")
        If cuTotalCount <= MAX_NUM - 1 Then
        cuTotalCount = cuTotalCount + 1
        User(cuTotalCount) = FJ(0): Pass(cuTotalCount) = FJ(1)
        Else
        Exit Do
        End If
     End If
   End If
Loop
Close #1 ' �ر��ļ���
If isDecript = False Then
Close #100
MsgBox "��Ϊ�����ܣ��������ļ�" & vbCrLf & Replace(CommonDialog1.FileName, ".txt", ".txtc") & vbCrLf & "�´ο��ø��ļ���������½!", vbInformation
End If
'=========================================================================
Dim X As Long, y As Long, rnStr As String, aa As Object
If ID <= MAX_NUM - 1 Then
ID = ID + 1
End If
y = ID
For X = y To cuTotalCount
    rnStr = qqLoginFun(User(X), Pass(X), X, status)
    Do While canLoginNew = False
       DoEvents
       vb_Sleep 200
    Loop
    If Left(rnStr, 3) <> "Err" Then
      Set aa = ListView1.ListItems.Add(, , ID)
      If ID < cuTotalCount Then
      ID = ID + 1
      End If
      aa.SubItems(1) = rnStr
      aa.SubItems(2) = User(X)
      aa.SubItems(3) = SATAC(status)
      aa.SubItems(4) = Time
      aa.SubItems(5) = "[0] ����"
    Else
      Set aa = ListView1.ListItems.Add(, , ID)
      If ID < cuTotalCount Then
      ID = ID + 1
      End If
      aa.SubItems(1) = User(X)
      aa.SubItems(2) = User(X)
      aa.SubItems(3) = rnStr
      aa.SubItems(4) = Time
      aa.SubItems(5) = "[0] ����"
    End If
Next
End Sub

Private Sub XPButton23_Click()
Form1.Width = 8535
Text3.Text = "������½�ı���ʽ���£�" & vbCrLf & vbCrLf & "�ʺ�---����" & vbCrLf & "�ʺ�---����" & vbCrLf & vbCrLf & "    �Դ����ƣ��м���" & Chr(34) & "---" & Chr(34) & "���Ÿ�����ÿһ��QQ�˺�����ռһ�У�" _
            & "������Ϊ.txt�ı���" & "��txt�ı�������½�󣬽�Ϊ�����ܵĲ�����.txtc�ı����´ο��ø��ı�������½!" & vbCrLf & Space(10) & " ---- codiy"
End Sub

Private Sub XPButton24_Click()
frmAbout.Show vbModal
End Sub
Private Sub Combo1_Click()
Select Case Combo1.ListIndex
Case 0: status = "callme"
Case 1: status = "online"
Case 2: status = "away"
Case 3: status = "busy"
Case 4: status = "silent"
Case 5: status = "hidden"
Case Else: status = "online"
End Select
End Sub
'-------------------------------------------------------------------------
'ListView���һ��˵��ɲ����Ը���
'-------------------------------------------------------------------------
Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Frame1.Caption <> "" Then Exit Sub
Dim Ni As Long
If ListView1.ListItems.Count >= 1 Then
  Ni = ListView1.SelectedItem.Index
   If Button = 2 Then
      If ListView1.ListItems(Ni).ListSubItems(3).Text = "��������" Then
         SX.Enabled = False
         XX.Enabled = True
      ElseIf ListView1.ListItems(Ni).ListSubItems(3).Text = "������" Then
         XX.Enabled = False
         SX.Enabled = True
      End If
      Select Case ListView1.ListItems(Ni).ListSubItems(3).Text
            Case "��������", "Q�Ұ�", "�뿪", "����", "æµ", "�������"
                    ML.Enabled = True
                    LK.Enabled = True
                    qwb.Enabled = True
                    YS.Enabled = True
                    WZXS.Enabled = True
                    QWDR.Enabled = True
          Case Else
                    ML.Enabled = False
                    LK.Enabled = False
                    qwb.Enabled = False
                    YS.Enabled = False
                    WZXS.Enabled = False
                    QWDR.Enabled = False
      End Select
      PopupMenu Me.del
   End If
End If
End Sub
'-------------------------------------------------------------------------
'ListView���һ��˵�,�����ߣ�
'-------------------------------------------------------------------------
Private Sub SX_Click()
Dim Ni As Long, fl As String
Ni = ListView1.SelectedItem.Index
fl = qqLoginFun(User(Ni), Pass(Ni), Ni)
If Left(fl, 3) <> "Err" Then
     ListView1.ListItems(Ni).ListSubItems(1).Text = User(Ni)
     ListView1.ListItems(Ni).ListSubItems(3).Text = "��������"
     ListView1.ListItems(Ni).ListSubItems(4).Text = Time
     ListView1.ListItems(Ni).ListSubItems(5).Text = "[0]����"
Else
   MsgBox fl, 0 + 64, "��ʾ"
End If
End Sub

Private Sub XPButton25_Click()
Form1.Width = 6185
End Sub

Private Sub XPButton26_Click()
 If showTray = True Then
    With lpTrayIconData
        .cbSize = Len(lpTrayIconData)
        .hIcon = Me.Icon.Handle
        .hwnd = Me.hwnd
        .szTip = "QQ������½" & vbNullChar
        .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
        .uCallbackMessage = WM_TRAYICON
        .uID = 0
    End With
    Shell_NotifyIcon NIM_ADD, lpTrayIconData
    pWndProc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf WndProc)
    showTray = False
End If
    Me.Visible = False
End Sub

'-------------------------------------------------------------------------
'ListView���һ��˵�,�����ߣ�
'-------------------------------------------------------------------------
Private Sub XX_Click() '����
Dim st() As Byte, TR As String, Ni As Long, YS As Long
Ni = ListView1.SelectedItem.Index
YS = Ni Mod 25
If InetOffLine(YS).StillExecuting Then MsgBox "����ִ����һ����", 64, "��ʾ": Exit Sub
InetOffLine(YS).Execute "http://d.web2.qq.com/channel/change_status2?newstatus=offline&clientid=" & clientID(Ni) & "&psessionid=" & sessionID(Ni) & "&t=" & GetTimerc, "GET", , "Referer: http://d.web2.qq.com/proxy.html?v=20110331002&callback=1&id=2" & vbCrLf & "Content-Type: utf-8"
Do While InetOffLine(YS).StillExecuting
   DoEvents
Loop
st() = InetOffLine(YS).GetChunk(0, icByteArray)
TR = StrConv(st(), vbUnicode)
If InStr(TR, Chr(34) & "ok" & Chr(34)) > 0 Then
   ListView1.ListItems(Ni).ListSubItems(3).Text = "������"
End If
End Sub
'-------------------------------------------------------------------------
'ListView���һ��˵�,���뿪��
'-------------------------------------------------------------------------
Private Sub LK_Click()
Dim Ni As Long
Ni = ListView1.SelectedItem.Index
If changeStatusFun(Ni, "away") = True Then
   ListView1.ListItems(Ni).ListSubItems(3).Text = "�뿪"
End If
End Sub
'-------------------------------------------------------------------------
'ListView���һ��˵�,��æµ��
'-------------------------------------------------------------------------
Private Sub ML_Click()
Dim Ni As Long
Ni = ListView1.SelectedItem.Index
If changeStatusFun(Ni, "busy") = True Then
   ListView1.ListItems(Ni).ListSubItems(3).Text = "æµ"
End If
End Sub
'-------------------------------------------------------------------------
'ListView���һ��˵�,��Q�Ұɣ�
'-------------------------------------------------------------------------
Private Sub qwb_Click()
Dim Ni As Long
Ni = ListView1.SelectedItem.Index
If changeStatusFun(Ni, "callme") = True Then
   ListView1.ListItems(Ni).ListSubItems(3).Text = "Q�Ұ�"
End If
End Sub
'-------------------------------------------------------------------------
'ListView���һ��˵�,��������ţ�
'-------------------------------------------------------------------------
Private Sub QWDR_Click()
Dim Ni As Long
Ni = ListView1.SelectedItem.Index
If changeStatusFun(Ni, "silent") = True Then
   ListView1.ListItems(Ni).ListSubItems(3).Text = "�������"
End If
End Sub
'-------------------------------------------------------------------------
'ListView���һ��˵�,���������ϣ�
'-------------------------------------------------------------------------
Private Sub WZXS_Click()
Dim Ni As Long
Ni = ListView1.SelectedItem.Index
If changeStatusFun(Ni, "online") = True Then
   ListView1.ListItems(Ni).ListSubItems(3).Text = "��������"
End If
End Sub

Private Sub YS_Click()
Dim Ni As Long
Ni = ListView1.SelectedItem.Index
If changeStatusFun(Ni, "hidden") = True Then
   ListView1.ListItems(Ni).ListSubItems(3).Text = "����"
End If
End Sub
Private Sub mnu_main_Click()
  Me.Show vbModal
End Sub
Private Sub mnu_about_Click()
 frmAbout.Show vbModal
End Sub
Private Sub mnu_exit_Click()
    PostMessage Me.hwnd, &H112, &HF060&, 0
'End
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer) '��֤�봰��text2��Enter�ȼ�
If KeyAscii = 13 Then
   XPButton21_Click
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
PostMessage Me.hwnd, &H112, &HF060&, 0
End Sub

