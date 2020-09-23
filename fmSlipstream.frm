VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form fmSlipstream 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Slipstreamer"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   Icon            =   "fmSlipstream.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   600
      Top             =   4560
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   60
      Top             =   4560
   End
   Begin VB.PictureBox Picture2 
      Height          =   75
      Left            =   120
      ScaleHeight     =   15
      ScaleWidth      =   7635
      TabIndex        =   2
      Top             =   600
      Width           =   7695
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   75
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   7845
      TabIndex        =   1
      Top             =   0
      Width           =   7905
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   6588
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   4
      TabsPerRow      =   7
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   128
      TabCaption(0)   =   "Start"
      TabPicture(0)   =   "fmSlipstream.frx":30DA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Combo1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbTimer"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label19"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Select"
      TabPicture(1)   =   "fmSlipstream.frx":30F6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Copy"
      TabPicture(2)   =   "fmSlipstream.frx":3112
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmAccept"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "ProgressBar1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmCopy"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame9"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label25(1)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label25(0)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "lbTsize"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "lbFsize"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label24"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "lbSSCopy"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label23"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "Extract"
      TabPicture(3)   =   "fmSlipstream.frx":312E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame11"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame10"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Slipstream"
      TabPicture(4)   =   "fmSlipstream.frx":314A
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame12"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame13"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Cleanup"
      TabPicture(5)   =   "fmSlipstream.frx":3166
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "?"
      TabPicture(6)   =   "fmSlipstream.frx":3182
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Picture3"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Label10"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Label9"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "Label8"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "Label7"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "Label6"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "Label5"
      Tab(6).Control(6).Enabled=   0   'False
      Tab(6).Control(7)=   "Label4"
      Tab(6).Control(7).Enabled=   0   'False
      Tab(6).Control(8)=   "Label3"
      Tab(6).Control(8).Enabled=   0   'False
      Tab(6).Control(9)=   "Label2"
      Tab(6).Control(9).Enabled=   0   'False
      Tab(6).Control(10)=   "Label1"
      Tab(6).Control(10).Enabled=   0   'False
      Tab(6).ControlCount=   11
      Begin VB.Frame Frame13 
         Height          =   375
         Left            =   120
         TabIndex        =   78
         Top             =   3180
         Width           =   7695
      End
      Begin VB.Frame Frame12 
         Caption         =   "SLIPSTREAM OPERATION"
         ForeColor       =   &H00000080&
         Height          =   2475
         Left            =   120
         TabIndex        =   77
         Top             =   660
         Width           =   7695
         Begin VB.PictureBox Picture6 
            Height          =   1035
            Left            =   1800
            ScaleHeight     =   975
            ScaleWidth      =   4095
            TabIndex        =   81
            Top             =   1020
            Width           =   4155
            Begin MSComCtl2.Animation Animation1 
               Height          =   915
               Left            =   0
               TabIndex        =   82
               Top             =   0
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   1614
               _Version        =   393216
               Center          =   -1  'True
               FullWidth       =   273
               FullHeight      =   61
            End
         End
         Begin VB.CommandButton cmSlipstream 
            Height          =   555
            Left            =   300
            Picture         =   "fmSlipstream.frx":319E
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   540
            Width           =   555
         End
         Begin VB.Label lbMsg 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Please be patient while the Service Pack is Slipstreamed"
            Height          =   195
            Left            =   1860
            TabIndex        =   83
            Top             =   2160
            Width           =   4035
         End
         Begin VB.Label Label26 
            Caption         =   "SLIPSTREAM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   1080
            TabIndex        =   80
            Top             =   660
            Width           =   1560
         End
      End
      Begin VB.Frame Frame11 
         Height          =   375
         Left            =   -74880
         TabIndex        =   75
         Top             =   2160
         Width           =   7695
      End
      Begin VB.Frame Frame10 
         Caption         =   "Extract Service Pack"
         ForeColor       =   &H00000080&
         Height          =   1455
         Left            =   -74880
         TabIndex        =   69
         Top             =   660
         Width           =   7695
         Begin VB.CommandButton cmExtractSP 
            Caption         =   "Extract SP"
            Height          =   315
            Left            =   180
            TabIndex        =   74
            Top             =   960
            Width           =   1515
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Product to Slipstream:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   5
            Left            =   180
            TabIndex        =   73
            Top             =   300
            Width           =   1545
         End
         Begin VB.Label lbProdA 
            Caption         =   "Product"
            Height          =   195
            Left            =   2400
            TabIndex        =   72
            Top             =   300
            Width           =   5055
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Path to Service Pack:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   71
            Top             =   540
            Width           =   1560
         End
         Begin VB.Label lbPathA 
            Caption         =   "Path"
            Height          =   195
            Left            =   2400
            TabIndex        =   70
            Top             =   540
            Width           =   5055
         End
      End
      Begin VB.CommandButton cmAccept 
         Caption         =   "Accept"
         Height          =   315
         Left            =   -68340
         TabIndex        =   66
         Top             =   2280
         Width           =   1095
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   -73320
         TabIndex        =   63
         Top             =   2700
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.CommandButton cmCopy 
         Caption         =   "COPY CD"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74820
         TabIndex        =   60
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Frame Frame9 
         Caption         =   "Selected Settings"
         ForeColor       =   &H00000080&
         Height          =   1575
         Left            =   -74880
         TabIndex        =   50
         Top             =   660
         Width           =   7635
         Begin VB.Label lbPath 
            Caption         =   "Path"
            Height          =   195
            Left            =   2400
            TabIndex        =   58
            Top             =   1200
            Width           =   5055
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Path to Service Pack:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   57
            Top             =   1200
            Width           =   1560
         End
         Begin VB.Label lbVoll 
            Caption         =   "Volume Label"
            Height          =   195
            Left            =   2400
            TabIndex        =   56
            Top             =   900
            Width           =   5055
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Volume Label in CDROM:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   55
            Top             =   900
            Width           =   1860
         End
         Begin VB.Label lbSvcp 
            Caption         =   "SP"
            Height          =   195
            Left            =   2400
            TabIndex        =   54
            Top             =   600
            Width           =   5055
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Service Pack to Slipstream:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   53
            Top             =   600
            Width           =   1950
         End
         Begin VB.Label lbProd 
            Caption         =   "Product"
            Height          =   195
            Left            =   2400
            TabIndex        =   52
            Top             =   300
            Width           =   5055
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Product to Slipstream:"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   51
            Top             =   300
            Width           =   1545
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Select Directory && File Name where the Service Pack Resides"
         ForeColor       =   &H00000080&
         Height          =   735
         Left            =   -74880
         TabIndex        =   47
         Top             =   2880
         Width           =   7695
         Begin VB.CommandButton cmSelect 
            Caption         =   "Select SP"
            Enabled         =   0   'False
            Height          =   315
            Left            =   180
            TabIndex        =   48
            Top             =   300
            Width           =   1035
         End
         Begin VB.Label lbSpPath 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1440
            TabIndex        =   49
            Top             =   300
            Width           =   6075
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1695
         Left            =   -72900
         TabIndex        =   37
         Top             =   1140
         Width           =   5715
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00E0E0E0&
            Height          =   855
            Left            =   3360
            ScaleHeight     =   795
            ScaleWidth      =   2115
            TabIndex        =   40
            Top             =   240
            Width           =   2175
            Begin VB.Label lbRoot 
               AutoSize        =   -1  'True
               BackColor       =   &H80000018&
               BackStyle       =   0  'Transparent
               Caption         =   "-Root Folder\"
               ForeColor       =   &H00404000&
               Height          =   195
               Left            =   240
               TabIndex        =   44
               Top             =   180
               Width           =   945
            End
            Begin VB.Label lbSpex 
               AutoSize        =   -1  'True
               BackColor       =   &H80000018&
               BackStyle       =   0  'Transparent
               Caption         =   "-SP Extract Folder"
               ForeColor       =   &H00404000&
               Height          =   195
               Left            =   660
               TabIndex        =   43
               Top             =   540
               Width           =   1275
            End
            Begin VB.Label lbCopy 
               AutoSize        =   -1  'True
               BackColor       =   &H80000018&
               BackStyle       =   0  'Transparent
               Caption         =   "-CD Copy Folder"
               ForeColor       =   &H00404000&
               Height          =   195
               Left            =   660
               TabIndex        =   42
               Top             =   360
               Width           =   1155
            End
            Begin VB.Label lbDrive 
               AutoSize        =   -1  'True
               BackColor       =   &H80000018&
               BackStyle       =   0  'Transparent
               Caption         =   "-Drive\"
               ForeColor       =   &H00404000&
               Height          =   195
               Left            =   120
               TabIndex        =   41
               Top             =   0
               Width           =   495
            End
         End
         Begin VB.ListBox List1 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   840
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   3135
         End
         Begin VB.CommandButton cmCreateTree 
            Caption         =   "Create Directories"
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   38
            Top             =   1260
            Width           =   1515
         End
         Begin VB.Label Label21 
            Caption         =   "Select drive to continue..."
            Height          =   195
            Left            =   3360
            TabIndex        =   46
            Top             =   1380
            Width           =   2175
         End
         Begin VB.Label Label20 
            Caption         =   "Directory Tree for Operation"
            Height          =   195
            Left            =   3360
            TabIndex        =   45
            Top             =   1140
            Width           =   2175
         End
      End
      Begin VB.Frame Frame6 
         Height          =   495
         Left            =   -72900
         TabIndex        =   35
         Top             =   600
         Width           =   5715
         Begin VB.Label lbDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   180
            Width           =   795
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Program Selection"
         ForeColor       =   &H00000080&
         Height          =   2235
         Left            =   -74880
         TabIndex        =   29
         Top             =   600
         Width           =   1875
         Begin VB.OptionButton opOFXP 
            Caption         =   "MS Office XP"
            ForeColor       =   &H00404000&
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1800
            Width           =   1575
         End
         Begin VB.OptionButton opO2K3 
            Caption         =   "MS Office 2003"
            ForeColor       =   &H00404000&
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1440
            Width           =   1575
         End
         Begin VB.OptionButton opMSXP 
            Caption         =   "MS Windows XP"
            ForeColor       =   &H00404000&
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton opS2K3 
            Caption         =   "MS Server 2003"
            ForeColor       =   &H00404000&
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton opMS2K 
            Caption         =   "MS Windows 2000"
            ForeColor       =   &H00404000&
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1080
            Width           =   1695
         End
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   -74880
         TabIndex        =   26
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Frame Frame4 
         Height          =   675
         Left            =   -71880
         TabIndex        =   25
         Top             =   2220
         Width           =   4575
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Note: The Service Pack used MUST be the FULL Version, not the Express Version - Links to the left are Full Version"
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   180
            Width           =   4335
         End
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   -71880
         TabIndex        =   23
         Top             =   1320
         Width           =   4575
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   $"fmSlipstream.frx":3E68
            ForeColor       =   &H00404000&
            Height          =   615
            Left            =   120
            TabIndex        =   24
            Top             =   180
            Width           =   4335
         End
      End
      Begin VB.Frame Frame1 
         Height          =   675
         Left            =   -71880
         TabIndex        =   21
         Top             =   600
         Width           =   4575
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Be sure to note the folder where the Service Pack is being saved during download."
            ForeColor       =   &H00404000&
            Height          =   435
            Left            =   120
            TabIndex        =   22
            Top             =   180
            Width           =   4335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Service Pack Download Links"
         ForeColor       =   &H00000080&
         Height          =   2295
         Left            =   -74880
         TabIndex        =   14
         Top             =   600
         Width           =   2895
         Begin VB.PictureBox Picture4 
            Height          =   75
            Left            =   180
            ScaleHeight     =   15
            ScaleWidth      =   2475
            TabIndex        =   20
            Top             =   1320
            Width           =   2535
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Download Office XP SP3"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   19
            ToolTipText     =   "OfficeXpSp3-kb832671-fullfile-enu.exe : 58.9mb"
            Top             =   1860
            Width           =   1785
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Download Windows 2000 SP4"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   18
            ToolTipText     =   "W2KSP4_EN.EXE : 129mb"
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Download Windows Server 2003 SP1"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   17
            ToolTipText     =   "WindowsServer2003-KB889101-SP1-x86-ENU.exe : 329mb"
            Top             =   660
            Width           =   2685
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Download Windows XP SP2"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   16
            ToolTipText     =   "WindowsXP-KB835935-SP2-ENU.exe : 266mb"
            Top             =   360
            Width           =   2025
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Download Office 2003 SP2"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   15
            ToolTipText     =   "Office2003SP2-KB887616-FullFile-ENU.exe : 101mb"
            Top             =   1560
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   75
         Left            =   -74580
         ScaleHeight     =   15
         ScaleWidth      =   6915
         TabIndex        =   13
         Top             =   3060
         Width           =   6975
      End
      Begin VB.Label lbTimer 
         AutoSize        =   -1  'True
         Caption         =   "F"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   -67260
         TabIndex        =   76
         Top             =   0
         Width           =   90
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
         Height          =   195
         Index           =   1
         Left            =   -68640
         TabIndex        =   68
         Top             =   3180
         Width           =   795
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source"
         Height          =   195
         Index           =   0
         Left            =   -70080
         TabIndex        =   67
         Top             =   3180
         Width           =   510
      End
      Begin VB.Label lbTsize 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         Height          =   255
         Left            =   -68640
         TabIndex        =   65
         Top             =   3360
         Width           =   1350
      End
      Begin VB.Label lbFsize 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         Height          =   255
         Left            =   -70080
         TabIndex        =   64
         Top             =   3360
         Width           =   1350
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "File Copy:"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   -73320
         TabIndex        =   62
         Top             =   3180
         Width           =   690
      End
      Begin VB.Label lbSSCopy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drive\Root\Folder\"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -73320
         TabIndex        =   61
         Top             =   2940
         Width           =   1380
      End
      Begin VB.Label Label23 
         Caption         =   "Please verify that the information above is correct and then press [Accept]"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   -74880
         TabIndex        =   59
         Top             =   2280
         Width           =   7635
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   $"fmSlipstream.frx":3F1F
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   -72540
         TabIndex        =   28
         Top             =   3000
         Width           =   5175
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Here"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -67980
         TabIndex        =   12
         Top             =   2700
         Width           =   345
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To view the webpage for further instructions go"
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   -71400
         TabIndex        =   11
         Top             =   2700
         Width           =   3330
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Here"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -72240
         TabIndex        =   10
         Top             =   2700
         Width           =   345
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "or"
         Height          =   195
         Left            =   -72480
         TabIndex        =   9
         Top             =   2700
         Width           =   135
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Here"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -72900
         TabIndex        =   8
         Top             =   2700
         Width           =   345
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Download [ bbie.exe]"
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   -74580
         TabIndex        =   7
         Top             =   2700
         Width           =   1515
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   $"fmSlipstream.frx":3FE0
         ForeColor       =   &H8000000D&
         Height          =   435
         Left            =   -74880
         TabIndex        =   6
         Top             =   2100
         Width           =   7635
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"fmSlipstream.frx":40A5
         ForeColor       =   &H8000000D&
         Height          =   795
         Left            =   -74880
         TabIndex        =   5
         Top             =   1260
         Width           =   7635
      End
      Begin VB.Label Label2 
         Caption         =   "* MicroSoft Office [All Versions] MUST be copied using a command line or the extracted files will be over 900mb !"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -74880
         TabIndex        =   4
         Top             =   3300
         Width           =   7005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"fmSlipstream.frx":41F4
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   -74880
         TabIndex        =   3
         Top             =   600
         Width           =   7635
      End
   End
End
Attribute VB_Name = "fmSlipstream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cBIT As Long

Private Sub cmAccept_Click()

 cmCopy.Enabled = True: cmAccept.Enabled = False

 sCDRM = Left(Combo1.Text, 3)
 sCOPY = lbDrive & lbRoot & Chr(92) & lbCopy & Chr(92)
 sSPEX = lbDrive & lbRoot & Chr(92) & lbSpex & Chr(92)
 sSVCP = lbSpPath
 sSPNM = lbSvcp
 
End Sub

Private Sub cmCopy_Click()
 
 cBIT = 0
 
 FSIZE = GetPathSize(sCDRM)
 TSIZE = GetPathSize(sCOPY)
 
 If UCase(Left(sSPNM, 5)) = "OFFIC" Then
  cBIT = 1
 End If
 
  With ProgressBar1
   .Min = 0
   .Max = FSIZE + 10
  End With

 lbFsize = FSIZE
 lbTsize = TSIZE
 
 If cBIT = 0 Then
  Timer2.Enabled = True
  Call XCopyFile(sCDRM, sCOPY)
  Timer2.Enabled = False
  ProgressBar1.Value = 0
  lbSSCopy = "Finsished - Proceed to [Extract]"
  lbTsize = GetPathSize(sCOPY)
 ElseIf cBIT = 1 Then
  Clipboard.Clear
  Clipboard.SetText sCOPY
  MsgBox "1: Follow the Setup routine to complete the process." & vbCrLf & _
         "2: DIRECT THE INTSALLATION TO: " & sCOPY & vbCrLf & _
         "   [The Path has copied to the clipboard... Paste it in the {Install Location} box." & vbCrLf & _
         "3: Have the SERIAL NUMBER ready", vbOKOnly, "Instructions"
  Call RunProcess(sCDRM & oSetup, 1)
 End If
 
End Sub

Private Sub cmCreateTree_Click()
 
 cmSelect.Enabled = True
 
 Call MakeDir(lbDrive & lbRoot)
 Call MakeDir(lbDrive & lbRoot & Chr(92) & lbCopy)
 Call MakeDir(lbDrive & lbRoot & Chr(92) & lbSpex)
 
End Sub

Private Sub cmExtractSP_Click()
 
 Dim cmdSTRING As String
 'OFFICE:
    'DRIVE:\FOLDER\ServicePack_EXE /t:DRIVE:\FOLDER /c
 'WINDOWS:
    'DRIVE:\FOLDER\ServicePack_EXE /x:DRIVE:\FOLDER
 
 If cBIT = 0 Then
  'WINDOWS
  cmdSTRING = sSVCP & Chr(32) & SWX & sSPEX
  Call RunProcess(cmdSTRING, 2)
  cBIT = 0
 ElseIf cBIT = 1 Then
  'OFFICE
  cmdSTRING = sSVCP & Chr(32) & SWT & sSPEX & Chr(32) & SWC
  Call RunProcess(cmdSTRING, 2)
  cBIT = 0
 End If
 
End Sub

Private Sub cmSelect_Click()

 lbSpPath = BrowseFolders(hwnd, "Select Service Pack File", BrowseForEverything, CSIDL_DRIVES)
 
 If Right(UCase(lbSpPath.Caption), 4) <> ".EXE" Then
  lbSpPath = "Please Select FILE From Browse Dialog"
 Else
  lbDesc = "PROCEED TO [COPY] PROCESS"
 End If
 
End Sub

Private Sub cmSlipstream_Click()
 
 Dim sSTR As Long
 Dim qBIT As Long
 
 Dim rPROC1 As String
 Dim rPROC2 As String
 
 Dim rPROCA As String
     rPROCA = "MSIEXEC /p "
 Dim rPROCB As String
     rPROCB = " ShortFileNames=True /qb /quiet"
     
 cBIT = 0
 lbMsg.ForeColor = &HFF&
 DoEvents
 
 If UCase(Left(sSPNM, 5)) = "OFFIC" Then
 sSTR = InStr(1, lbSvcp, "887616", vbTextCompare)
  If sSTR > 0 Then
   rPROC1 = rPROCA & sSPEX & "MAINSP2ff.msp" & SWA & sCOPY & "PRO11.MSI" & rPROCB
   rPROC2 = rPROCA & sSPEX & "OWC11SP2ff.msp" & SWA & sCOPY & "OWC11.MSI" & rPROCB
   qBIT = 10
   GoTo SEL_OfficePack
  Else
    sSTR = InStr(1, lbSvcp, "832671", vbTextCompare)
     If sSTR > 0 Then
      rPROC1 = rPROCA & sSPEX & "MAINSP3FF.MSP" & SWA & sCOPY & "PRO11.MSI" & rPROCB
      rPROC2 = rPROCA & sSPEX & "OWC10SP3FF.MSP" & SWA & sCOPY & "OWC11.MSI" & rPROCB
      qBIT = 20
      GoTo SEL_OfficePack
     Else
      Exit Sub
     End If
  End If
 Else
   GoTo SEL_WindowsPack
 End If
  
SEL_OfficePack:

Call FFResouce(App.Path & Chr(92) & "check_file.avi", RFA, "AVI")

   With Animation1
    .Open App.Path & "\check_file.avi"
    .Play
     DoEvents
   End With

 Select Case qBIT
  Case 10
   Call RunProcess(rPROC1, 2)
   Call RunProcess(rPROC2, 2)
   Animation1.Stop
   lbMsg = "FINISHED"
   Kill App.Path & "\check_file.avi"
   qBIT = 0: rPROC1 = Empty: rPROC2 = Empty
  Case 20
   Call RunProcess(rPROC1, 2)
   Call RunProcess(rPROC2, 2)
   Animation1.Stop
   lbMsg = "FINISHED"
   Kill App.Path & "\check_file.avi"
   qBIT = 0: rPROC1 = Empty: rPROC2 = Empty
 End Select
 Exit Sub
 
SEL_WindowsPack:

   With Animation1
    .Open App.Path & "\check_file.avi"
    .Play
     DoEvents
   End With

'DRIVE:\FOLDER\i386\Update\Update.exe -s:DRIVE:\FOLDER\?i386
'DRIVE:\FOLDER\i386\Update\Update.exe -s:DRIVE:\FOLDER\?i386
'DRIVE:\FOLDER\i386\Update\Update.exe -s:DRIVE:\FOLDER\?i386

 If InStr(1, lbSvcp, "W2KSP4", vbTextCompare) > 0 Or _
    InStr(1, lbSvcp, "889101", vbTextCompare) > 0 Or _
    InStr(1, lbSvcp, "835935", vbTextCompare) > 0 Then
     rPROC1 = sSPEX & i386 & "Update.exe" & SWS & sCOPY
     sSTR = 100
      Call RunProcess(rPROC1, 2)
 End If
  
  If sSTR > 0 Then
   Animation1.Stop
   lbMsg = "FINISHED"
   Kill App.Path & "\check_file.avi"
   qBIT = 0: rPROC1 = Empty
   sSTR = 0
  End If
  
End Sub

Private Sub Form_Load()

 If App.PrevInstance Then End
    App.TaskVisible = True
    
    SSTab1.Tab = 0
    Call Get_CdrList(Combo1)
    Call Get_HddList(List1)
    
    If IsDriveReady(Left(Combo1.Text, 3)) = False Then
     Timer1.Enabled = True
    Else
     Timer1.Enabled = False
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Call Shutdown(False)
End Sub

Private Sub Shutdown(Optional ByVal Force As Boolean = False)
 
 Dim i As Long

 On Error Resume Next
 
      For i = Forms.Count - 1 To 0 Step -1
         Unload Forms(i)
          If Not Force Then
            If Forms.Count > i Then
               Exit Sub
            End If
          End If
      Next i

 If Force Or (Forms.Count = 0) Then Close
 If Force Or (Forms.Count > 0) Then End
      
End Sub

Private Sub MLink()
    MakeLink Label6, Startup
    MakeLink Label8, Startup
    MakeLink Label10, Startup
    MakeLink Label11, Startup
    MakeLink Label12, Startup
    MakeLink Label13, Startup
    MakeLink Label14, Startup
    MakeLink Label15, Startup
End Sub
Private Sub MMove()
    MakeLink Label6, FormMove
    MakeLink Label8, FormMove
    MakeLink Label10, FormMove
    MakeLink Label11, FormMove
    MakeLink Label12, FormMove
    MakeLink Label13, FormMove
    MakeLink Label14, FormMove
    MakeLink Label15, FormMove
End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 lbDesc = "Description"
End Sub

Private Sub Label6_Click()
    MakeLink Label6, Click, Me, "http://69.90.47.6/mybootdisks.com/mybootdisks_com/nu2/bbie10.zip"
    '"http://69.90.47.6/mybootdisks.com/mybootdisks_com/nu2/bbie10.zip"
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MakeLink Label6, LinkMove
End Sub

Private Sub Label8_Click()
    MakeLink Label8, Click, Me, "http://downloadmirror.dll-downloads.com/~webpromo/mirrors/nu2.nu/bbie10.zip"
    '"http://downloadmirror.dll-downloads.com/~webpromo/mirrors/nu2.nu/bbie10.zip"
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MakeLink Label8, LinkMove
End Sub

Private Sub Label10_Click()
    MakeLink Label10, Click, Me, "http://nu2.nu/bbie/index.php"
    '"http://nu2.nu/bbie/index.php"
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MakeLink Label10, LinkMove
End Sub

Private Sub Label11_Click()
    MakeLink Label11, Click, Me, "http://download.microsoft.com/download/9/b/3/9b37f157-123d-41fd-a3f4-f4aedd0cc847/Office2003SP2-KB887616-FullFile-ENU.exe"
    '"http://download.microsoft.com/download/9/b/3/9b37f157-123d-41fd-a3f4-f4aedd0cc847/Office2003SP2-KB887616-FullFile-ENU.exe"
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MakeLink Label11, LinkMove
End Sub

Private Sub Label12_Click()
    MakeLink Label12, Click, Me, "http://download.microsoft.com/download/1/6/5/165b076b-aaa9-443d-84f0-73cf11fdcdf8/WindowsXP-KB835935-SP2-ENU.exe"
    '"http://download.microsoft.com/download/1/6/5/165b076b-aaa9-443d-84f0-73cf11fdcdf8/WindowsXP-KB835935-SP2-ENU.exe"
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MakeLink Label12, LinkMove
End Sub

Private Sub Label13_Click()
    MakeLink Label13, Click, Me, "http://download.microsoft.com/download/1/2/7/127c5938-d36a-4405-9df1-f00d57495652/WindowsServer2003-KB889101-SP1-x86-ENU.exe"
    '"http://download.microsoft.com/download/1/2/7/127c5938-d36a-4405-9df1-f00d57495652/WindowsServer2003-KB889101-SP1-x86-ENU.exe"
End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MakeLink Label13, LinkMove
End Sub

Private Sub Label14_Click()
    MakeLink Label14, Click, Me, "http://download.microsoft.com/download/E/6/A/E6A04295-D2A8-40D0-A0C5-241BFECD095E/W2KSP4_EN.EXE"
    '"http://download.microsoft.com/download/E/6/A/E6A04295-D2A8-40D0-A0C5-241BFECD095E/W2KSP4_EN.EXE"
End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MakeLink Label14, LinkMove
End Sub

Private Sub Label15_Click()
    MakeLink Label15, Click, Me, "http://download.microsoft.com/download/9/1/F/91FFC6B2-0745-470B-8DD3-1285B85DB12B/OfficeXpSp3-kb832671-fullfile-enu.exe"
    '"http://download.microsoft.com/download/9/1/F/91FFC6B2-0745-470B-8DD3-1285B85DB12B/OfficeXpSp3-kb832671-fullfile-enu.exe"
End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MakeLink Label15, LinkMove
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call MMove
End Sub

Private Sub List1_Click()
 
 Dim X As Long
 For X = 0 To List1.ListCount - 1
  If List1.Selected(X) = True Then
   lbDrive = Left(List1.Text, 3)
   cmCreateTree.Enabled = True
  End If
 Next X
 
 lbRoot = "~SS_Temp"
 lbCopy = "~SS_Copy"
 lbSpex = "~SP_Copy"
 
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 lbDesc = "Please select drive to create directory tree and copy files..."
End Sub

Private Sub opMS2K_Click()
 List1.Enabled = True
 lbProd = opMS2K.Caption
End Sub

Private Sub opMS2K_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 lbDesc = "Windows 2000 [Server or Professional version]"
End Sub

Private Sub opMSXP_Click()
 List1.Enabled = True
 lbProd = opMSXP.Caption
 lbSvcp = "WindowsXP-KB835935-SP2-ENU"
End Sub

Private Sub opMSXP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 lbDesc = "Windows XP [Home or Professional version]"
End Sub

Private Sub opO2K3_Click()
 List1.Enabled = True
 lbProd = opO2K3.Caption
 lbSvcp = "Office2003SP2-KB887616-FullFile-ENU"
End Sub

Private Sub opO2K3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 lbDesc = "Office 2003 [All versions]"
End Sub

Private Sub opOFXP_Click()
 List1.Enabled = True
 lbProd = opOFXP.Caption
 lbSvcp = "OfficeXpSp3-kb832671-fullfile-enu"
End Sub

Private Sub opOFXP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 lbDesc = "Office XP [NOTE: Only works with ENTERPRISE version (2002)]"
End Sub

Private Sub opS2K3_Click()
 List1.Enabled = True
 lbProd = opS2K3.Caption
 lbSvcp = "WindowsServer2003-KB889101-SP1-x86-ENU"
End Sub

Private Sub opS2K3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 lbDesc = "Windows Server 2003 [All versions]"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
 
 Dim sSTR As Long
 
 Select Case SSTab1.Tab
 
  Case 0
  Case 1
   SSTab1.TabEnabled(0) = False
  Case 2
   lbVoll = Combo1.Text
   lbPath = lbSpPath
   lbSSCopy = lbDrive & lbRoot & Chr(92) & lbCopy & Chr(92)
   sSTR = InStr(1, lbSpPath, lbSvcp, vbTextCompare)
   
   If sSTR <= 0 Then
    MsgBox "The Service Pack selected does NOT match the MS Product", vbCritical, "Product Mismatch"
    cmAccept.Enabled = False
    cmExtractSP.Enabled = False
    SSTab1.TabEnabled(PreviousTab) = True
   Else
    SSTab1.TabEnabled(PreviousTab) = False
    cmAccept.Enabled = True
    cmExtractSP.Enabled = True
   End If
  
   If IsDriveReady(Left(Combo1.Text, 3)) = False Then
    Timer1.Enabled = True
    cmCopy.Enabled = False
   Else
    Timer1.Enabled = False
    'cmCopy.Enabled = True
   End If
  
  
  Case 3
   lbPathA = sSVCP
   lbProdA = lbProd
   cBIT = 0
 
   If UCase(Left(sSPNM, 5)) = "OFFIC" Then
    cBIT = 1
   End If
  
   'SSTab1.TabEnabled(0) = False
   'SSTab1.TabEnabled(1) = False
   'SSTab1.TabEnabled(2) = False
  
  Case 4
   SSTab1.TabEnabled(0) = False
   SSTab1.TabEnabled(1) = False
   SSTab1.TabEnabled(2) = False
   SSTab1.TabEnabled(3) = False
  Case 5
  Case 6
 
 End Select
 
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call MMove
End Sub

Private Sub Timer1_Timer()
 
 If IsDriveReady(Left(Combo1.Text, 3)) = True Then
  lbTimer = "F"
  Call Get_CdrList(Combo1)
  cmCopy.Enabled = True
  lbVoll = Combo1.Text
  Timer1.Enabled = False
 Else
  lbTimer = "T"
  lbVoll = "<NO CD>"
  cmCopy.Enabled = False
  Timer1.Enabled = True
 End If
 
End Sub

Private Sub Timer2_Timer()
 
 TSIZE = GetPathSize(sCOPY)
 
 With lbSSCopy
  .Caption = vbNullString
  .Caption = sCOPY & sFHND
 End With
 
 ProgressBar1.Value = TSIZE
 lbTsize = TSIZE
 DoEvents
 
End Sub
