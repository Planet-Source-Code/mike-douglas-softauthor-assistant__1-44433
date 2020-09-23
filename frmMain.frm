VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   BackColor       =   &H8000000B&
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12360
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9645
   ScaleWidth      =   12360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraRegistered 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   7500
      TabIndex        =   235
      Top             =   30
      Visible         =   0   'False
      Width           =   2325
      Begin VB.Label lblReg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   60
         TabIndex        =   236
         Top             =   0
         Width           =   90
      End
   End
   Begin VB.CommandButton cmdDataHELP 
      Height          =   350
      Left            =   11640
      Picture         =   "frmMain.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   219
      TabStop         =   0   'False
      Top             =   4920
      Width           =   350
   End
   Begin VB.Frame frmBrowseToolbar 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   300
      TabIndex        =   83
      Top             =   6960
      Width           =   10695
      Begin VB.CommandButton cmdBrowseHelp 
         Height          =   350
         Left            =   1650
         Picture         =   "frmMain.frx":0356
         Style           =   1  'Graphical
         TabIndex        =   216
         TabStop         =   0   'False
         Top             =   480
         Width           =   350
      End
      Begin VB.Timer tmrNavigate 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   10320
         Top             =   0
      End
      Begin VB.CommandButton cmdBrowseRefresh 
         Height          =   350
         Left            =   1200
         Picture         =   "frmMain.frx":03A2
         Style           =   1  'Graphical
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   480
         Width           =   350
      End
      Begin VB.CommandButton cmdBrowseStop 
         Height          =   350
         Left            =   810
         Picture         =   "frmMain.frx":03F3
         Style           =   1  'Graphical
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   480
         Width           =   350
      End
      Begin VB.CommandButton cmdBrowseForward 
         Height          =   350
         Left            =   420
         Picture         =   "frmMain.frx":0441
         Style           =   1  'Graphical
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   480
         Width           =   350
      End
      Begin VB.CommandButton cmdBrowseBack 
         Height          =   350
         Left            =   30
         Picture         =   "frmMain.frx":0488
         Style           =   1  'Graphical
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   480
         Width           =   350
      End
      Begin VB.CommandButton cmdBrowseTo 
         Caption         =   "Go"
         Height          =   315
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   90
         Width           =   495
      End
      Begin VB.TextBox txtBrowserURL 
         Height          =   315
         Left            =   60
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   90
         Width           =   9195
      End
      Begin VB.Frame fraWebSubmit 
         Height          =   435
         Left            =   2130
         TabIndex        =   158
         Top             =   390
         Width           =   8115
         Begin VB.CommandButton cmdBrowseFill 
            Height          =   350
            Left            =   1290
            Picture         =   "frmMain.frx":04CF
            Style           =   1  'Graphical
            TabIndex        =   232
            TabStop         =   0   'False
            Top             =   90
            Width           =   350
         End
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   6630
            Picture         =   "frmMain.frx":0519
            ScaleHeight     =   240
            ScaleWidth      =   1350
            TabIndex        =   197
            Top             =   150
            Width           =   1350
         End
         Begin VB.CommandButton cmdBrowseListAdd 
            Height          =   350
            Left            =   2460
            Picture         =   "frmMain.frx":0636
            Style           =   1  'Graphical
            TabIndex        =   196
            TabStop         =   0   'False
            Top             =   90
            Width           =   350
         End
         Begin VB.CommandButton cmdBrowseListCut 
            Height          =   350
            Left            =   2070
            Picture         =   "frmMain.frx":067F
            Style           =   1  'Graphical
            TabIndex        =   166
            TabStop         =   0   'False
            Top             =   90
            Width           =   350
         End
         Begin VB.CommandButton cmdBrowseListUpdate 
            Height          =   350
            Left            =   1680
            Picture         =   "frmMain.frx":06D1
            Style           =   1  'Graphical
            TabIndex        =   165
            TabStop         =   0   'False
            Top             =   90
            Width           =   350
         End
         Begin VB.CheckBox chkWebList 
            Height          =   255
            Left            =   2880
            TabIndex        =   161
            TabStop         =   0   'False
            ToolTipText     =   "Mark this site as complete."
            Top             =   120
            Width           =   3465
         End
         Begin VB.CommandButton cmdBrowseListBack 
            Height          =   350
            Left            =   510
            Picture         =   "frmMain.frx":0724
            Style           =   1  'Graphical
            TabIndex        =   160
            TabStop         =   0   'False
            Top             =   90
            Width           =   350
         End
         Begin VB.CommandButton cmdBrowseListForward 
            Height          =   350
            Left            =   900
            Picture         =   "frmMain.frx":076B
            Style           =   1  'Graphical
            TabIndex        =   159
            TabStop         =   0   'False
            Top             =   90
            Width           =   350
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "List:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   162
            Top             =   150
            Width           =   375
         End
      End
      Begin VB.PictureBox pctNavigate 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   0
         Left            =   9870
         Picture         =   "frmMain.frx":07B2
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   163
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox pctNavigate 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   10170
         Picture         =   "frmMain.frx":0A34
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   164
         Top             =   270
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1515
      Left            =   930
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   7650
      Width           =   3315
      ExtentX         =   5847
      ExtentY         =   2672
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Frame frmEditButtons 
      Height          =   2445
      Left            =   11370
      TabIndex        =   76
      Top             =   1440
      Width           =   975
      Begin VB.CommandButton cmdSetDefault 
         Caption         =   "Default"
         Height          =   345
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   1980
         Width           =   795
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   345
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   1530
         Width           =   795
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load"
         Height          =   345
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   1080
         Width           =   795
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   345
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   630
         Width           =   795
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   345
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   180
         Width           =   795
      End
   End
   Begin VB.Frame frmToolbar 
      Height          =   435
      Left            =   270
      TabIndex        =   74
      Top             =   6210
      Width           =   10605
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   60
         Picture         =   "frmMain.frx":0CB6
         ScaleHeight     =   240
         ScaleWidth      =   1875
         TabIndex        =   88
         Top             =   150
         Width           =   1875
      End
      Begin VB.CommandButton cmdShowConfig 
         Caption         =   "Hide"
         Height          =   345
         Left            =   9810
         Style           =   1  'Graphical
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   90
         Width           =   765
      End
      Begin VB.Label lblStatus 
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   2580
         TabIndex        =   149
         Top             =   150
         Width           =   6375
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5265
      Left            =   810
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   630
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   9287
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   6
      TabsPerRow      =   7
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Program"
      TabPicture(0)   =   "frmMain.frx":0E0F
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdFindFile"
      Tab(0).Control(1)=   "cmdSelectLanguage"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSelectInstallSupport"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdSelectDistIncludes"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSelectCompany"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdSelectAuthor"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdSelectProgramType"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdSelectPermissions"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdSelectOSSupport"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdSelectBlurbs"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdViewExpirationInfo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtFileSize"
      Tab(0).Control(12)=   "txtFilenameLong"
      Tab(0).Control(13)=   "txtFilenameGeneric"
      Tab(0).Control(14)=   "txtFilenamePrevious"
      Tab(0).Control(15)=   "txtFilenameVersioned"
      Tab(0).Control(16)=   "txtChangeInfo"
      Tab(0).Control(17)=   "txtCompany"
      Tab(0).Control(18)=   "txtAuthor"
      Tab(0).Control(19)=   "txtLanguage"
      Tab(0).Control(20)=   "txtInstallSupport"
      Tab(0).Control(21)=   "cmdSelectReleaseStatus"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtReleaseStatus"
      Tab(0).Control(23)=   "txtProgramType"
      Tab(0).Control(24)=   "txtPermissions"
      Tab(0).Control(25)=   "txtDistributionIncludes"
      Tab(0).Control(26)=   "txtSystemRequirements"
      Tab(0).Control(27)=   "txtRegistrationCostOther"
      Tab(0).Control(28)=   "txtOSSupport"
      Tab(0).Control(29)=   "txtBlurbs"
      Tab(0).Control(30)=   "txtProgramVersion"
      Tab(0).Control(31)=   "txtReleaseDate"
      Tab(0).Control(32)=   "txtRegistrationCostUSD"
      Tab(0).Control(33)=   "txtProductName"
      Tab(0).Control(34)=   "Frame2"
      Tab(0).Control(35)=   "Label64"
      Tab(0).Control(36)=   "Label63"
      Tab(0).Control(37)=   "Label62"
      Tab(0).Control(38)=   "Label61"
      Tab(0).Control(39)=   "Label60"
      Tab(0).Control(40)=   "Label59"
      Tab(0).Control(41)=   "Label58"
      Tab(0).Control(42)=   "Label49"
      Tab(0).Control(43)=   "Label48"
      Tab(0).Control(44)=   "Label47"
      Tab(0).Control(45)=   "Label46"
      Tab(0).Control(46)=   "Label36"
      Tab(0).Control(47)=   "Label35"
      Tab(0).Control(48)=   "Label40"
      Tab(0).Control(49)=   "Label39"
      Tab(0).Control(50)=   "Label37"
      Tab(0).Control(51)=   "Label34"
      Tab(0).Control(52)=   "Label33"
      Tab(0).Control(53)=   "Label32"
      Tab(0).Control(54)=   "Label31"
      Tab(0).Control(55)=   "Label30"
      Tab(0).Control(56)=   "Label29"
      Tab(0).ControlCount=   57
      TabCaption(1)   =   "Program (cont)"
      TabPicture(1)   =   "frmMain.frx":0E2B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label38"
      Tab(1).Control(1)=   "Label41"
      Tab(1).Control(2)=   "Label42"
      Tab(1).Control(3)=   "Label43"
      Tab(1).Control(4)=   "Label44"
      Tab(1).Control(5)=   "Label45"
      Tab(1).Control(6)=   "Label55"
      Tab(1).Control(7)=   "Label56"
      Tab(1).Control(8)=   "Label57"
      Tab(1).Control(9)=   "txtDownloadURL(0)"
      Tab(1).Control(10)=   "txtDownloadURL(1)"
      Tab(1).Control(11)=   "txtDownloadURL(2)"
      Tab(1).Control(12)=   "txtDownloadURL(3)"
      Tab(1).Control(13)=   "txtXMLfileURL"
      Tab(1).Control(14)=   "txtIconFileURL"
      Tab(1).Control(15)=   "txtScreenshotURL"
      Tab(1).Control(16)=   "txtOrderURL"
      Tab(1).Control(17)=   "txtInfoURL"
      Tab(1).Control(18)=   "fraProgramDescriptions"
      Tab(1).Control(19)=   "txtKeywords"
      Tab(1).Control(20)=   "txtSpecificCategory"
      Tab(1).Control(21)=   "cmdSelectSpecificCategory"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtCategories"
      Tab(1).ControlCount=   23
      TabCaption(2)   =   "Company"
      TabPicture(2)   =   "frmMain.frx":0E47
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtCompanyLogoURL"
      Tab(2).Control(1)=   "cmdSelectOrganizations"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtOrganizations"
      Tab(2).Control(3)=   "txtCompanyName"
      Tab(2).Control(4)=   "txtAddress1"
      Tab(2).Control(5)=   "txtAddress2"
      Tab(2).Control(6)=   "txtCityTown"
      Tab(2).Control(7)=   "txtStateProvince"
      Tab(2).Control(8)=   "txtZIPPostal"
      Tab(2).Control(9)=   "txtCountry"
      Tab(2).Control(10)=   "txtCompanyWebsiteURL"
      Tab(2).Control(11)=   "txtSupportEmail"
      Tab(2).Control(12)=   "txtContactPhone"
      Tab(2).Control(13)=   "txtCompanyAbout"
      Tab(2).Control(14)=   "txtSalesEmail"
      Tab(2).Control(15)=   "txtGeneralEmail"
      Tab(2).Control(16)=   "txtSalesPhone"
      Tab(2).Control(17)=   "txtSupportPhone"
      Tab(2).Control(18)=   "txtFaxPhone"
      Tab(2).Control(19)=   "Label75"
      Tab(2).Control(20)=   "Label68"
      Tab(2).Control(21)=   "Label1"
      Tab(2).Control(22)=   "Label2"
      Tab(2).Control(23)=   "Label3"
      Tab(2).Control(24)=   "Label4"
      Tab(2).Control(25)=   "Label5"
      Tab(2).Control(26)=   "Label6"
      Tab(2).Control(27)=   "Label7"
      Tab(2).Control(28)=   "Label8"
      Tab(2).Control(29)=   "Label10"
      Tab(2).Control(30)=   "Label12"
      Tab(2).Control(31)=   "Label13"
      Tab(2).Control(32)=   "Label14"
      Tab(2).Control(33)=   "Label15"
      Tab(2).Control(34)=   "Label16"
      Tab(2).Control(35)=   "Label18"
      Tab(2).Control(36)=   "Label9"
      Tab(2).ControlCount=   37
      TabCaption(3)   =   "Author"
      TabPicture(3)   =   "frmMain.frx":0E63
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label23"
      Tab(3).Control(1)=   "Label22"
      Tab(3).Control(2)=   "Label21"
      Tab(3).Control(3)=   "Label20"
      Tab(3).Control(4)=   "Label19"
      Tab(3).Control(5)=   "Label11"
      Tab(3).Control(6)=   "txtContactFName"
      Tab(3).Control(7)=   "txtContactEMail"
      Tab(3).Control(8)=   "txtContactLName"
      Tab(3).Control(9)=   "txtAuthorLName"
      Tab(3).Control(10)=   "txtAuthorEmail"
      Tab(3).Control(11)=   "txtAuthorFName"
      Tab(3).ControlCount=   12
      TabCaption(4)   =   "Permissions"
      TabPicture(4)   =   "frmMain.frx":0E7F
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label26"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label25"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label24"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "txtPermissionName"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "txtDistributionPermisions"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "txtEULA"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Blurbs"
      TabPicture(5)   =   "frmMain.frx":0E9B
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label28"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label27"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "txtBlurbText"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "txtBlurbName"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "cmdBlurbOrder"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "Submit"
      TabPicture(6)   =   "frmMain.frx":0EB7
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "fraEmail"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "fraWeb"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "fraGenDocs"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "cmdAutoUpdate"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "cmdSubmit"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "Frame4"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).ControlCount=   6
      Begin VB.CommandButton cmdFindFile 
         Caption         =   "..."
         Height          =   300
         Left            =   -66120
         TabIndex        =   237
         Top             =   4620
         Width           =   285
      End
      Begin VB.TextBox txtCompanyLogoURL 
         Height          =   300
         Left            =   -73200
         TabIndex        =   47
         Top             =   3960
         Width           =   2595
      End
      Begin VB.Frame Frame4 
         Caption         =   "Submission Type"
         Height          =   1005
         Left            =   8760
         TabIndex        =   228
         Top             =   570
         Width           =   1605
         Begin VB.OptionButton optSubmitType 
            Caption         =   "E-Mail"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   230
            TabStop         =   0   'False
            Top             =   270
            Width           =   1335
         End
         Begin VB.OptionButton optSubmitType 
            Caption         =   "Web"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   229
            TabStop         =   0   'False
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdSubmit 
         BackColor       =   &H00E0E0E0&
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6990
         Style           =   1  'Graphical
         TabIndex        =   218
         TabStop         =   0   'False
         Top             =   4680
         Width           =   1395
      End
      Begin VB.CommandButton cmdAutoUpdate 
         Caption         =   "Auto Update"
         Height          =   405
         Left            =   5490
         Style           =   1  'Graphical
         TabIndex        =   217
         TabStop         =   0   'False
         Top             =   4680
         Width           =   1395
      End
      Begin VB.CommandButton cmdSelectOrganizations 
         Height          =   300
         Left            =   -66090
         Picture         =   "frmMain.frx":0ED3
         Style           =   1  'Graphical
         TabIndex        =   202
         TabStop         =   0   'False
         Top             =   3120
         Width           =   345
      End
      Begin VB.TextBox txtOrganizations 
         Height          =   300
         Left            =   -68370
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   3120
         Width           =   2355
      End
      Begin VB.CommandButton cmdSelectLanguage 
         Height          =   300
         Left            =   -66180
         Picture         =   "frmMain.frx":0F46
         Style           =   1  'Graphical
         TabIndex        =   146
         TabStop         =   0   'False
         Top             =   1244
         Width           =   345
      End
      Begin VB.CommandButton cmdSelectInstallSupport 
         Height          =   300
         Left            =   -66180
         Picture         =   "frmMain.frx":1228
         Style           =   1  'Graphical
         TabIndex        =   145
         TabStop         =   0   'False
         Top             =   2038
         Width           =   345
      End
      Begin VB.CommandButton cmdSelectDistIncludes 
         Height          =   300
         Left            =   -66180
         Picture         =   "frmMain.frx":150A
         Style           =   1  'Graphical
         TabIndex        =   140
         TabStop         =   0   'False
         Top             =   1641
         Width           =   345
      End
      Begin VB.CommandButton cmdSelectCompany 
         Height          =   300
         Left            =   -71010
         Picture         =   "frmMain.frx":17EC
         Style           =   1  'Graphical
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   1680
         Width           =   345
      End
      Begin VB.CommandButton cmdSelectAuthor 
         Height          =   300
         Left            =   -71010
         Picture         =   "frmMain.frx":185F
         Style           =   1  'Graphical
         TabIndex        =   147
         TabStop         =   0   'False
         Top             =   2065
         Width           =   345
      End
      Begin VB.CommandButton cmdSelectProgramType 
         Height          =   300
         Left            =   -71010
         Picture         =   "frmMain.frx":1B41
         Style           =   1  'Graphical
         TabIndex        =   143
         TabStop         =   0   'False
         Top             =   2450
         Width           =   345
      End
      Begin VB.CommandButton cmdSelectPermissions 
         Height          =   300
         Left            =   -71010
         Picture         =   "frmMain.frx":1E23
         Style           =   1  'Graphical
         TabIndex        =   142
         TabStop         =   0   'False
         Top             =   2835
         Width           =   345
      End
      Begin VB.CommandButton cmdSelectOSSupport 
         Height          =   300
         Left            =   -71010
         Picture         =   "frmMain.frx":2105
         Style           =   1  'Graphical
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   3600
         Width           =   345
      End
      Begin VB.CommandButton cmdSelectBlurbs 
         Height          =   300
         Left            =   -71010
         Picture         =   "frmMain.frx":23E7
         Style           =   1  'Graphical
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   3220
         Width           =   345
      End
      Begin VB.CommandButton cmdBlurbOrder 
         Caption         =   "Order"
         Height          =   300
         Left            =   -67050
         Style           =   1  'Graphical
         TabIndex        =   199
         Top             =   600
         Width           =   1275
      End
      Begin VB.CommandButton cmdViewExpirationInfo 
         Height          =   300
         Left            =   -71010
         Picture         =   "frmMain.frx":26C9
         Style           =   1  'Graphical
         TabIndex        =   198
         TabStop         =   0   'False
         Top             =   1260
         Width           =   345
      End
      Begin VB.TextBox txtFileSize 
         Height          =   300
         Left            =   -68430
         TabIndex        =   21
         Top             =   4620
         Width           =   1755
      End
      Begin VB.TextBox txtFilenameLong 
         Height          =   300
         Left            =   -68430
         TabIndex        =   20
         Top             =   4223
         Width           =   2595
      End
      Begin VB.TextBox txtFilenameGeneric 
         Height          =   300
         Left            =   -68430
         MaxLength       =   10
         TabIndex        =   19
         Top             =   3826
         Width           =   2595
      End
      Begin VB.TextBox txtFilenamePrevious 
         Height          =   300
         Left            =   -68430
         MaxLength       =   12
         TabIndex        =   18
         Top             =   3429
         Width           =   2595
      End
      Begin VB.TextBox txtFilenameVersioned 
         Height          =   300
         Left            =   -68430
         MaxLength       =   12
         TabIndex        =   17
         Top             =   3032
         Width           =   2595
      End
      Begin VB.TextBox txtChangeInfo 
         Height          =   915
         Left            =   -73260
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   3990
         Width           =   2595
      End
      Begin VB.TextBox txtCategories 
         Height          =   300
         Left            =   -73470
         TabIndex        =   27
         Top             =   2550
         Width           =   3495
      End
      Begin VB.CommandButton cmdSelectSpecificCategory 
         Height          =   300
         Left            =   -65100
         Picture         =   "frmMain.frx":273C
         Style           =   1  'Graphical
         TabIndex        =   181
         TabStop         =   0   'False
         Top             =   2520
         Width           =   345
      End
      Begin VB.TextBox txtSpecificCategory 
         Height          =   300
         Left            =   -68250
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   2520
         Width           =   3165
      End
      Begin VB.Frame fraGenDocs 
         Caption         =   "Generate Docs"
         Height          =   2820
         Left            =   8760
         TabIndex        =   174
         ToolTipText     =   "Generate documentation based on program info."
         Top             =   1620
         Width           =   1605
         Begin VB.CommandButton cmdGenHTML 
            Caption         =   "HTML"
            Height          =   315
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   195
            TabStop         =   0   'False
            Top             =   1890
            Width           =   1395
         End
         Begin VB.CommandButton cmdGenPAD 
            Caption         =   "PAD"
            Height          =   315
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   180
            TabStop         =   0   'False
            Top             =   300
            Width           =   1395
         End
         Begin VB.CommandButton cmdGenDIZ 
            Caption         =   "DIZ"
            Height          =   315
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   179
            TabStop         =   0   'False
            Top             =   615
            Width           =   1395
         End
         Begin VB.CommandButton cmdGenReadMe 
            Caption         =   "ReadME"
            Height          =   315
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   178
            TabStop         =   0   'False
            Top             =   930
            Width           =   1395
         End
         Begin VB.CommandButton cmdGenAll 
            Caption         =   "ALL"
            Height          =   315
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   177
            TabStop         =   0   'False
            Top             =   2340
            Width           =   1395
         End
         Begin VB.CommandButton cmdGenLicense 
            Caption         =   "License"
            Height          =   315
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   176
            TabStop         =   0   'False
            Top             =   1245
            Width           =   1395
         End
         Begin VB.CommandButton cmdGenPressRelease 
            Caption         =   "Press"
            Height          =   315
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   175
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1395
         End
      End
      Begin VB.TextBox txtKeywords 
         Height          =   300
         Left            =   -68250
         MaxLength       =   250
         TabIndex        =   32
         Top             =   2160
         Width           =   3495
      End
      Begin VB.Frame fraProgramDescriptions 
         Caption         =   "Descriptions"
         Height          =   2325
         Left            =   -74940
         TabIndex        =   167
         Top             =   2880
         Width           =   10335
         Begin VB.TextBox txtDescription 
            Height          =   1515
            Index           =   4
            Left            =   5220
            MaxLength       =   2000
            MultiLine       =   -1  'True
            TabIndex        =   38
            Top             =   720
            Width           =   4245
         End
         Begin VB.TextBox txtDescription 
            Height          =   1125
            Index           =   3
            Left            =   90
            MaxLength       =   450
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   1110
            Width           =   4245
         End
         Begin VB.TextBox txtDescription 
            Height          =   495
            Index           =   2
            Left            =   5220
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   37
            Top             =   180
            Width           =   4245
         End
         Begin VB.TextBox txtDescription 
            Height          =   495
            Index           =   1
            Left            =   90
            MaxLength       =   80
            MultiLine       =   -1  'True
            TabIndex        =   35
            Top             =   570
            Width           =   4245
         End
         Begin VB.TextBox txtDescription 
            Height          =   345
            Index           =   0
            Left            =   90
            MaxLength       =   45
            TabIndex        =   34
            Top             =   180
            Width           =   4245
         End
         Begin VB.Label lblDescriptionRemaining 
            Caption         =   "2000"
            Height          =   225
            Index           =   4
            Left            =   9510
            TabIndex        =   172
            Top             =   720
            Width           =   435
         End
         Begin VB.Label lblDescriptionRemaining 
            Caption         =   "450"
            Height          =   225
            Index           =   3
            Left            =   4410
            TabIndex        =   171
            Top             =   1110
            Width           =   435
         End
         Begin VB.Label lblDescriptionRemaining 
            Caption         =   "250"
            Height          =   225
            Index           =   2
            Left            =   9510
            TabIndex        =   170
            Top             =   180
            Width           =   435
         End
         Begin VB.Label lblDescriptionRemaining 
            Caption         =   "80"
            Height          =   225
            Index           =   1
            Left            =   4410
            TabIndex        =   169
            Top             =   570
            Width           =   435
         End
         Begin VB.Label lblDescriptionRemaining 
            Caption         =   "45"
            Height          =   225
            Index           =   0
            Left            =   4410
            TabIndex        =   168
            Top             =   180
            Width           =   435
         End
      End
      Begin VB.TextBox txtCompany 
         Height          =   300
         Left            =   -73260
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1680
         Width           =   2355
      End
      Begin VB.TextBox txtAuthor 
         Height          =   300
         Left            =   -73260
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2065
         Width           =   2355
      End
      Begin VB.TextBox txtLanguage 
         Height          =   300
         Left            =   -68430
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1244
         Width           =   2355
      End
      Begin VB.TextBox txtInstallSupport 
         Height          =   300
         Left            =   -68430
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2038
         Width           =   2355
      End
      Begin VB.CommandButton cmdSelectReleaseStatus 
         Height          =   300
         Left            =   -71010
         Picture         =   "frmMain.frx":2A1E
         Style           =   1  'Graphical
         TabIndex        =   144
         TabStop         =   0   'False
         Top             =   840
         Width           =   345
      End
      Begin VB.TextBox txtReleaseStatus 
         Height          =   300
         Left            =   -72180
         TabIndex        =   2
         Top             =   840
         Width           =   1185
      End
      Begin VB.TextBox txtProgramType 
         Height          =   300
         Left            =   -73260
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2450
         Width           =   2355
      End
      Begin VB.TextBox txtPermissions 
         Height          =   300
         Left            =   -73260
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2835
         Width           =   2355
      End
      Begin VB.TextBox txtDistributionIncludes 
         Height          =   300
         Left            =   -68430
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1641
         Width           =   2355
      End
      Begin VB.TextBox txtSystemRequirements 
         Height          =   500
         Left            =   -68430
         MaxLength       =   80
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   2435
         Width           =   2595
      End
      Begin VB.TextBox txtRegistrationCostOther 
         Height          =   300
         Left            =   -66690
         TabIndex        =   12
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtOSSupport 
         Height          =   300
         Left            =   -73260
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   3600
         Width           =   2355
      End
      Begin VB.TextBox txtBlurbs 
         Height          =   300
         Left            =   -73260
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3220
         Width           =   2355
      End
      Begin VB.TextBox txtInfoURL 
         Height          =   300
         Left            =   -73470
         TabIndex        =   26
         Top             =   2160
         Width           =   3495
      End
      Begin VB.TextBox txtOrderURL 
         Height          =   300
         Left            =   -73470
         TabIndex        =   25
         Top             =   1770
         Width           =   3495
      End
      Begin VB.TextBox txtScreenshotURL 
         Height          =   300
         Left            =   -73470
         TabIndex        =   24
         Top             =   1380
         Width           =   3495
      End
      Begin VB.TextBox txtIconFileURL 
         Height          =   300
         Left            =   -73470
         TabIndex        =   23
         Top             =   990
         Width           =   3495
      End
      Begin VB.TextBox txtXMLfileURL 
         Height          =   300
         Left            =   -73470
         TabIndex        =   22
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txtDownloadURL 
         Height          =   300
         Index           =   3
         Left            =   -68250
         TabIndex        =   31
         Top             =   1770
         Width           =   3495
      End
      Begin VB.TextBox txtDownloadURL 
         Height          =   300
         Index           =   2
         Left            =   -68250
         TabIndex        =   30
         Top             =   1380
         Width           =   3495
      End
      Begin VB.TextBox txtDownloadURL 
         Height          =   300
         Index           =   1
         Left            =   -68250
         TabIndex        =   29
         Top             =   990
         Width           =   3495
      End
      Begin VB.TextBox txtDownloadURL 
         Height          =   300
         Index           =   0
         Left            =   -68250
         TabIndex        =   28
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txtProgramVersion 
         Height          =   300
         Left            =   -73260
         TabIndex        =   1
         Top             =   840
         Width           =   1035
      End
      Begin VB.TextBox txtReleaseDate 
         Height          =   300
         Left            =   -68430
         TabIndex        =   10
         Top             =   450
         Width           =   2595
      End
      Begin VB.TextBox txtCompanyName 
         Height          =   300
         Left            =   -73200
         TabIndex        =   39
         Top             =   600
         Width           =   2595
      End
      Begin VB.TextBox txtAddress1 
         Height          =   300
         Left            =   -73200
         TabIndex        =   40
         Top             =   1020
         Width           =   2595
      End
      Begin VB.TextBox txtAddress2 
         Height          =   300
         Left            =   -73200
         TabIndex        =   41
         Top             =   1440
         Width           =   2595
      End
      Begin VB.TextBox txtCityTown 
         Height          =   300
         Left            =   -73200
         TabIndex        =   42
         Top             =   1860
         Width           =   2595
      End
      Begin VB.TextBox txtStateProvince 
         Height          =   300
         Left            =   -73200
         TabIndex        =   43
         Top             =   2280
         Width           =   2595
      End
      Begin VB.TextBox txtZIPPostal 
         Height          =   300
         Left            =   -73200
         TabIndex        =   44
         Top             =   2700
         Width           =   2595
      End
      Begin VB.TextBox txtCountry 
         Height          =   300
         Left            =   -73200
         TabIndex        =   45
         Top             =   3120
         Width           =   2595
      End
      Begin VB.TextBox txtCompanyWebsiteURL 
         Height          =   300
         Left            =   -73200
         TabIndex        =   46
         Top             =   3540
         Width           =   2595
      End
      Begin VB.TextBox txtSupportEmail 
         Height          =   300
         Left            =   -68370
         TabIndex        =   52
         Top             =   2280
         Width           =   2595
      End
      Begin VB.TextBox txtContactPhone 
         Height          =   300
         Left            =   -68370
         TabIndex        =   49
         Top             =   1020
         Width           =   2595
      End
      Begin VB.TextBox txtCompanyAbout 
         Height          =   705
         Left            =   -73200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   56
         Top             =   4380
         Width           =   7425
      End
      Begin VB.TextBox txtSalesEmail 
         Height          =   300
         Left            =   -68370
         TabIndex        =   50
         Top             =   1440
         Width           =   2595
      End
      Begin VB.TextBox txtGeneralEmail 
         Height          =   300
         Left            =   -68370
         TabIndex        =   48
         Top             =   600
         Width           =   2595
      End
      Begin VB.TextBox txtSalesPhone 
         Height          =   300
         Left            =   -68370
         TabIndex        =   51
         Top             =   1860
         Width           =   2595
      End
      Begin VB.TextBox txtSupportPhone 
         Height          =   300
         Left            =   -68370
         TabIndex        =   53
         Top             =   2700
         Width           =   2595
      End
      Begin VB.TextBox txtFaxPhone 
         Height          =   300
         Left            =   -68370
         TabIndex        =   55
         Top             =   3540
         Width           =   2595
      End
      Begin VB.TextBox txtAuthorFName 
         Height          =   300
         Left            =   -71010
         TabIndex        =   57
         Top             =   1410
         Width           =   1155
      End
      Begin VB.TextBox txtAuthorEmail 
         Height          =   300
         Left            =   -71010
         TabIndex        =   59
         Top             =   1830
         Width           =   2595
      End
      Begin VB.TextBox txtAuthorLName 
         Height          =   300
         Left            =   -69780
         TabIndex        =   58
         Top             =   1410
         Width           =   1365
      End
      Begin VB.TextBox txtContactLName 
         Height          =   300
         Left            =   -69780
         TabIndex        =   61
         Top             =   3090
         Width           =   1365
      End
      Begin VB.TextBox txtContactEMail 
         Height          =   300
         Left            =   -71010
         TabIndex        =   62
         Top             =   3510
         Width           =   2595
      End
      Begin VB.TextBox txtContactFName 
         Height          =   300
         Left            =   -71010
         TabIndex        =   60
         Top             =   3090
         Width           =   1155
      End
      Begin VB.TextBox txtEULA 
         Height          =   3285
         Left            =   -73200
         MaxLength       =   60000
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   65
         Top             =   1440
         Width           =   7425
      End
      Begin VB.TextBox txtDistributionPermisions 
         Height          =   300
         Left            =   -73200
         TabIndex        =   64
         Top             =   1020
         Width           =   7365
      End
      Begin VB.TextBox txtPermissionName 
         Height          =   300
         Left            =   -73200
         TabIndex        =   63
         Top             =   600
         Width           =   2595
      End
      Begin VB.TextBox txtBlurbName 
         Height          =   300
         Left            =   -73200
         TabIndex        =   66
         Top             =   600
         Width           =   2595
      End
      Begin VB.TextBox txtBlurbText 
         Height          =   3705
         Left            =   -73200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   67
         Top             =   1020
         Width           =   7425
      End
      Begin VB.TextBox txtRegistrationCostUSD 
         Height          =   300
         Left            =   -68430
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtProductName 
         Height          =   300
         Left            =   -73260
         TabIndex        =   0
         Top             =   450
         Width           =   2595
      End
      Begin VB.Frame fraWeb 
         Height          =   4335
         Left            =   450
         TabIndex        =   153
         Top             =   1170
         Width           =   8415
         Begin VB.CheckBox chkDisablePopUps 
            Caption         =   "Disable pop-up windows."
            Height          =   255
            Left            =   5490
            TabIndex        =   234
            Top             =   2820
            Value           =   1  'Checked
            Width           =   2835
         End
         Begin VB.CommandButton cmdSuggestSite 
            Caption         =   "Suggest Site"
            Height          =   405
            Left            =   6480
            Style           =   1  'Graphical
            TabIndex        =   226
            TabStop         =   0   'False
            Top             =   480
            Width           =   1695
         End
         Begin VB.Frame Frame3 
            Caption         =   "Refresh Submit List"
            Height          =   1335
            Left            =   5340
            TabIndex        =   221
            Top             =   1410
            Width           =   2985
            Begin VB.CommandButton cmdRefreshAllSites 
               Caption         =   "All Sites"
               Height          =   285
               Left            =   1530
               Style           =   1  'Graphical
               TabIndex        =   225
               TabStop         =   0   'False
               Top             =   930
               Width           =   1335
            End
            Begin VB.CommandButton cmdRefreshCurrentSite 
               Caption         =   "Current Site"
               Height          =   285
               Left            =   150
               Style           =   1  'Graphical
               TabIndex        =   224
               TabStop         =   0   'False
               Top             =   930
               Width           =   1335
            End
            Begin VB.CheckBox chkRefreshCompletion 
               Caption         =   "Completion Flags"
               Height          =   285
               Left            =   150
               TabIndex        =   223
               Top             =   570
               Value           =   1  'Checked
               Width           =   2535
            End
            Begin VB.CheckBox chkRefreshURLs 
               Caption         =   "URLs"
               Height          =   285
               Left            =   150
               TabIndex        =   222
               Top             =   240
               Width           =   2325
            End
         End
         Begin VB.CheckBox chkIncompleteSites 
            Caption         =   "Do not submit to completed sites."
            Height          =   255
            Left            =   5490
            TabIndex        =   220
            Top             =   3120
            Value           =   1  'Checked
            Width           =   2835
         End
         Begin VB.CommandButton cmdSelectSiteList 
            Height          =   300
            Left            =   3990
            Picture         =   "frmMain.frx":2A91
            Style           =   1  'Graphical
            TabIndex        =   154
            TabStop         =   0   'False
            Top             =   570
            Width           =   345
         End
         Begin VB.TextBox txtSiteList 
            Height          =   2805
            Left            =   1140
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   72
            Top             =   570
            Width           =   2835
         End
         Begin VB.Label Label53 
            BackStyle       =   0  'Transparent
            Caption         =   "Web Submission"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   157
            Top             =   90
            Width           =   2205
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sites:"
            Height          =   255
            Left            =   480
            TabIndex        =   155
            Top             =   570
            Width           =   465
         End
      End
      Begin VB.Frame fraEmail 
         Height          =   3945
         Left            =   180
         TabIndex        =   150
         Top             =   570
         Width           =   8415
         Begin VB.CommandButton cmdSuggestEmail 
            Caption         =   "Suggest Email"
            Height          =   405
            Left            =   6480
            Style           =   1  'Graphical
            TabIndex        =   227
            TabStop         =   0   'False
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton cmdViewLog 
            Caption         =   "View Log"
            Height          =   405
            Left            =   6480
            Style           =   1  'Graphical
            TabIndex        =   208
            TabStop         =   0   'False
            Top             =   1050
            Width           =   1695
         End
         Begin VB.TextBox txtFromAddress 
            Height          =   300
            Left            =   1140
            TabIndex        =   206
            Top             =   1410
            Width           =   2895
         End
         Begin VB.TextBox txtFromName 
            Height          =   300
            Left            =   1140
            TabIndex        =   204
            Top             =   1050
            Width           =   2895
         End
         Begin VB.Frame Frame1 
            Caption         =   "Server Settings"
            Height          =   1785
            Left            =   180
            TabIndex        =   191
            Top             =   2070
            Width           =   8115
            Begin VB.TextBox txtEmailRetries 
               Height          =   300
               Left            =   960
               TabIndex        =   214
               Top             =   1350
               Width           =   2895
            End
            Begin VB.TextBox txtSMTPUser 
               Height          =   300
               Left            =   5070
               TabIndex        =   210
               Top             =   990
               Width           =   2895
            End
            Begin VB.TextBox txtSMTPPassword 
               Height          =   300
               Left            =   5070
               TabIndex        =   209
               Top             =   1350
               Width           =   2895
            End
            Begin VB.TextBox txtTimeOut 
               Height          =   300
               Left            =   960
               TabIndex        =   71
               Top             =   990
               Width           =   2895
            End
            Begin VB.TextBox txtSMTPPort 
               Height          =   300
               Left            =   960
               TabIndex        =   70
               Top             =   630
               Width           =   2895
            End
            Begin VB.TextBox txtMailserver 
               Height          =   300
               Left            =   960
               TabIndex        =   69
               Top             =   270
               Width           =   2895
            End
            Begin VB.Label Label74 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Retries:"
               Height          =   255
               Left            =   -120
               TabIndex        =   215
               Top             =   1350
               Width           =   975
            End
            Begin VB.Label Label73 
               Caption         =   "SMTP Authentication:"
               Height          =   255
               Left            =   4110
               TabIndex        =   213
               Top             =   690
               Width           =   2955
            End
            Begin VB.Label Label72 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "User Name:"
               Height          =   255
               Left            =   3990
               TabIndex        =   212
               Top             =   990
               Width           =   975
            End
            Begin VB.Label Label71 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Password:"
               Height          =   255
               Left            =   3990
               TabIndex        =   211
               Top             =   1350
               Width           =   975
            End
            Begin VB.Label Label67 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Timeout:"
               Height          =   255
               Left            =   -120
               TabIndex        =   194
               Top             =   990
               Width           =   975
            End
            Begin VB.Label Label66 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Port:"
               Height          =   255
               Left            =   -120
               TabIndex        =   193
               Top             =   630
               Width           =   975
            End
            Begin VB.Label Label65 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Address:"
               Height          =   255
               Left            =   -120
               TabIndex        =   192
               Top             =   270
               Width           =   975
            End
         End
         Begin VB.TextBox txtEmailAddress 
            Height          =   500
            Left            =   1140
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   68
            Top             =   510
            Width           =   2895
         End
         Begin VB.CommandButton cmdSelectEmailAddress 
            Height          =   300
            Left            =   4035
            Picture         =   "frmMain.frx":2D73
            Style           =   1  'Graphical
            TabIndex        =   151
            TabStop         =   0   'False
            Top             =   510
            Width           =   345
         End
         Begin VB.Label Label70 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "From E-mail:"
            Height          =   255
            Left            =   90
            TabIndex        =   207
            Top             =   1410
            Width           =   975
         End
         Begin VB.Label Label69 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "From Name:"
            Height          =   255
            Left            =   90
            TabIndex        =   205
            Top             =   1050
            Width           =   975
         End
         Begin VB.Label Label52 
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail Submission"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   0
            TabIndex        =   156
            Top             =   90
            Width           =   3375
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "To:"
            Height          =   255
            Left            =   630
            TabIndex        =   152
            Top             =   510
            Width           =   435
         End
      End
      Begin VB.Frame Frame2 
         Height          =   465
         Left            =   -73260
         TabIndex        =   200
         Top             =   1140
         Width           =   2625
         Begin VB.CheckBox chkProgramExpires 
            Caption         =   "Expiration Information"
            Height          =   195
            Left            =   120
            TabIndex        =   201
            Top             =   180
            Width           =   2205
         End
      End
      Begin VB.Label Label75 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Logo URL:"
         Height          =   255
         Left            =   -74850
         TabIndex        =   233
         Top             =   3960
         Width           =   1515
      End
      Begin VB.Label Label68 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Organization Membership:"
         Height          =   255
         Left            =   -70350
         TabIndex        =   203
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label Label64 
         Caption         =   "bytes"
         Height          =   285
         Left            =   -66600
         TabIndex        =   190
         Top             =   4680
         Width           =   405
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FileSize:"
         Height          =   255
         Left            =   -70470
         TabIndex        =   189
         Top             =   4620
         Width           =   1965
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Filename Long:"
         Height          =   255
         Left            =   -70470
         TabIndex        =   188
         Top             =   4223
         Width           =   1965
      End
      Begin VB.Label Label61 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Filename Generic(6.3):"
         Height          =   255
         Left            =   -70470
         TabIndex        =   187
         Top             =   3826
         Width           =   1965
      End
      Begin VB.Label Label60 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Filename Previous(8.3):"
         Height          =   255
         Left            =   -70470
         TabIndex        =   186
         Top             =   3429
         Width           =   1965
      End
      Begin VB.Label Label59 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Filename Versioned(8.3):"
         Height          =   255
         Left            =   -70470
         TabIndex        =   185
         Top             =   3032
         Width           =   1965
      End
      Begin VB.Label Label58 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Change Info:"
         Height          =   255
         Left            =   -74820
         TabIndex        =   184
         Top             =   3990
         Width           =   1515
      End
      Begin VB.Label Label57 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Categories:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   183
         Top             =   2550
         Width           =   1275
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Specific Category:"
         Height          =   255
         Left            =   -69900
         TabIndex        =   182
         Top             =   2520
         Width           =   1515
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Keywords:"
         Height          =   255
         Left            =   -69900
         TabIndex        =   173
         Top             =   2160
         Width           =   1515
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Includes:"
         Height          =   255
         Left            =   -70020
         TabIndex        =   141
         Top             =   1641
         Width           =   1515
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Requirements:"
         Height          =   255
         Left            =   -70020
         TabIndex        =   139
         Top             =   2435
         Width           =   1515
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(Other):"
         Height          =   255
         Left            =   -67380
         TabIndex        =   138
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "OS Support:"
         Height          =   255
         Left            =   -74820
         TabIndex        =   137
         Top             =   3600
         Width           =   1515
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Info URL:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   134
         Top             =   2160
         Width           =   1275
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Order URL:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   133
         Top             =   1770
         Width           =   1275
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Screenshot URL:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   132
         Top             =   1380
         Width           =   1275
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Icon File URL:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   131
         Top             =   990
         Width           =   1275
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "XML File URL:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   130
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Download URL(s):"
         Height          =   255
         Left            =   -69900
         TabIndex        =   129
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Version:"
         Height          =   255
         Left            =   -74220
         TabIndex        =   128
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Release Date:"
         Height          =   255
         Left            =   -70020
         TabIndex        =   127
         Top             =   450
         Width           =   1515
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   -74850
         TabIndex        =   126
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address1:"
         Height          =   255
         Left            =   -74850
         TabIndex        =   125
         Top             =   1020
         Width           =   1515
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address2:"
         Height          =   255
         Left            =   -74850
         TabIndex        =   124
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "City/Town:"
         Height          =   255
         Left            =   -74850
         TabIndex        =   123
         Top             =   1860
         Width           =   1515
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "State/Province:"
         Height          =   255
         Left            =   -74850
         TabIndex        =   122
         Top             =   2280
         Width           =   1515
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ZIP/Postal Code:"
         Height          =   255
         Left            =   -74850
         TabIndex        =   121
         Top             =   2700
         Width           =   1515
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Country:"
         Height          =   255
         Left            =   -74850
         TabIndex        =   120
         Top             =   3120
         Width           =   1515
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Website:"
         Height          =   255
         Left            =   -74850
         TabIndex        =   119
         Top             =   3540
         Width           =   1515
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Support Email:"
         Height          =   255
         Left            =   -70020
         TabIndex        =   118
         Top             =   2280
         Width           =   1515
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Phone:"
         Height          =   255
         Left            =   -70020
         TabIndex        =   117
         Top             =   1020
         Width           =   1515
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "About:"
         Height          =   255
         Left            =   -74850
         TabIndex        =   116
         Top             =   4380
         Width           =   1515
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Email:"
         Height          =   255
         Left            =   -69840
         TabIndex        =   115
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Phone:"
         Height          =   255
         Left            =   -69750
         TabIndex        =   114
         Top             =   1860
         Width           =   1245
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Support Phone:"
         Height          =   255
         Left            =   -70020
         TabIndex        =   113
         Top             =   2700
         Width           =   1515
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         Height          =   255
         Left            =   -70020
         TabIndex        =   112
         Top             =   3540
         Width           =   1515
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Email:"
         Height          =   255
         Left            =   -69960
         TabIndex        =   111
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   -72360
         TabIndex        =   110
         Top             =   1410
         Width           =   1215
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         Height          =   255
         Left            =   -72450
         TabIndex        =   109
         Top             =   1830
         Width           =   1305
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         Height          =   255
         Left            =   -72450
         TabIndex        =   108
         Top             =   3510
         Width           =   1305
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   -72360
         TabIndex        =   107
         Top             =   3090
         Width           =   1215
      End
      Begin VB.Label Label22 
         Caption         =   "Contact"
         Height          =   255
         Left            =   -72510
         TabIndex        =   106
         Top             =   2820
         Width           =   1335
      End
      Begin VB.Label Label23 
         Caption         =   "Author"
         Height          =   255
         Left            =   -72510
         TabIndex        =   105
         Top             =   1140
         Width           =   1335
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EULA:"
         Height          =   255
         Left            =   -74850
         TabIndex        =   104
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Distribution Perm.:"
         Height          =   195
         Left            =   -74610
         TabIndex        =   103
         Top             =   1020
         Width           =   1275
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   -74850
         TabIndex        =   102
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   -74850
         TabIndex        =   101
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Text:"
         Height          =   255
         Left            =   -74850
         TabIndex        =   100
         Top             =   1020
         Width           =   1515
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Language:"
         Height          =   255
         Left            =   -70020
         TabIndex        =   99
         Top             =   1244
         Width           =   1515
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Install Support:"
         Height          =   255
         Left            =   -70020
         TabIndex        =   98
         Top             =   2038
         Width           =   1515
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Registration(USD):"
         Height          =   255
         Left            =   -69840
         TabIndex        =   97
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Type:"
         Height          =   255
         Left            =   -74820
         TabIndex        =   96
         Top             =   2450
         Width           =   1515
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Blurbs:"
         Height          =   255
         Left            =   -74820
         TabIndex        =   95
         Top             =   3220
         Width           =   1515
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Permission:"
         Height          =   255
         Left            =   -74820
         TabIndex        =   94
         Top             =   2835
         Width           =   1515
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Author:"
         Height          =   255
         Left            =   -74820
         TabIndex        =   93
         Top             =   2065
         Width           =   1515
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Company:"
         Height          =   255
         Left            =   -74820
         TabIndex        =   92
         Top             =   1680
         Width           =   1515
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   -74820
         TabIndex        =   91
         Top             =   450
         Width           =   1515
      End
   End
   Begin VB.Label lblBrowserStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "lblBrowserStatus"
      Height          =   285
      Left            =   10
      TabIndex        =   231
      Top             =   9300
      Width           =   6315
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Predef As New INI

Private Const PredefFile = "selections.dat"

Dim ShowConfig As Boolean
Dim defProgram As String
Dim defCompany As String
Dim defAuthor As String
Dim defPermission As String
Dim defBlurb As String

Private Sub chkProgramExpires_Click()
    If chkProgramExpires.Value = 1 Then frmExpireInfo.Show
    ProgramModified = True
End Sub

Private Sub chkWebList_Click()
    XX = chkWebList.Value
    SubmitList(SubmitListIndex).Completed = IIf(XX = 1, True, False)
End Sub

Private Sub cmdAutoUpdate_Click()
    frmUpdate.Show vbModal, Me
End Sub

Private Sub cmdBrowseBack_Click()
    On Error Resume Next
    Me.SetFocus
    WebBrowser1.GoBack
End Sub

Private Sub cmdBlurbOrder_Click()
    Dim tmpBlurbOrder As String
    Dim BlurbOrderFile As String
    Dim OldOrder As String
    Dim eOldOrder() As String
    Dim BlurbDir As String
    Dim eBlurbDir() As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    Me.SetFocus
    
    'load old order, verify contents, add new files
    BlurbOrderFile = DataPath & "blurborder.dat"
    If Dir(BlurbOrderFile) > "" Then
        Open BlurbOrderFile For Input As #FileNum
            Line Input #FileNum, OldOrder
        Close #FileNum
    End If
    
    eOldOrder = Split(OldOrder, ",")
    OldOrder = ""
    BlurbDir = GetDirContents(DataPath, ".blurb")
    For X = 0 To UBound(eOldOrder)
        If InStr(1, BlurbDir, eOldOrder(X)) > 0 Then OldOrder = OldOrder & eOldOrder(X) & ","
    Next X
    
    eBlurbDir = Split(BlurbDir, ",")
    For X = 0 To UBound(eBlurbDir)
        If InStr(1, OldOrder, eBlurbDir(X)) = 0 Then OldOrder = OldOrder & eBlurbDir(X) & ","
    Next X
    
    tmpBlurbOrder = frmSelect.SelectVals(OldOrder, "", False, "Blurb Order", , , True)

    Open BlurbOrderFile For Output As #FileNum
        Print #FileNum, tmpBlurbOrder
    Close #FileNum
    
    BlurbOrder = tmpBlurbOrder
    Status "Blurb order saved."
End Sub

Private Sub cmdBrowseFill_Click()
    FillBrowserForm
End Sub

Private Sub cmdBrowseForward_Click()
    On Error Resume Next
    Me.SetFocus
    WebBrowser1.GoForward
End Sub

Private Sub cmdBrowseHelp_Click()
    frmMain.WebBrowser1.Navigate HelpPath & "INDEX.HTML"
End Sub

Private Sub cmdBrowseListAdd_Click()
    Me.SetFocus
'    If txtBrowserURL <> SubmitList(SubmitListIndex).URL Then Surf2ListSite (SubmitListIndex)
    frmPostIt.Caption = SubmitList(SubmitListIndex).Name
    frmPostIt.txtPostIt.Text = SubmitList(SubmitListIndex).Notes
    frmPostIt.Show
End Sub

Private Sub cmdBrowseListBack_Click()
    Me.SetFocus
    GoPrevSite
End Sub

Private Sub cmdBrowseListCut_Click()
    Me.SetFocus
    If chkWebList.Caption = "" Then Exit Sub
    If MsgBox("Cut '" & SubmitList(SubmitListIndex).Name & "' from the submission list?", vbYesNo, "Cut") = vbNo Then Exit Sub
    
    X = UBound(SubmitList)
    If X = 0 Or SubmitListIndex = 0 Then Exit Sub
    
    For y = SubmitListIndex To X - 1
        SubmitList(y) = SubmitList(y + 1)
    Next y
    ReDim Preserve SubmitList(X - 1)
    If SubmitListIndex > UBound(SubmitList) Then SubmitListIndex = UBound(SubmitList)
    
    For X = 0 To UBound(SubmitList)
        If SubmitList(X).Name > "" Then tmp = tmp & SubmitList(X).Name & vbCrLf
    Next X
    If Len(tmp) > 1 Then tmp = Left$(tmp, Len(tmp) - 2)
    txtSiteList = tmp
    
'    Surf2ListSite SubmitListIndex
    GoNextSite
End Sub

Private Sub cmdBrowseListForward_Click()
    Me.SetFocus
    GoNextSite
End Sub

Private Sub cmdBrowseListUpdate_Click()
    Me.SetFocus
    
    XX = MsgBox("Update URL for: " & SubmitList(SubmitListIndex).Name & vbCrLf & "to: " & frmMain.txtBrowserURL, vbYesNo, "Update URL")
    If XX = vbYes Then SubmitList(SubmitListIndex).URL = frmMain.txtBrowserURL
    Debug.Print SubmitList(SubmitListIndex).URL
End Sub

Private Sub cmdBrowseRefresh_Click()
    On Error Resume Next
    Me.SetFocus
    WebBrowser1.Refresh
End Sub

Private Sub cmdBrowseStop_Click()
    On Error Resume Next
    Me.SetFocus
    WebBrowser1.Stop
    tmrNavigate.Enabled = False
    pctNavigate(0).Visible = True
    pctNavigate(1).Visible = False
End Sub

Private Sub cmdBrowseTo_Click()
    Me.SetFocus
    On Error Resume Next
    Select Case Trim(UCase(txtBrowserURL))
        Case "ADON"
            Status "AD ON"
            lblStatus.BackColor = &HC0C0FF
            AdminMode = True
        Case "ADOFF"
            Status "AD OFF"
            lblStatus.BackColor = &H8000000F
            AdminMode = False
        Case Else
            WebBrowser1.Navigate txtBrowserURL
    End Select
End Sub

Private Sub cmdDataHELP_Click()
    Me.SetFocus
    If (Me.Height < SSTab1.Height * 2) Then Me.Height = SSTab1.Height * 2
    Select Case UCase(SSTab1.TabCaption(SSTab1.Tab))
        Case "PROGRAM"
            frmMain.WebBrowser1.Navigate HelpPath & "PROGRAM.HTM"
        Case "PROGRAM (CONT)"
            frmMain.WebBrowser1.Navigate HelpPath & "PROGRAMCONT.HTM"
        Case "COMPANY"
            frmMain.WebBrowser1.Navigate HelpPath & "COMPANY.HTM"
        Case "AUTHOR"
            frmMain.WebBrowser1.Navigate HelpPath & "AUTHOR.HTM"
        Case "PERMISSIONS"
            frmMain.WebBrowser1.Navigate HelpPath & "PERMISSIONS.HTM"
        Case "BLURBS"
            frmMain.WebBrowser1.Navigate HelpPath & "BLURBS.HTM"
        Case "SUBMIT"
            frmMain.WebBrowser1.Navigate HelpPath & "SUBMIT.HTM"
    End Select
End Sub

Private Sub cmdDelete_Click()
    Me.SetFocus
    Dim FileName As String
    Dim Extension As String
    Dim DelType As String
    Dim tmpList As String
    
    Select Case UCase(SSTab1.TabCaption(SSTab1.Tab))
        Case "PROGRAM"
            tmpList = GetDirContents(DataPath, ".program")
            DeleteDataFile frmSelect.SelectVals(tmpList, "", False, "Delete Program"), ".program"
            
        Case "COMPANY"
            tmpList = GetDirContents(DataPath, ".company")
            DeleteDataFile frmSelect.SelectVals(tmpList, "", False, "Delete Company"), ".company"
            
        Case "AUTHOR"
            tmpList = GetDirContents(DataPath, ".author")
            DeleteDataFile frmSelect.SelectVals(tmpList, "", False, "Delete Author"), ".author"

        Case "PERMISSIONS"
            tmpList = GetDirContents(DataPath, ".permission")
            DeleteDataFile frmSelect.SelectVals(tmpList, "", False, "Delete Permission"), ".permission"
            
        Case "BLURBS"
            tmpList = GetDirContents(DataPath, ".blurb")
            DeleteDataFile frmSelect.SelectVals(tmpList, "", False, "Delete Blurb"), ".blurb"
    End Select
End Sub

Private Sub cmdFindFile_Click()
    Dim tmpName As String
    
    Me.SetFocus
    tmpName = SelectFile
    If tmpName > "" Then txtFileSize = FileLen(tmpName)
End Sub

Private Sub cmdGenAll_Click()
    Me.SetFocus
    If CheckForUnsavedData = False Then Exit Sub
    cmdGenHTML_Click
    cmdGenLicense_Click
    cmdGenPAD_Click
    cmdGenPressRelease_Click
    cmdGenReadMe_Click
    cmdGenDIZ_Click
    Status "Documents Generated"
End Sub

Private Sub cmdGenDIZ_Click()
    Dim tmpContents As String
    Dim Lines() As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    On Error Resume Next
    Me.SetFocus
    Status "Generating FILE_ID.DIZ..."
    If CheckForUnsavedData = False Then Exit Sub
    
    SetPageTuples
    ParseTemplate "FILE_ID_DIZ.template", "FILE_ID.DIZ"
    Open OutputPath & "FILE_ID.DIZ" For Input As #FileNum
        Do Until EOF(FileNum)
            Line Input #FileNum, tmp
            If Trim(tmp) > "" Then tmpContents = tmpContents & Trim(tmp) & " "
        Loop
    Close #FileNum
    
    tmpContents = WordWrap(tmpContents, 45)
    Lines = Split(tmpContents, vbCrLf)
    
    Open OutputPath & "FILE_ID.DIZ" For Output As #FileNum
        Print #FileNum, tmpContents;
    Close #FileNum
    
    If UBound(Lines) > 9 Then
        Status "FILE_ID.DIZ created, DIZ length error, (Greater than 10 lines)"
        MsgBox "DIZ length error, (Greater than 10 lines)", vbOKOnly, "Generate DIZ"
    Else
        Status "FILE_ID.DIZ created."
    End If
End Sub

Private Sub cmdGenHTML_Click()
    Dim tmp As String
    Dim FileNum As Integer
    
    Me.SetFocus
    Status "Generating INDEX.HTML..."
    If CheckForUnsavedData = False Then Exit Sub
    
    SetPageTuples
    tmp = strParseTemplate("HTML.template")
    tmp = Txt2HTMLCodes(tmp)
    
    FileNum = FreeFile
    Open OutputPath & "INDEX.HTML" For Output As #FileNum
        Print #FileNum, tmp
    Close #FileNum
    
    Status "INDEX.HTML created."
End Sub

Private Sub cmdGenLicense_Click()
    Dim tmp As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    Me.SetFocus
    Status "Generating LICENSE.TXT..."
    If CheckForUnsavedData = False Then Exit Sub
    
    SetPageTuples
    tmp = strParseTemplateText(txtEULA.Text)
    Open OutputPath & "LICENSE.TXT" For Output As #FileNum
        Print #FileNum, tmp
    Close #FileNum
    
    Status "LICENSE.TXT created."
End Sub

Private Sub cmdGenPAD_Click()
    Dim tmp As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    Me.SetFocus
    Status "Generating PAD files..."
    If CheckForUnsavedData = False Then Exit Sub
    
    SetPageTuples
    
    For X = 0 To UBound(PageTuple)
        PageTuple(X).Value = Txt2XMLCodes(PageTuple(X).Value)
    Next X
    
    ParseTemplate "PAD_FILE_XML.template", "pad_file.xml"
    
    Status "PAD_FILE.XML created. Genrating PAD_FILE.HTM..."
    
    SetPageTuples
    
    For X = 0 To UBound(PageTuple)
        PageTuple(X).Value = Txt2HTMLCodes(PageTuple(X).Value)
    Next X
    
    tmp = strParseTemplate("PAD_FILE_HTM.template")
    Open OutputPath & "pad_file.htm" For Output As #FileNum
        Print #FileNum, tmp
    Close #FileNum
    
    Status "PAD files created."
End Sub

Private Sub cmdGenPressRelease_Click()
    Me.SetFocus
    Status "Generating press release..."
    If CheckForUnsavedData = False Then Exit Sub
    
    SetPageTuples
    ParseTemplate "PressRelease.template", "PRESSRELEASE.TXT"
    Status "Press release created."
End Sub

Private Sub cmdGenReadMe_Click()
    Me.SetFocus
    Status "Generating README.TXT..."
    If CheckForUnsavedData = False Then Exit Sub
    LoadBlurbOrder
    
    TOCandArticles txtBlurbs, BlurbOrder, DataPath, ".blurb"
    SetPageTuples
    ParseTemplate "README.template", "README.TXT"
    TOC = "": Articles = ""
    Status "README.TXT created."
End Sub

Private Sub cmdLoad_Click()
    Me.SetFocus
    Dim tmp As String
    Dim tmp2 As String
    Dim tmpFile As String
    Dim tmpList As String
    
    On Error Resume Next
    Select Case UCase(SSTab1.TabCaption(SSTab1.Tab))
        Case "PROGRAM"
            tmpList = GetDirContents(DataPath, ".program")
            LoadProgram frmSelect.SelectVals(tmpList, "", False, "Load Program")
            
        Case "COMPANY"
            tmpList = GetDirContents(DataPath, ".company")
            LoadCompany frmSelect.SelectVals(tmpList, "", False, "Load Company")
            
        Case "AUTHOR"
            tmpList = GetDirContents(DataPath, ".author")
            LoadAuthor frmSelect.SelectVals(tmpList, "", False, "Load Author")
            
        Case "PERMISSIONS"
            tmpList = GetDirContents(DataPath, ".permission")
            LoadPermission frmSelect.SelectVals(tmpList, "", False, "Load Permission")
            
        Case "BLURBS"
            tmpList = GetDirContents(DataPath, ".blurb")
            LoadBlurb frmSelect.SelectVals(tmpList, "", False, "Load Blurb")
                    
    End Select
End Sub

Private Sub cmdNew_Click()
    Me.SetFocus
    Dim tmp As String
    Dim tmp2 As String
    Dim tmpFile As String
    Dim tmpList As String
    
    On Error Resume Next
    Select Case UCase(SSTab1.TabCaption(SSTab1.Tab))
        Case "PROGRAM"
            LoadProgram "DEFAULT"
            Status "New Program"
        Case "COMPANY"
            LoadCompany "DEFAULT"
            Status "New Company"
        Case "AUTHOR"
            txtAuthorFName = ""
            txtAuthorLName = ""
            txtAuthorEmail = ""
            txtContactFName = ""
            txtContactLName = ""
            txtContactEMail = ""
        Case "PERMISSIONS"
            txtPermissionName = ""
            txtDistributionPermisions = ""
            txtEULA = ""
        Case "BLURBS"
            txtBlurbName = ""
            txtBlurbText = ""
    End Select
End Sub

Private Sub cmdRefreshAllSites_Click()
    Dim tmp As String
    
    If SubmitListIndex = 0 Then Exit Sub
    If chkRefreshURLs.Value = 1 Then
        XX = MsgBox("Revert URL for all sites to their original values?", vbYesNo, "Refresh URLs")
        If XX = vbYes Then
            For X = 0 To UBound(SubmitList)
                tmp = SubmitList(X).Name
                yy = FindSiteListSite(tmp)
                If yy > 0 Then
                    SubmitList(X).URL = SiteList(yy).URL
                    Surf2ListSite SubmitListIndex
                End If
            Next X
        End If
    End If
    
    If chkRefreshCompletion.Value = 1 Then
        XX = MsgBox("Clear ALL completions?", vbYesNo, "Refresh")
        If XX = vbYes Then
            For X = 0 To UBound(SubmitList)
                SubmitList(X).Completed = False
                chkWebList.Value = 0
            Next X
        End If
    End If
End Sub

Private Sub cmdRefreshCurrentSite_Click()
    Dim tmp As String
    
    If SubmitListIndex = 0 Then Exit Sub
    tmp = SubmitList(SubmitListIndex).Name
    If chkRefreshURLs.Value = 1 Then
        XX = MsgBox("Revert URL for '" & tmp & "' to its original value?", vbYesNo, "Refresh URL")
        If XX = vbYes Then
            yy = FindSiteListSite(tmp)
            If yy > 0 Then
                SubmitList(SubmitListIndex).URL = SiteList(yy).URL
                Surf2ListSite SubmitListIndex
            End If
        End If
    End If
    
    If chkRefreshCompletion.Value = 1 Then
        SubmitList(SubmitListIndex).Completed = False
        chkWebList.Value = 0
    End If
End Sub

Private Sub cmdSave_Click()
    Me.SetFocus
    Select Case UCase(SSTab1.TabCaption(SSTab1.Tab))
        Case "PROGRAM"
            If Trim(txtProductName.Text) = "" Then Exit Sub
            SaveProgram Trim(txtProductName.Text) & " v" & txtProgramVersion.Text
            
        Case "COMPANY"
            SaveCompany txtCompanyName.Text
            
        Case "AUTHOR"
            SaveAuthor Trim(txtAuthorFName.Text & " " & txtAuthorLName.Text)
            
        Case "PERMISSIONS"
            SavePermission txtPermissionName.Text
            
        Case "BLURBS"
            SaveBlurb txtBlurbName.Text
    End Select
End Sub

Private Sub cmdSelectAuthor_Click()
    Dim tmpAuthor As String
    
    tmpAuthor = GetDirContents(DataPath, ".author")
    txtAuthor = frmSelect.SelectVals(tmpAuthor, txtAuthor, False, "Program Author")
    LoadAuthor txtAuthor
End Sub

Private Sub cmdSelectBlurbs_Click()
    Dim tmpBlurbs As String
    
    tmpBlurbs = GetDirContents(DataPath, ".blurb")
    txtBlurbs = frmSelect.SelectVals(tmpBlurbs, txtBlurbs, True, "Blurbs")
End Sub

Private Sub cmdSelectCompany_Click()
    Dim tmpCompany As String
    
    tmpCompany = GetDirContents(DataPath, ".company")
    txtCompany = frmSelect.SelectVals(tmpCompany, txtCompany, False, "Company")
    LoadCompany txtCompany
End Sub

Private Sub cmdSelectDistIncludes_Click()
    txtDistributionIncludes = frmSelect.SelectVals(Predef.INIValue("Dist_Includes"), txtDistributionIncludes, True, "Distribution Includes")
End Sub

Private Sub cmdSelectEmailAddress_Click()
    Dim tmpAddress As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    If Dir(DataPath & "Maillist.dat") = "" Then Exit Sub
    Open DataPath & "Maillist.dat" For Input As #FileNum
        Do Until EOF(FileNum)
            Line Input #FileNum, X$
            If Trim(X$) > "" Then tmpAddress = tmpAddress & Trim(X$) & ";"
        Loop
    Close #FileNum
    
    If tmpAddress = "" Then Exit Sub
    txtEmailAddress = frmSelect.SelectVals(tmpAddress, txtEmailAddress, True, "Email (To:)", ";", ";")
End Sub

Private Sub cmdSelectInstallSupport_Click()
    txtInstallSupport = frmSelect.SelectVals(Predef.INIValue("Program_Install_Support"), txtInstallSupport, False, "Install Support")
End Sub

Private Sub cmdSelectLanguage_Click()
    txtLanguage = frmSelect.SelectVals(Predef.INIValue("Program_Language"), txtLanguage, True, "Language")
End Sub

Private Sub cmdSelectOrganizations_Click()
    txtOrganizations = frmSelect.SelectVals(Predef.INIValue("Organization_Membership"), txtOrganizations, True, "Program Type")
End Sub

Private Sub cmdSelectOSSupport_Click()
    txtOSSupport = frmSelect.SelectVals(Predef.INIValue("Program_OS_Support"), txtOSSupport, True, "OS Support")
End Sub

Private Sub cmdSelectPermissions_Click()
    Dim tmpPerm As String
    
    tmpPerm = GetDirContents(DataPath, ".permission")
    txtPermissions = frmSelect.SelectVals(tmpPerm, txtPermissions, False, "Permissions")
    LoadPermission txtPermissions
End Sub

Private Sub cmdSelectProgramType_Click()
    txtProgramType = frmSelect.SelectVals(Predef.INIValue("Program_Type"), txtProgramType, False, "Program Type")
End Sub

Private Sub cmdSelectReleaseStatus_Click()
    txtReleaseStatus = frmSelect.SelectVals(Predef.INIValue("Program_Release_Status"), txtReleaseStatus, False, "Release Status")
End Sub

Private Sub cmdSelectSiteList_Click()
    Dim tmpList As String
    
    WebSubmiting = False
    UpdateDisplay
    For y = 1 To UBound(SiteList)
        tmpList = tmpList & SiteList(y).Name & vbCrLf
    Next y
    
    txtSiteList = frmSelect.SelectVals(tmpList, txtSiteList, True, "Submission Sites", vbCrLf, vbCrLf)
End Sub

Private Sub cmdSelectSpecificCategory_Click()
    txtSpecificCategory = frmSelect.SelectVals(Predef.INIValue("Program_Specific_Category"), txtSpecificCategory, False, "Specific Category")
End Sub


Private Sub cmdSetDefault_Click()
    Me.SetFocus
    Select Case UCase(SSTab1.TabCaption(SSTab1.Tab))
        Case "PROGRAM"
            SaveProgram "DEFAULT"
        Case "COMPANY"
            SaveCompany "DEFAULT"
    End Select
End Sub

Private Sub cmdShowConfig_Click()
    Me.SetFocus
    ShowConfig = Not (ShowConfig)
    lblStatus = ""
    UpdateDisplay
End Sub

Public Sub UpdateDisplay()
    Dim ShowEditbuttons As Boolean
    
    SSTab1.Top = 100
    SSTab1.Left = 80
    cmdShowConfig.Caption = IIf(ShowConfig, "Hide", "Show")
    cmdDataHELP.Top = SSTab1.Height - cmdDataHELP.Height + SSTab1.Top - 150
    cmdDataHELP.Left = SSTab1.Width - cmdDataHELP.Width + SSTab1.Left - 150
    
    'do show/hide config-------------
    If ShowConfig = True Then
        frmEditButtons.Left = SSTab1.Left + SSTab1.Width - frmEditButtons.Width - 150
        frmEditButtons.Top = SSTab1.Top + 350

        frmToolbar.Top = SSTab1.Top + SSTab1.Height
        SSTab1.Visible = True
        cmdDataHELP.Visible = True
        ShowEditbuttons = True
        If DemoVersion = True Then fraRegistered.Visible = True
    Else
        frmToolbar.Top = 0
        SSTab1.Visible = False
        cmdDataHELP.Visible = False
        ShowEditbuttons = False
        fraRegistered.Visible = False
    End If
    
    If UBound(SubmitList) = 0 Then WebSubmiting = False
    If WebSubmiting = True And optSubmitType(1).Value = True Then
        fraWebSubmit.Visible = True
    Else
        fraWebSubmit.Visible = False
    End If
    
    Select Case UCase(SSTab1.TabCaption(SSTab1.Tab))
        Case "PROGRAM"
            cmdNew.Visible = True
            cmdSetDefault.Visible = True
        Case "PROGRAM (CONT)"
            ShowEditbuttons = False
        Case "COMPANY"
            cmdNew.Visible = True
            cmdSetDefault.Visible = True
        Case "AUTHOR"
'            cmdNew.Visible = False
            cmdSetDefault.Visible = False
        Case "PERMISSIONS"
'            cmdNew.Visible = False
            cmdSetDefault.Visible = False
        Case "BLURBS"
'            cmdNew.Visible = False
            cmdSetDefault.Visible = False
        Case "SUBMIT"
            fraEmail.Visible = optSubmitType(0).Value
            fraWeb.Visible = Not (optSubmitType(0).Value)
            ShowEditbuttons = False
            fraEmail.Top = 500
            fraEmail.Left = 250
            fraWeb.Top = fraEmail.Top
            fraWeb.Left = fraEmail.Left
            fraWeb.Width = fraEmail.Width
            fraWeb.Height = fraEmail.Height
        Case Else
            ShowEditbuttons = True
    End Select
    
    frmEditButtons.Visible = ShowEditbuttons
    
    Form_Resize
End Sub

Private Sub cmdSubmit_Click()
    Me.SetFocus
    If CheckForUnsavedData = False Then Exit Sub
    If optSubmitType(0).Value = True Then
        'email submit
        Me.Height = 1
        Me.Width = 1
        
        SetPageTuples
        tmp = strParseTemplate("email.template")
        frmPreviewEmail.txtMailMessage = tmp
        frmPreviewEmail.Show 'vbModal, Me
    Else
        'web submit
'        If (Me.Height < SSTab1.Height * 2) Then Me.Height = SSTab1.Height * 2
        ShowConfig = False
        WebSubmiting = True
        UpdateDisplay
        Me.Refresh
        InitSubmitList
        GoNextSite
        'Surf2ListSite 1
    End If
        
    UpdateDisplay
End Sub

Private Sub cmdSuggestEmail_Click()
    Me.SetFocus
    XX = ShellExecute(hwnd, vbNullString, "http://www.aesgard.com/product/softauthor/suggestmail.htm", vbNullString, App.Path, 1)
End Sub

Private Sub cmdSuggestSite_Click()
    Me.SetFocus
    XX = ShellExecute(hwnd, vbNullString, "http://www.aesgard.com/product/softauthor/suggestsite.htm", vbNullString, App.Path, 1)
End Sub

Private Sub cmdViewExpirationInfo_Click()
    Me.SetFocus
    frmExpireInfo.Show
End Sub

Private Sub cmdViewLog_Click()
    Dim XX As Long
    
    If Dir(App.Path & "\EMAILSUBMIT.TXT") > "" Then
        XX = ShellExecute(hwnd, vbNullString, "EMAILSUBMIT.TXT", vbNullString, App.Path, 1)
    Else
        MsgBox "Log not found."
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    ReDim SiteList(0)
    ReDim SubmitList(0)
    Me.Width = 0
    Me.Height = 0
    ShowConfig = True
    
    DataPath = App.Path & "\DATA\"
    MkDir DataPath
    TemplatePath = App.Path & "\TEMPLATE\"
    MkDir TemplatePath
    OutputPath = App.Path & "\OUTPUT\"
    MkDir OutputPath
    HelpPath = App.Path & "\HELP\"
    MkDir HelpPath
    
    optSubmitType(1).Value = True
    SSTab1.Tab = 0
    Predef.ReadINI DataPath & PredefFile
    
    '...load defaults and .cfg
    LoadSiteList
    Debug.Print 1
    LoadMainCfg
    Debug.Print 2
    
    Me.Caption = App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision
    Debug.Print 3
    
    Do2000Look Me
    Debug.Print 4
    WebBrowser1.Silent = True
    WebBrowser1.Navigate HelpPath & "INDEX.HTML"
    UpdateDisplay
    Status ""
End Sub

Private Sub Form_QueryUnload(CANCEL As Integer, UnloadMode As Integer)
    SaveMainCfg
    CheckModifiedData
    For Each Form In Forms
        Unload Form
    Next
    End
End Sub

Private Sub Form_Resize()
    Dim MinX As Long
    Dim MinY As Long
    Dim tmpYBrowser As Long
    
    If Me.WindowState = 1 Then Exit Sub
    
    MinX = SSTab1.Width + 250
    MinY = IIf(ShowConfig, SSTab1.Height + frmToolbar.Height + 550, frmToolbar.Height + 550)
   
    If Me.Width < MinX Then Me.Width = MinX
    If Me.Height < MinY Then Me.Height = MinY
    
    'toolbar movements---------------
    frmToolbar.Width = Me.Width - 200
    frmToolbar.Left = 50
    cmdShowConfig.Left = frmToolbar.Width - cmdShowConfig.Width ' - 100

    'do browser-------------------
    frmBrowseToolbar.Top = frmToolbar.Top + frmToolbar.Height
    frmBrowseToolbar.Width = Me.Width
    txtBrowserURL.Width = Me.Width - cmdBrowseTo.Width - pctNavigate(0).Width - 400
    frmBrowseToolbar.Left = 0
    cmdBrowseTo.Left = txtBrowserURL.Width + 150
    pctNavigate(0).Left = cmdBrowseTo.Left + cmdBrowseTo.Width + 50
    WebBrowser1.Width = Me.Width - 120
    WebBrowser1.Top = frmBrowseToolbar.Top + frmBrowseToolbar.Height
    tmpYBrowser = Me.Height - WebBrowser1.Top - 450 - lblBrowserStatus.Height
    If tmpYBrowser < 1 Then tmpYBrowser = 0
    WebBrowser1.Height = tmpYBrowser
    WebBrowser1.Left = 0
    lblBrowserStatus.Top = WebBrowser1.Top + WebBrowser1.Height + 100
    lblBrowserStatus.Width = WebBrowser1.Width
    fraWebSubmit.Width = Me.Width - fraWebSubmit.Left - 200
    Picture2.Left = fraWebSubmit.Width - Picture2.Width - 100
    chkWebList.Width = fraWebSubmit.Width - chkWebList.Left - Picture2.Width - 100
End Sub

Private Sub optSubmitType_Click(Index As Integer)
    UpdateDisplay
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    UpdateDisplay
End Sub

Private Sub tmrNavigate_Timer()
    pctNavigate(1).Top = pctNavigate(0).Top
    pctNavigate(1).Left = pctNavigate(0).Left
    X = IIf(pctNavigate(0).Visible = True, 1, 0)
    y = IIf(X = 1, 0, 1)
    
    pctNavigate(X).Visible = True
    pctNavigate(y).Visible = False
End Sub

Private Sub txtAddress1_Change()
    CompanyModified = True
End Sub

Private Sub txtAddress2_Change()
    CompanyModified = True
End Sub

Private Sub txtAuthor_Change()
    ProgramModified = True
End Sub

Private Sub txtAuthorEmail_Change()
    AuthorModified = True
End Sub

Private Sub txtAuthorFName_Change()
    AuthorModified = True
End Sub

Private Sub txtAuthorLName_Change()
    AuthorModified = True
End Sub

Private Sub txtBlurbName_Change()
    BlurbModified = True
End Sub

Private Sub txtBlurbs_Change()
    ProgramModified = True
End Sub

Private Sub txtBlurbText_Change()
    BlurbModified = True
End Sub

Private Sub txtBrowserURL_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdBrowseTo_Click
End Sub

Private Sub txtCategories_Change()
    ProgramModified = True
End Sub

Private Sub txtChangeInfo_Change()
    ProgramModified = True
End Sub

Private Sub txtCityTown_Change()
    CompanyModified = True
End Sub

Private Sub txtCompany_Change()
    ProgramModified = True
End Sub

Private Sub txtCompanyAbout_Change()
    CompanyModified = True
End Sub

Private Sub txtCompanyLogoURL_Change()
    CompanyModified = True
End Sub

Private Sub txtCompanyName_Change()
    CompanyModified = True
End Sub

Private Sub txtCompanyWebsiteURL_Change()
    CompanyModified = True
End Sub

Private Sub txtContactEMail_Change()
    AuthorModified = True
End Sub

Private Sub txtContactFName_Change()
    AuthorModified = True
End Sub

Private Sub txtContactLName_Change()
    AuthorModified = True
End Sub

Private Sub txtContactPhone_Change()
    CompanyModified = True
End Sub

Private Sub txtCountry_Change()
    CompanyModified = True
End Sub

Private Sub txtDescription_Change(Index As Integer)
    ProgramModified = True
    lblDescriptionRemaining(Index) = txtDescription(Index).MaxLength - Len(txtDescription(Index).Text)
End Sub

Private Sub txtDistributionIncludes_Change()
    ProgramModified = True
End Sub

Private Sub txtDistributionPermisions_Change()
    PermissionsModified = True
End Sub

Private Sub txtDownloadURL_Change(Index As Integer)
    ProgramModified = True
End Sub

Private Sub txtEmailAddress_Change()
    SubmitModified = True
End Sub

Private Sub txtEmailRetries_Change()
    SubmitModified = True
End Sub

Private Sub txtEmailRetries_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumbers(KeyAscii, True)
End Sub

Private Sub txtEULA_Change()
    PermissionsModified = True
End Sub

Private Sub txtFaxPhone_Change()
    CompanyModified = True
End Sub

Private Sub txtFilenameGeneric_Change()
    ProgramModified = True
End Sub

Private Sub txtFilenameLong_Change()
    ProgramModified = True
End Sub

Private Sub txtFilenamePrevious_Change()
    ProgramModified = True
End Sub

Private Sub txtFilenameVersioned_Change()
    ProgramModified = True
End Sub

Private Sub txtFileSize_Change()
    ProgramModified = True
End Sub

Private Sub txtFileSize_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumbers(KeyAscii, True)
End Sub

Private Sub txtFromAddress_Change()
    SubmitModified = True
End Sub

Private Sub txtFromName_Change()
    SubmitModified = True
End Sub

Private Sub txtGeneralEmail_Change()
    CompanyModified = True
End Sub

Private Sub txtIconFileURL_Change()
    ProgramModified = True
End Sub

Private Sub txtInfoURL_Change()
    ProgramModified = True
End Sub

Private Sub txtInstallSupport_Change()
    ProgramModified = True
End Sub

Private Sub txtKeywords_Change()
    ProgramModified = True
End Sub

Private Sub txtLanguage_Change()
    ProgramModified = True
End Sub

Private Sub txtMailserver_Change()
    SubmitModified = True
End Sub

Private Sub txtOrderURL_Change()
    ProgramModified = True
End Sub

Private Sub txtOrganizations_Change()
    CompanyModified = True
End Sub

Private Sub txtOSSupport_Change()
    ProgramModified = True
End Sub

Private Sub txtPermissionName_Change()
    PermissionsModified = True
End Sub

Private Sub txtPermissions_Change()
    ProgramModified = True
End Sub

Private Sub txtProductName_Change()
    ProgramModified = True
End Sub

Private Sub txtProgramType_Change()
    ProgramModified = True
End Sub

Private Sub txtProgramVersion_Change()
    ProgramModified = True
End Sub

Private Sub txtRegistrationCostOther_Change()
    ProgramModified = True
End Sub

Private Sub txtRegistrationCostUSD_Change()
    ProgramModified = True
End Sub

Private Sub txtRegistrationCostUSD_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumbers(KeyAscii, True)
End Sub

Private Sub txtReleaseDate_Change()
    ProgramModified = True
End Sub

Private Sub txtReleaseStatus_Change()
    ProgramModified = True
End Sub

Private Sub txtSalesEmail_Change()
    CompanyModified = True
End Sub

Private Sub txtSalesPhone_Change()
    CompanyModified = True
End Sub

Private Sub txtScreenshotURL_Change()
    ProgramModified = True
End Sub

Private Sub txtSiteList_Change()
    SubmitModified = True
End Sub

Private Sub txtSMTPPassword_Change()
    SubmitModified = True
End Sub

Private Sub txtSMTPPort_Change()
    SubmitModified = True
End Sub

Private Sub txtSMTPPort_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumbers(KeyAscii, True)
End Sub

Private Sub txtSMTPUser_Change()
    SubmitModified = True
End Sub

Private Sub txtSpecificCategory_Change()
    ProgramModified = True
End Sub

Private Sub txtStateProvince_Change()
    CompanyModified = True
End Sub

Private Sub txtSupportEmail_Change()
    CompanyModified = True
End Sub

Private Sub txtSupportPhone_Change()
    CompanyModified = True
End Sub

Private Sub txtSystemRequirements_Change()
    ProgramModified = True
End Sub

Private Sub txtTimeOut_Change()
    SubmitModified = True
End Sub

Private Sub txtTimeOut_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumbers(KeyAscii, True)
End Sub

Private Sub txtXMLfileURL_Change()
    ProgramModified = True
End Sub

Private Sub txtZIPPostal_Change()
    CompanyModified = True
End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, CANCEL As Boolean)
    Dim eHeaders() As String
    Dim tmpHeaders As Variant
    
    On Error Resume Next
    eHeaders = Split(Headers, vbCrLf)
    For X = 0 To UBound(eHeaders)
        If InStr(1, eHeaders(X), "Referer: ") > 0 Then
            eHeaders(X) = "Referer: http://www.aesgard.com/"
        End If
        tmpHeaders = tmpHeaders & eHeaders(X) & vbCrLf
    Next X
    Headers = tmpHeaders
    
'    txtBrowserURL = WebBrowser1.LocationURL
    tmrNavigate.Enabled = True
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    txtBrowserURL = WebBrowser1.LocationURL
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    tmrNavigate.Enabled = False
    pctNavigate(0).Visible = True
    pctNavigate(1).Visible = False
    txtBrowserURL = WebBrowser1.LocationURL
End Sub

Private Sub SaveMainCfg()
    Dim FileNum As Integer
    
    Status "Saving current configuration data..."
    FileNum = FreeFile
'    On Error Resume Next
    Open DataPath & "main.dat" For Output As #FileNum
        Print #FileNum, Trim(txtProductName.Text) & " v" & txtProgramVersion.Text
        Print #FileNum, txtCompanyName.Text
        Print #FileNum, Trim(txtAuthorFName.Text & " " & txtAuthorLName.Text)
        Print #FileNum, txtPermissionName.Text
        Print #FileNum, txtBlurbName.Text
        Print #FileNum, PostitLeft
        Print #FileNum, PostitTop
        Print #FileNum, PostitHeight
        Print #FileNum, PostitWidth
        Print #FileNum, LastUpdateNo
    Close #FileNum
    
    Status "Saving current e-mail data..."
    SaveSMTPcfg
    Status "Data saved."
    
    Status "Saving current submit data..."
    SaveSubmitList
    Status "Data saved."
    
    If AdminMode = True Then
        Status "Saving current site data..."
        SaveSiteList
        Status "Data saved."
    End If
    
    Status "Saving current sticky notes data..."
    SaveSubmitListNotes
    Status "Data saved."
End Sub

Private Sub LoadMainCfg()
    Dim tmp As String
    Dim FileNum As Integer
    Dim ProgName As String
    Dim CompName As String
    Dim AuthName As String
    Dim PermName As String
    Dim BlurbName As String
    
    FileNum = FreeFile
    
    If Dir(DataPath & "main.dat") > "" Then
        Open DataPath & "main.dat" For Input As #FileNum
            Line Input #FileNum, ProgName
            Line Input #FileNum, CompName
            Line Input #FileNum, AuthName
            Line Input #FileNum, PermName
            Line Input #FileNum, BlurbName
            Line Input #FileNum, tmp
                PostitLeft = CLng(tmp)
            Line Input #FileNum, tmp
                PostitTop = CLng(tmp)
            Line Input #FileNum, tmp
                PostitHeight = CLng(tmp)
            Line Input #FileNum, tmp
                PostitWidth = CLng(tmp)
            Line Input #FileNum, tmp
                LastUpdateNo = CInt(tmp)
        Close #FileNum
    End If
    
    Debug.Print 10
    LoadProgram ProgName
    Debug.Print 11
    LoadCompany CompName
    Debug.Print 12
    LoadAuthor AuthName
    Debug.Print 13
    LoadPermission PermName
    Debug.Print 14
    LoadBlurb BlurbName
    Debug.Print 15

    LoadSMTPcfg
    Debug.Print 16
    LoadSubmitList
    Debug.Print 17
    LoadSubmitListNotes
    Debug.Print 18
    SetTooltips Me
    Debug.Print 19
End Sub

Private Sub SaveSMTPcfg()
    Dim FileNum As Integer
    
    FileNum = FreeFile
    On Error Resume Next
    Open DataPath & "SMTP.dat" For Output As #FileNum
        Print #FileNum, frmMain.txtEmailAddress
        Print #FileNum, frmMain.txtFromName
        Print #FileNum, frmMain.txtFromAddress
        Print #FileNum, frmMain.txtMailserver
        Print #FileNum, frmMain.txtSMTPPort
        Print #FileNum, frmMain.txtTimeOut
        Print #FileNum, frmMain.txtEmailRetries
        Print #FileNum, frmMain.txtSMTPUser
        Print #FileNum, frmMain.txtSMTPPassword
    Close #FileNum
End Sub

Private Sub LoadSMTPcfg()
    Dim tmpAddress As String
    Dim etmp() As String
    Dim y As Integer
    Dim FileNum As Integer
    
    FileNum = FreeFile
    On Error Resume Next
   
    'open mailist.dat to build confirmation list for items in SMTP.dat
    tmpAddress = ";"
    If Dir(DataPath & "Maillist.dat") > "" Then
        Open DataPath & "Maillist.dat" For Input As #FileNum
            Do Until EOF(FileNum)
                Line Input #FileNum, X$
                If Trim(X$) > "" Then tmpAddress = tmpAddress & Trim(X$) & ";"
            Loop
        Close #FileNum
    End If
    
    If Dir(DataPath & "SMTP.dat") > "" Then
        Open DataPath & "SMTP.dat" For Input As #FileNum
            Line Input #FileNum, tmp
                etmp = Split(tmp, ";")
                tmp = ""
                For y = 0 To UBound(etmp)
                    If etmp(y) > "" And InStr(1, tmpAddress, ";" & etmp(y) & ";") > 0 Then tmp = tmp & etmp(y) & ";"
                Next y
                
                frmMain.txtEmailAddress = tmp
            Line Input #FileNum, tmp
                frmMain.txtFromName = tmp
            Line Input #FileNum, tmp
                frmMain.txtFromAddress = tmp
            Line Input #FileNum, tmp
                frmMain.txtMailserver = tmp
            Line Input #FileNum, tmp
                frmMain.txtSMTPPort = Val(tmp)
            Line Input #FileNum, tmp
                frmMain.txtTimeOut = Val(tmp)
            Line Input #FileNum, tmp
                frmMain.txtEmailRetries = Val(tmp)
            Line Input #FileNum, tmp
                frmMain.txtSMTPUser = tmp
            Line Input #FileNum, tmp
                frmMain.txtSMTPPassword = tmp
        Close #FileNum
    End If
    
    If Val(frmMain.txtSMTPPort) = 0 Then frmMain.txtSMTPPort = "25"
    If Val(frmMain.txtEmailRetries) = 0 Then frmMain.txtEmailRetries = "3"
    If Val(frmMain.txtTimeOut) = 0 Then frmMain.txtTimeOut = "30"
    
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, CANCEL As Boolean)
    Dim frmWB As frmNewBrowseWindow
    
    On Error Resume Next
    If chkDisablePopUps.Value = 1 Then
        Set frmWB = New frmNewBrowseWindow
        frmWB.WebBrowser1.RegisterAsBrowser = True
        Set ppDisp = frmWB.WebBrowser1.Object
'        frmWB.Visible = True
    End If
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
     lblBrowserStatus = Text
End Sub
