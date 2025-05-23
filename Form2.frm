VERSION 5.00
Object = "{BCE37951-37DF-4D69-A8A3-2CFABEE7B3CC}#1.0#0"; "CCRSlider.ocx"
Begin VB.Form FireCallPrefs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fire Call Win Preferences"
   ClientHeight    =   10650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10650
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraHousekeeping 
      Caption         =   "Housekeeping"
      Height          =   7755
      Left            =   690
      TabIndex        =   184
      Top             =   1215
      Width           =   8655
      Begin VB.Frame fraHiddenHousekeeping 
         BorderStyle     =   0  'None
         Height          =   7230
         Left            =   600
         TabIndex        =   185
         Top             =   255
         Width           =   7245
         Begin VB.CommandButton btnArchiveNow 
            Caption         =   "Archive Now"
            Height          =   525
            Left            =   5205
            TabIndex        =   285
            Top             =   2040
            Width           =   1320
         End
         Begin VB.Frame fraHouseKeepingBackups 
            BorderStyle     =   0  'None
            Height          =   4125
            Left            =   165
            TabIndex        =   253
            Top             =   3015
            Width           =   6855
            Begin VB.CommandButton btnBackupNow 
               Caption         =   "Backup Now"
               Height          =   525
               Left            =   5145
               TabIndex        =   284
               Top             =   3285
               Width           =   1320
            End
            Begin VB.CheckBox chkAutomaticBackups 
               Caption         =   "  Automatic Backups"
               Height          =   225
               Left            =   1800
               TabIndex        =   256
               ToolTipText     =   "Check this box to enable advice messages. If enabled, advice messages are sent periodically to this address."
               Top             =   1185
               Width           =   1950
            End
            Begin VB.CommandButton btnBackupLocation 
               Caption         =   "..."
               Height          =   300
               Left            =   1800
               Style           =   1  'Graphical
               TabIndex        =   255
               ToolTipText     =   "Open a file explorer at the Backup folder location."
               Top             =   3345
               Width           =   315
            End
            Begin VB.CheckBox chkBackupOnStart 
               Caption         =   "  Backup on Start"
               Height          =   225
               Left            =   1785
               TabIndex        =   254
               ToolTipText     =   "Check this box to enable advice messages. If enabled, advice messages are sent periodically to this address."
               Top             =   315
               Width           =   1725
            End
            Begin vb6projectCCRSlider.Slider sliAutomaticBackupInterval 
               Height          =   390
               Left            =   1680
               TabIndex        =   257
               ToolTipText     =   "Set the hourly interval "
               Top             =   1935
               Width           =   3870
               _ExtentX        =   6826
               _ExtentY        =   688
               Min             =   1
               Max             =   24
               Value           =   24
               SelStart        =   20
            End
            Begin VB.Label lblHousekeepingDesc 
               Caption         =   "Check this box to enable automatic hourly backups"
               Height          =   450
               Index           =   3
               Left            =   1800
               TabIndex        =   266
               Tag             =   "lblAutomaticBackupsDesc"
               ToolTipText     =   "Check this box to enable automatic hourly backups"
               Top             =   1530
               Width           =   4335
            End
            Begin VB.Label lblHousekeepingDesc 
               Caption         =   "Set the automatic backup interval in hours."
               Height          =   330
               Index           =   4
               Left            =   1770
               TabIndex        =   265
               Tag             =   "lblIntervalDesc"
               Top             =   2760
               Width           =   3810
            End
            Begin VB.Label lblHousekeepingTab 
               Caption         =   "Interval:"
               Height          =   315
               Index           =   1
               Left            =   990
               TabIndex        =   264
               Tag             =   "lblInterval"
               Top             =   1995
               Width           =   780
            End
            Begin VB.Label lblIntervalMid 
               Caption         =   "12"
               Height          =   315
               Left            =   3435
               TabIndex        =   263
               Top             =   2430
               Width           =   840
            End
            Begin VB.Label lblIntervalMax 
               Caption         =   "24"
               Height          =   315
               Left            =   5250
               TabIndex        =   262
               Top             =   2445
               Width           =   405
            End
            Begin VB.Label lblIntervalMin 
               Caption         =   "1"
               Height          =   315
               Left            =   1770
               TabIndex        =   261
               Top             =   2445
               Width           =   345
            End
            Begin VB.Label lblHousekeepingTab 
               Caption         =   "Backup Location:"
               Height          =   375
               Index           =   2
               Left            =   420
               TabIndex        =   260
               Tag             =   "lblBackupLocation"
               Top             =   3345
               Width           =   1425
            End
            Begin VB.Label lblHousekeepingDesc 
               Caption         =   "Open file explorer at the backup folder location. This will allow you to select a backup file for restoring if required."
               Height          =   675
               Index           =   5
               Left            =   2265
               TabIndex        =   259
               Tag             =   "lblBackupLocationDesc"
               Top             =   3345
               Width           =   2895
            End
            Begin VB.Label lblHousekeepingDesc 
               Caption         =   "Check this box to enable automatic backups on each startup"
               Height          =   450
               Index           =   2
               Left            =   1785
               TabIndex        =   258
               Tag             =   "lblBackupOnStartDesc"
               ToolTipText     =   "Check this box to enable automatic backups on each startup"
               Top             =   660
               Width           =   4335
            End
         End
         Begin VB.ComboBox cmbArchiveDays 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   251
            Top             =   945
            Width           =   1665
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   300
            Left            =   1905
            Style           =   1  'Graphical
            TabIndex        =   194
            ToolTipText     =   "Open a file explorer at the Archive folder location."
            Top             =   2070
            Width           =   315
         End
         Begin VB.CheckBox chkAutomaticHousekeeping 
            Caption         =   "Send Emails"
            Height          =   225
            Left            =   1935
            TabIndex        =   186
            ToolTipText     =   "Check this box to enable advice messages. If enabled, advice messages are sent periodically to this address."
            Top             =   150
            Width           =   225
         End
         Begin VB.Label lblHousekeepingTab 
            Caption         =   "-oOo-"
            Height          =   375
            Index           =   3
            Left            =   2820
            TabIndex        =   286
            Tag             =   "lblArchiveLocation"
            Top             =   2820
            Width           =   1425
         End
         Begin VB.Label lblHousekeepingDesc 
            Caption         =   "Select the number of days after which all old chats will be archived from your selected files."
            Height          =   450
            Index           =   6
            Left            =   1935
            TabIndex        =   252
            Tag             =   "lblAutomaticHousekeepingDesc"
            ToolTipText     =   "Check this box to enable automatic housekeeping"
            Top             =   1410
            Width           =   3525
         End
         Begin VB.Label lblHousekeepingTab 
            Caption         =   "Archive Location:"
            Height          =   375
            Index           =   0
            Left            =   525
            TabIndex        =   196
            Tag             =   "lblArchiveLocation"
            Top             =   2070
            Width           =   1425
         End
         Begin VB.Label lblHousekeepingDesc 
            Caption         =   "Open file explorer at the archive folder location. This will allow you to view archive files."
            Height          =   675
            Index           =   1
            Left            =   2370
            TabIndex        =   195
            Top             =   2070
            Width           =   2895
         End
         Begin VB.Label lblHousekeepingDesc 
            Caption         =   "Check this box to enable automatic housekeeping"
            Height          =   450
            Index           =   0
            Left            =   1920
            TabIndex        =   188
            Tag             =   "lblAutomaticHousekeepingDesc"
            ToolTipText     =   "Check this box to enable automatic housekeeping"
            Top             =   495
            Width           =   4335
         End
         Begin VB.Label lblAutomaticHousekeeping 
            Caption         =   "Automatic Housekeeping"
            Height          =   270
            Left            =   2280
            TabIndex        =   187
            ToolTipText     =   "Check this box to enable automatic housekeeping"
            Top             =   150
            Width           =   3120
         End
      End
   End
   Begin VB.Frame fraGeneral 
      Caption         =   "General"
      Height          =   7110
      Left            =   2295
      TabIndex        =   0
      Top             =   1260
      Width           =   8640
      Begin VB.Frame fraGeneralInner 
         BorderStyle     =   0  'None
         Height          =   6585
         Left            =   930
         TabIndex        =   78
         Top             =   390
         Width           =   5985
         Begin VB.CheckBox chkServiceProcesses 
            Caption         =   "Check the above network processes are running"
            Height          =   225
            Left            =   1485
            TabIndex        =   282
            ToolTipText     =   "Check this box to enable alarms when the above network processes are not running. Uncheck the check box to suppress the alarm."
            Top             =   4695
            Width           =   4035
         End
         Begin VB.CheckBox chkGenStartup 
            Caption         =   "Run Fire Call at Windows Startup"
            Height          =   225
            Left            =   1485
            TabIndex        =   280
            ToolTipText     =   "Check this box to enable the automatic start of the program when Windows is started."
            Top             =   5865
            Width           =   3555
         End
         Begin VB.Frame fraServiceProvider 
            Height          =   1575
            Left            =   1500
            TabIndex        =   189
            Top             =   2940
            Width           =   4005
            Begin VB.Frame fraNone 
               BorderStyle     =   0  'None
               Height          =   345
               Left            =   420
               TabIndex        =   208
               Top             =   1065
               Width           =   1665
               Begin VB.Label lblNone 
                  Caption         =   "None"
                  Height          =   270
                  Left            =   105
                  TabIndex        =   209
                  Top             =   60
                  Width           =   1020
               End
            End
            Begin VB.Frame fraOneDrive 
               BorderStyle     =   0  'None
               Height          =   345
               Left            =   420
               TabIndex        =   206
               Top             =   780
               Width           =   1665
               Begin VB.Label lblOneDrive 
                  Caption         =   "One Drive"
                  Height          =   270
                  Left            =   105
                  TabIndex        =   207
                  Top             =   60
                  Width           =   1020
               End
            End
            Begin VB.Frame fraGoogleDrive 
               BorderStyle     =   0  'None
               Height          =   345
               Left            =   420
               TabIndex        =   204
               Top             =   495
               Width           =   1665
               Begin VB.Label lblGoogleDrive 
                  Caption         =   "Google Drive"
                  Height          =   270
                  Left            =   105
                  TabIndex        =   205
                  Top             =   60
                  Width           =   1020
               End
            End
            Begin VB.Frame fraDropbox 
               BorderStyle     =   0  'None
               Height          =   345
               Left            =   420
               TabIndex        =   202
               Top             =   195
               Width           =   1695
               Begin VB.Label lblDropbox 
                  Caption         =   "Dropbox"
                  Height          =   270
                  Left            =   120
                  TabIndex        =   203
                  Top             =   60
                  Width           =   915
               End
            End
            Begin VB.OptionButton optServiceProvider 
               Height          =   285
               Index           =   0
               Left            =   195
               TabIndex        =   201
               ToolTipText     =   "Will report an error if the Dropbox processes are missing."
               Top             =   225
               Width           =   255
            End
            Begin VB.OptionButton optServiceProvider 
               Enabled         =   0   'False
               Height          =   315
               Index           =   3
               Left            =   195
               TabIndex        =   192
               ToolTipText     =   "Will not report missing process errors."
               Top             =   1080
               Width           =   315
            End
            Begin VB.OptionButton optServiceProvider 
               Enabled         =   0   'False
               Height          =   315
               Index           =   2
               Left            =   195
               TabIndex        =   191
               ToolTipText     =   "Will report an error if the OneDrive processes are missing."
               Top             =   795
               Width           =   315
            End
            Begin VB.OptionButton optServiceProvider 
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               Left            =   195
               TabIndex        =   190
               ToolTipText     =   "Will report an error if the Google Drive processes are missing."
               Top             =   510
               Width           =   270
            End
            Begin VB.Label lblGeneralTab 
               Caption         =   "Select which utility you are using to share the files and folders. Fire Call for Windows will check if the processes exist."
               Height          =   1245
               Index           =   8
               Left            =   2145
               TabIndex        =   210
               Top             =   225
               Width           =   1740
            End
         End
         Begin VB.TextBox txtSharedInputFile 
            Height          =   315
            Left            =   1485
            TabIndex        =   85
            ToolTipText     =   "Select the shared input file."
            Top             =   0
            Width           =   3660
         End
         Begin VB.CommandButton btnSharedInputFile 
            Caption         =   "..."
            Height          =   300
            Left            =   5145
            Style           =   1  'Graphical
            TabIndex        =   84
            ToolTipText     =   "Select the shared input file."
            Top             =   15
            Width           =   315
         End
         Begin VB.TextBox txtSharedOutputFile 
            Height          =   315
            Left            =   1485
            TabIndex        =   83
            ToolTipText     =   "Select the shared output file."
            Top             =   705
            Width           =   3660
         End
         Begin VB.CommandButton btnSharedOutputFile 
            Caption         =   "..."
            Height          =   300
            Left            =   5145
            Style           =   1  'Graphical
            TabIndex        =   82
            ToolTipText     =   "Select the shared output file."
            Top             =   720
            Width           =   315
         End
         Begin VB.TextBox txtExchangeFolder 
            Height          =   315
            Left            =   1485
            TabIndex        =   81
            ToolTipText     =   "Choose a shared folder for the exchange of images and text files."
            Top             =   1470
            Width           =   3660
         End
         Begin VB.CommandButton btnExchangeFolder 
            Caption         =   "..."
            Height          =   300
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   80
            ToolTipText     =   "Choose a shared folder for the exchange of images and text files."
            Top             =   1485
            Width           =   315
         End
         Begin VB.ComboBox cmbRefreshInterval 
            Height          =   315
            ItemData        =   "Form2.frx":000C
            Left            =   1485
            List            =   "Form2.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   79
            ToolTipText     =   "Set the refresh interval"
            Top             =   5085
            Width           =   4035
         End
         Begin VB.Label lblGeneralTab 
            Caption         =   "Test Provider :"
            Height          =   255
            Index           =   10
            Left            =   375
            TabIndex        =   283
            Tag             =   "lblServiceProvider"
            ToolTipText     =   "Check this box to enable regular testing of the above network processes."
            Top             =   4695
            Width           =   1470
         End
         Begin VB.Label lblGeneralTab 
            Caption         =   "Auto Start :"
            Height          =   375
            Index           =   11
            Left            =   615
            TabIndex        =   281
            Tag             =   "lblRefreshInterval"
            Top             =   5865
            Width           =   1740
         End
         Begin VB.Label lblGeneralTab 
            Caption         =   "Process to Check :"
            Height          =   255
            Index           =   3
            Left            =   90
            TabIndex        =   193
            Tag             =   "lblServiceProvider"
            Top             =   3135
            Width           =   1350
         End
         Begin VB.Label lblGeneralTab 
            Caption         =   "Shared Input File :"
            Height          =   255
            Index           =   0
            Left            =   105
            TabIndex        =   94
            Tag             =   "lblSharedInputFile"
            Top             =   45
            Width           =   1350
         End
         Begin VB.Label lblGeneralTab 
            Caption         =   "Select the shared input file."
            Height          =   300
            Index           =   5
            Left            =   1515
            TabIndex        =   93
            Tag             =   "lblSharedInputFileDesc"
            Top             =   420
            Width           =   3420
         End
         Begin VB.Label lblGeneralTab 
            Caption         =   "Shared Output File :"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   92
            Tag             =   "lblSharedOutputFolder"
            Top             =   750
            Width           =   1440
         End
         Begin VB.Label lblGeneralTab 
            Caption         =   "Select the shared output file."
            Height          =   300
            Index           =   6
            Left            =   1515
            TabIndex        =   91
            Tag             =   "lblSharedOutputFileDesc"
            Top             =   1125
            Width           =   3420
         End
         Begin VB.Label lblGeneralTab 
            Caption         =   "Exchange Folder :"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   90
            Tag             =   "lblExchangeFolder"
            Top             =   1515
            Width           =   1350
         End
         Begin VB.Label lblGeneralTab 
            Caption         =   $"Form2.frx":0010
            Height          =   900
            Index           =   7
            Left            =   1545
            TabIndex        =   89
            Tag             =   "lblExchangeFolderDesc"
            Top             =   1980
            Width           =   3600
         End
         Begin VB.Label lblGeneralTab 
            Caption         =   "Refresh Interval :"
            Height          =   375
            Index           =   4
            Left            =   195
            TabIndex        =   88
            Tag             =   "lblRefreshInterval"
            Top             =   5145
            Width           =   1740
         End
         Begin VB.Label lblGeneralTab 
            Caption         =   "Set the program's refresh interval"
            Height          =   300
            Index           =   9
            Left            =   1470
            TabIndex        =   87
            Top             =   5535
            Width           =   3750
         End
         Begin VB.Label lblExchangeFolderDesc2 
            Height          =   450
            Left            =   1530
            TabIndex        =   86
            Top             =   2415
            Width           =   3945
         End
      End
   End
   Begin VB.Frame fraConfiguration 
      Caption         =   "Configuration"
      Height          =   7260
      Left            =   360
      TabIndex        =   21
      Top             =   1365
      Width           =   8640
      Begin VB.Frame fraConfigInner 
         BorderStyle     =   0  'None
         Height          =   6660
         Left            =   750
         TabIndex        =   95
         Top             =   360
         Width           =   7260
         Begin VB.Frame fraAllowShutdowns 
            BorderStyle     =   0  'None
            Height          =   1245
            Left            =   1425
            TabIndex        =   198
            Top             =   5370
            Width           =   4575
            Begin VB.CheckBox chkAllowShutdowns 
               Caption         =   "Allow Remote Partner to Shutdown Fire Call"
               Height          =   225
               Left            =   285
               TabIndex        =   199
               Top             =   135
               Width           =   3960
            End
            Begin VB.Label lblConfigurationTab 
               Caption         =   $"Form2.frx":00BD
               Height          =   660
               Index           =   8
               Left            =   270
               TabIndex        =   200
               Top             =   525
               Width           =   3720
            End
         End
         Begin VB.CheckBox chkEnableBalloonTooltips 
            Caption         =   "Enable Balloon Tooltips on all Controls"
            Height          =   225
            Left            =   1710
            TabIndex        =   197
            ToolTipText     =   "Check the box to enable larger balloon tooltips for all controls on the main program"
            Top             =   5085
            Width           =   3405
         End
         Begin VB.Frame fraTargetClient 
            Height          =   855
            Left            =   1695
            TabIndex        =   176
            Top             =   -90
            Width           =   3675
            Begin VB.OptionButton optHandleData 
               Caption         =   " Unix Client         (UTF8)"
               Height          =   270
               Index           =   1
               Left            =   105
               TabIndex        =   178
               Top             =   495
               Width           =   2640
            End
            Begin VB.OptionButton optHandleData 
               Caption         =   " Windows Client  (ANSI)"
               Height          =   270
               Index           =   0
               Left            =   105
               TabIndex        =   177
               Top             =   180
               Width           =   2670
            End
         End
         Begin VB.CheckBox chkSingleListBox 
            Caption         =   "Single Chat Window"
            Height          =   225
            Left            =   1710
            TabIndex        =   172
            ToolTipText     =   "Check this box to merge the two chatboxes into one larger box"
            Top             =   2400
            Width           =   3270
         End
         Begin VB.TextBox txtPrefixString 
            Height          =   315
            Left            =   1695
            TabIndex        =   100
            ToolTipText     =   "Enter a prefix/nickname for outgoing messages."
            Top             =   1230
            Width           =   3660
         End
         Begin VB.CheckBox chkLoadBottom 
            Caption         =   "Load Text at Bottom of chatbox"
            Height          =   225
            Left            =   1710
            TabIndex        =   99
            ToolTipText     =   "Check this box to load new messages at the bottom of the text display."
            Top             =   1965
            Width           =   3210
         End
         Begin VB.CheckBox chkEnableScrollbars 
            Caption         =   "Enable Scrollbars on Chat Boxes"
            Height          =   225
            Left            =   1710
            TabIndex        =   98
            ToolTipText     =   "Check the box to enable the optional horizontal and vertical scrollbars"
            Top             =   4185
            Width           =   3555
         End
         Begin VB.CheckBox chkEnableTooltips 
            Caption         =   "Enable Tooltips on all Controls"
            Height          =   225
            Left            =   1710
            TabIndex        =   97
            ToolTipText     =   "Check the box to enable tooltips for all controls on the main program"
            Top             =   4650
            Width           =   3345
         End
         Begin VB.ComboBox cmbMaxLineLength 
            Height          =   315
            ItemData        =   "Form2.frx":015B
            Left            =   1710
            List            =   "Form2.frx":015D
            Style           =   2  'Dropdown List
            TabIndex        =   96
            ToolTipText     =   "The program will cut your text to a new line when this limit is reached"
            Top             =   3330
            Width           =   1575
         End
         Begin VB.Label lblConfigurationTab 
            Caption         =   "Select the type of client your chat partner will be using."
            Height          =   375
            Index           =   4
            Left            =   1695
            TabIndex        =   180
            ToolTipText     =   "You can use an ADO record stream or the FileSystemObject (FSO) to write UTF8 or ANSI compatible files"
            Top             =   840
            Width           =   3990
         End
         Begin VB.Label lblConfigurationTab 
            Caption         =   "Target Client :"
            Height          =   255
            Index           =   0
            Left            =   645
            TabIndex        =   179
            Tag             =   "lblTargetClient"
            Top             =   105
            Width           =   1065
         End
         Begin VB.Label lblConfigurationTab 
            Caption         =   "This determines whether the separate chat boxes are merged into one larger box"
            Height          =   660
            Index           =   6
            Left            =   1710
            TabIndex        =   173
            Top             =   2745
            Width           =   3615
         End
         Begin VB.Label lblConfigurationTab 
            Caption         =   "Enter a prefix/nickname for outgoing messages."
            Height          =   375
            Index           =   5
            Left            =   1755
            TabIndex        =   104
            Top             =   1650
            Width           =   3705
         End
         Begin VB.Label lblConfigurationTab 
            Caption         =   "Prefix String :"
            Height          =   255
            Index           =   1
            Left            =   705
            TabIndex        =   103
            Tag             =   "lblPrefixString"
            Top             =   1275
            Width           =   1065
         End
         Begin VB.Label lblConfigurationTab 
            Caption         =   "Maximum Line Length : "
            Height          =   285
            Index           =   2
            Left            =   30
            TabIndex        =   102
            Tag             =   "lblMaxLineLength"
            Top             =   3375
            Width           =   1740
         End
         Begin VB.Label lblConfigurationTab 
            Caption         =   "Choose the maximum length for your texts."
            Height          =   300
            Index           =   7
            Left            =   1710
            TabIndex        =   101
            Top             =   3780
            Width           =   3750
         End
      End
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6105
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Save the changes you have made to the preferences"
      Top             =   10140
      Width           =   1320
   End
   Begin VB.Frame fraHousekeepingButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   6840
      TabIndex        =   181
      ToolTipText     =   "Opens the Housekeeping tab"
      Top             =   -90
      Width           =   930
      Begin VB.PictureBox picHousekeeping 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   165
         Picture         =   "Form2.frx":015F
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   182
         ToolTipText     =   "Opens the Housekeeping tab"
         Top             =   225
         Width           =   600
      End
      Begin VB.Label lblHousekeeping 
         Caption         =   "House"
         Height          =   225
         Left            =   210
         TabIndex        =   183
         ToolTipText     =   "Opens the Housekeeping tab"
         Top             =   810
         Width           =   540
      End
   End
   Begin VB.CommandButton btnHelp 
      Caption         =   "Help"
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   174
      ToolTipText     =   "Open the help utility"
      Top             =   10155
      Width           =   1320
   End
   Begin VB.Frame fraSoundsButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   5880
      TabIndex        =   30
      Top             =   -105
      Width           =   930
      Begin VB.PictureBox picSounds 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   630
         Left            =   165
         Picture         =   "Form2.frx":0D7F
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   31
         ToolTipText     =   "Opens the Window tab"
         Top             =   210
         Width           =   630
      End
      Begin VB.Label lblSounds 
         Caption         =   "Sounds"
         Height          =   240
         Left            =   210
         TabIndex        =   32
         Top             =   825
         Width           =   615
      End
   End
   Begin VB.Frame fraTextsButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   4920
      TabIndex        =   27
      Top             =   -105
      Width           =   930
      Begin VB.PictureBox picTexts 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   630
         Left            =   165
         Picture         =   "Form2.frx":133E
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   28
         ToolTipText     =   "Opens the Window tab"
         Top             =   225
         Width           =   630
      End
      Begin VB.Label lblTexts 
         Caption         =   "Texts"
         Height          =   240
         Left            =   270
         TabIndex        =   29
         Top             =   825
         Width           =   615
      End
   End
   Begin VB.Timer themeTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1440
      Top             =   10125
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Close"
      Height          =   360
      Left            =   7500
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Close the utility"
      Top             =   10155
      Width           =   1320
   End
   Begin VB.Frame fraWindowButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   7815
      TabIndex        =   16
      Top             =   -90
      Width           =   930
      Begin VB.PictureBox picWindow 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   165
         Picture         =   "Form2.frx":1940
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   17
         ToolTipText     =   "Opens the Window tab"
         Top             =   225
         Width           =   600
      End
      Begin VB.Label lblWindow 
         Caption         =   "Window"
         Height          =   240
         Left            =   180
         TabIndex        =   18
         Top             =   825
         Width           =   615
      End
   End
   Begin VB.Frame fraFontsButtons 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   3960
      TabIndex        =   13
      Top             =   -90
      Width           =   930
      Begin VB.PictureBox picFonts 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   180
         Picture         =   "Form2.frx":2188
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   14
         ToolTipText     =   "Opens the Fonts tab"
         Top             =   210
         Width           =   600
      End
      Begin VB.Label lblFonts 
         Caption         =   "Fonts"
         Height          =   240
         Left            =   270
         TabIndex        =   15
         Top             =   825
         Width           =   510
      End
   End
   Begin VB.Frame fraEmojisButtons 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   3000
      TabIndex        =   10
      Top             =   -90
      Width           =   930
      Begin VB.PictureBox picEmojis 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   180
         Picture         =   "Form2.frx":2974
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   11
         ToolTipText     =   "Opens the Emojis tab"
         Top             =   210
         Width           =   600
      End
      Begin VB.Label lblEmojis 
         Caption         =   "Emojis"
         Height          =   240
         Left            =   270
         TabIndex        =   12
         Top             =   825
         Width           =   510
      End
   End
   Begin VB.Frame fraEmailButtons 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   2040
      TabIndex        =   7
      Top             =   -90
      Width           =   930
      Begin VB.PictureBox picEmail 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   180
         Picture         =   "Form2.frx":2E57
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   8
         ToolTipText     =   "Opens the email tab"
         Top             =   210
         Width           =   600
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email"
         Height          =   240
         Left            =   270
         TabIndex        =   9
         Top             =   825
         Width           =   510
      End
   End
   Begin VB.Frame fraConfigurationButtons 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   1080
      TabIndex        =   4
      Top             =   -90
      Width           =   930
      Begin VB.PictureBox picConfig 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   180
         Picture         =   "Form2.frx":33F3
         ScaleHeight     =   600
         ScaleWidth      =   600
         TabIndex        =   5
         ToolTipText     =   "Opens the configuration tab"
         Top             =   195
         Width           =   600
      End
      Begin VB.Label lblConfig 
         Caption         =   "Config."
         Height          =   240
         Left            =   195
         TabIndex        =   6
         Top             =   825
         Width           =   630
      End
   End
   Begin VB.Frame fraGeneralButtons 
      Height          =   1140
      Left            =   120
      TabIndex        =   1
      Top             =   -90
      Width           =   930
      Begin VB.PictureBox picGeneral 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   240
         Picture         =   "Form2.frx":3C55
         ScaleHeight     =   405
         ScaleWidth      =   420
         TabIndex        =   2
         ToolTipText     =   "Opens the general tab"
         Top             =   300
         Width           =   420
      End
      Begin VB.Label lblGeneral 
         Caption         =   "General"
         Height          =   240
         Left            =   195
         TabIndex        =   3
         Top             =   825
         Width           =   705
      End
   End
   Begin VB.Frame fraTexts 
      Caption         =   "Texts"
      Height          =   7035
      Left            =   1530
      TabIndex        =   26
      Top             =   2910
      Width           =   8640
      Begin VB.Frame fraTextsInner 
         BorderStyle     =   0  'None
         Height          =   6585
         Left            =   990
         TabIndex        =   113
         Top             =   345
         Width           =   5940
         Begin VB.ComboBox cmbTTFN 
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   148
            Top             =   480
            Width           =   3240
         End
         Begin VB.TextBox txtStringToAdd 
            Height          =   330
            Left            =   1500
            TabIndex        =   147
            Text            =   "Enter text here and click + button below"
            ToolTipText     =   "Enter text here and click + button to add to any list below"
            Top             =   0
            Width           =   3225
         End
         Begin VB.CommandButton btnTtfnAdd 
            Caption         =   "+"
            Height          =   300
            Left            =   4860
            Style           =   1  'Graphical
            TabIndex        =   146
            ToolTipText     =   "Add the new text from the text box above into the button's list of pre-defined texts"
            Top             =   495
            Width           =   315
         End
         Begin VB.CommandButton btnTtfnRemove 
            Caption         =   "-"
            Height          =   300
            Left            =   5205
            Style           =   1  'Graphical
            TabIndex        =   145
            ToolTipText     =   "Delete the currently selected text"
            Top             =   495
            Width           =   315
         End
         Begin VB.ComboBox cmbWell 
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   144
            Top             =   1020
            Width           =   3240
         End
         Begin VB.CommandButton btnWellAdd 
            Caption         =   "+"
            Height          =   300
            Left            =   4860
            Style           =   1  'Graphical
            TabIndex        =   143
            ToolTipText     =   "Add the new text from the text box above into the button's list of pre-defined texts"
            Top             =   1020
            Width           =   315
         End
         Begin VB.CommandButton btnWellRemove 
            Caption         =   "-"
            Height          =   300
            Left            =   5205
            Style           =   1  'Graphical
            TabIndex        =   142
            ToolTipText     =   "Delete the currently selected text"
            Top             =   1020
            Width           =   315
         End
         Begin VB.ComboBox cmbNews 
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   141
            Top             =   1530
            Width           =   3240
         End
         Begin VB.CommandButton btnNewsAdd 
            Caption         =   "+"
            Height          =   300
            Left            =   4860
            Style           =   1  'Graphical
            TabIndex        =   140
            ToolTipText     =   "Add the new text from the text box above into the button's list of pre-defined texts"
            Top             =   1530
            Width           =   315
         End
         Begin VB.CommandButton btnNewsRemove 
            Caption         =   "-"
            Height          =   300
            Left            =   5205
            Style           =   1  'Graphical
            TabIndex        =   139
            ToolTipText     =   "Delete the currently selected text"
            Top             =   1530
            Width           =   315
         End
         Begin VB.ComboBox cmbMorn 
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   138
            Top             =   2040
            Width           =   3240
         End
         Begin VB.CommandButton btnMornAdd 
            Caption         =   "+"
            Height          =   300
            Left            =   4860
            Style           =   1  'Graphical
            TabIndex        =   137
            ToolTipText     =   "Add the new text from the text box above into the button's list of pre-defined texts"
            Top             =   2040
            Width           =   315
         End
         Begin VB.CommandButton btnMornRemove 
            Caption         =   "-"
            Height          =   300
            Left            =   5205
            Style           =   1  'Graphical
            TabIndex        =   136
            ToolTipText     =   "Delete the currently selected text"
            Top             =   2040
            Width           =   315
         End
         Begin VB.ComboBox cmbWot 
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   135
            Top             =   2535
            Width           =   3240
         End
         Begin VB.CommandButton btnWotAdd 
            Caption         =   "+"
            Height          =   300
            Left            =   4860
            Style           =   1  'Graphical
            TabIndex        =   134
            ToolTipText     =   "Add the new text from the text box above into the button's list of pre-defined texts"
            Top             =   2535
            Width           =   315
         End
         Begin VB.CommandButton btnWotRemove 
            Caption         =   "-"
            Height          =   300
            Left            =   5205
            Style           =   1  'Graphical
            TabIndex        =   133
            ToolTipText     =   "Delete the currently selected text"
            Top             =   2535
            Width           =   315
         End
         Begin VB.ComboBox cmbWth 
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   132
            Top             =   3030
            Width           =   3240
         End
         Begin VB.CommandButton btnWthAdd 
            Caption         =   "+"
            Height          =   300
            Left            =   4860
            Style           =   1  'Graphical
            TabIndex        =   131
            ToolTipText     =   "Add the new text from the text box above into the button's list of pre-defined texts"
            Top             =   3030
            Width           =   315
         End
         Begin VB.CommandButton btnWthRemove 
            Caption         =   "-"
            Height          =   300
            Left            =   5205
            Style           =   1  'Graphical
            TabIndex        =   130
            ToolTipText     =   "Delete the currently selected text"
            Top             =   3030
            Width           =   315
         End
         Begin VB.ComboBox cmbPrg 
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   129
            Top             =   3525
            Width           =   3240
         End
         Begin VB.CommandButton btnPrgAdd 
            Caption         =   "+"
            Height          =   300
            Left            =   4860
            Style           =   1  'Graphical
            TabIndex        =   128
            ToolTipText     =   "Add the new text from the text box above into the button's list of pre-defined texts"
            Top             =   3510
            Width           =   315
         End
         Begin VB.CommandButton btnPrgRemove 
            Caption         =   "-"
            Height          =   300
            Left            =   5205
            Style           =   1  'Graphical
            TabIndex        =   127
            ToolTipText     =   "Delete the currently selected text"
            Top             =   3510
            Width           =   315
         End
         Begin VB.ComboBox cmbGdn 
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   126
            Top             =   4005
            Width           =   3240
         End
         Begin VB.CommandButton btnGdnAdd 
            Caption         =   "+"
            Height          =   300
            Left            =   4860
            Style           =   1  'Graphical
            TabIndex        =   125
            ToolTipText     =   "Add the new text from the text box above into the button's list of pre-defined texts"
            Top             =   4035
            Width           =   315
         End
         Begin VB.CommandButton btnGdnRemove 
            Caption         =   "-"
            Height          =   300
            Left            =   5205
            Style           =   1  'Graphical
            TabIndex        =   124
            ToolTipText     =   "Delete the currently selected text"
            Top             =   4035
            Width           =   315
         End
         Begin VB.ComboBox cmbBusy 
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   123
            Top             =   4500
            Width           =   3240
         End
         Begin VB.CommandButton btnBusyAdd 
            Caption         =   "+"
            Height          =   300
            Left            =   4860
            Style           =   1  'Graphical
            TabIndex        =   122
            ToolTipText     =   "Add the new text from the text box above into the button's list of pre-defined texts"
            Top             =   4500
            Width           =   315
         End
         Begin VB.CommandButton btnBusyRemove 
            Caption         =   "-"
            Height          =   300
            Left            =   5205
            Style           =   1  'Graphical
            TabIndex        =   121
            ToolTipText     =   "Delete the currently selected text"
            Top             =   4500
            Width           =   315
         End
         Begin VB.ComboBox cmbCod 
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   120
            Top             =   4980
            Width           =   3240
         End
         Begin VB.CommandButton btnCodAdd 
            Caption         =   "+"
            Height          =   300
            Left            =   4860
            Style           =   1  'Graphical
            TabIndex        =   119
            ToolTipText     =   "Add the new text from the text box above into the button's list of pre-defined texts"
            Top             =   4980
            Width           =   315
         End
         Begin VB.CommandButton btnCodRemove 
            Caption         =   "-"
            Height          =   300
            Left            =   5205
            Style           =   1  'Graphical
            TabIndex        =   118
            ToolTipText     =   "Delete the currently selected text"
            Top             =   4980
            Width           =   315
         End
         Begin VB.ComboBox cmbOut 
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   117
            Top             =   5475
            Width           =   3240
         End
         Begin VB.CommandButton btnOutAdd 
            Caption         =   "+"
            Height          =   300
            Left            =   4860
            Style           =   1  'Graphical
            TabIndex        =   116
            ToolTipText     =   "Add the new text from the text box above into the button's list of pre-defined texts"
            Top             =   5475
            Width           =   315
         End
         Begin VB.CommandButton btnOutRemove 
            Caption         =   "-"
            Height          =   300
            Left            =   5205
            Style           =   1  'Graphical
            TabIndex        =   115
            ToolTipText     =   "Delete the currently selected text"
            Top             =   5475
            Width           =   315
         End
         Begin VB.CommandButton btnDeleteText 
            Caption         =   "-"
            Height          =   300
            Left            =   4860
            Style           =   1  'Graphical
            TabIndex        =   114
            ToolTipText     =   "Delete Text"
            Top             =   30
            Width           =   315
         End
         Begin VB.Label lblTextsTab 
            Caption         =   "TTFN Button :"
            Height          =   255
            Index           =   0
            Left            =   15
            TabIndex        =   160
            Tag             =   "lblTTFNButton"
            Top             =   525
            Width           =   1350
         End
         Begin VB.Label lblTextsTab 
            Caption         =   "Well Button :"
            Height          =   255
            Index           =   1
            Left            =   15
            TabIndex        =   159
            Tag             =   "lblWellButton"
            Top             =   1035
            Width           =   1350
         End
         Begin VB.Label lblTextsTab 
            Caption         =   "News Button :"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   158
            Tag             =   "lblNewsButton"
            Top             =   1560
            Width           =   1350
         End
         Begin VB.Label lblTextsTab 
            Caption         =   "Morn Button :"
            Height          =   255
            Index           =   3
            Left            =   15
            TabIndex        =   157
            Tag             =   "lblMornButton"
            Top             =   2070
            Width           =   1350
         End
         Begin VB.Label lblTextsTab 
            Caption         =   "Wot Button :"
            Height          =   255
            Index           =   4
            Left            =   15
            TabIndex        =   156
            Tag             =   "lblWotButton"
            Top             =   2580
            Width           =   1350
         End
         Begin VB.Label lblTextsTab 
            Caption         =   "Wth Button :"
            Height          =   255
            Index           =   5
            Left            =   15
            TabIndex        =   155
            Tag             =   "lblWthButton"
            Top             =   3090
            Width           =   1350
         End
         Begin VB.Label lblTextsTab 
            Caption         =   "Prg Button :"
            Height          =   255
            Index           =   6
            Left            =   15
            TabIndex        =   154
            Tag             =   "lblPrgButton"
            Top             =   3570
            Width           =   1350
         End
         Begin VB.Label lblTextsTab 
            Caption         =   "Gdn Button :"
            Height          =   255
            Index           =   7
            Left            =   15
            TabIndex        =   153
            Tag             =   "lblGdnButton"
            Top             =   4065
            Width           =   1350
         End
         Begin VB.Label lblTextsTab 
            Caption         =   "Busy Button :"
            Height          =   255
            Index           =   8
            Left            =   15
            TabIndex        =   152
            Tag             =   "lblBusyButton"
            Top             =   4530
            Width           =   1350
         End
         Begin VB.Label lblTextsTab 
            Caption         =   "Cod Button :"
            Height          =   255
            Index           =   9
            Left            =   15
            TabIndex        =   151
            Tag             =   "lblCodButton"
            Top             =   5010
            Width           =   1350
         End
         Begin VB.Label lblTextsTab 
            Caption         =   "Out Button :"
            Height          =   255
            Index           =   10
            Left            =   15
            TabIndex        =   150
            Tag             =   "lblOutButton"
            Top             =   5505
            Width           =   1350
         End
         Begin VB.Label lblTextsDesc 
            Caption         =   "Here you can change or add to the pre-defined text buttons that appear at the bottom of the program."
            Height          =   570
            Left            =   1500
            TabIndex        =   149
            Top             =   5970
            Width           =   4050
         End
      End
   End
   Begin VB.Frame fraFonts 
      Caption         =   "Fonts"
      Height          =   3705
      Left            =   1515
      TabIndex        =   24
      Top             =   2985
      Width           =   8640
      Begin VB.Frame fraFontsInner 
         BorderStyle     =   0  'None
         Height          =   3210
         Left            =   915
         TabIndex        =   59
         Top             =   300
         Width           =   5895
         Begin VB.CommandButton btnTextFont 
            Caption         =   "..."
            Height          =   300
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Choose a font."
            Top             =   0
            Width           =   315
         End
         Begin VB.TextBox txtFontSize 
            Height          =   315
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   64
            Text            =   "8"
            ToolTipText     =   "Choose the font size in the two chat windows"
            Top             =   990
            Width           =   510
         End
         Begin VB.TextBox txtPrefsFontSize 
            Height          =   315
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   63
            Text            =   "8"
            ToolTipText     =   "Choose a font size to be used within this preferences window only"
            Top             =   2700
            Width           =   510
         End
         Begin VB.CommandButton btnPrefsFont 
            Caption         =   "..."
            Height          =   300
            Left            =   4935
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Choose a font."
            Top             =   1710
            Width           =   315
         End
         Begin VB.TextBox txtPrefsFont 
            Height          =   315
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   61
            Text            =   "Times New Roman"
            ToolTipText     =   "Choose a font to be used only for this preferences window"
            Top             =   1725
            Width           =   3285
         End
         Begin VB.TextBox txtTextFont 
            Height          =   315
            Left            =   1635
            Locked          =   -1  'True
            TabIndex        =   60
            Text            =   "Times New Roman"
            ToolTipText     =   "Choose a font to be used for the text in the two chat windows"
            Top             =   15
            Width           =   3240
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Chat Box Font:"
            Height          =   330
            Index           =   0
            Left            =   330
            TabIndex        =   73
            Tag             =   "lblTextFont"
            ToolTipText     =   "We suggest Linux Biolinum G at 8pt - which you will find in the FCW program folder"
            Top             =   60
            Width           =   1245
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Choose a font to be used for the text in the chat window"
            Height          =   480
            Index           =   4
            Left            =   1695
            TabIndex        =   72
            ToolTipText     =   "We suggest Linux Biolinum G at 8pt - which you will find in the FCW program folder"
            Top             =   420
            Width           =   3915
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Font Size :"
            Height          =   330
            Index           =   1
            Left            =   705
            TabIndex        =   71
            Tag             =   "lblFontSize"
            Top             =   1005
            Width           =   885
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "The chosen font size"
            Height          =   480
            Index           =   7
            Left            =   2295
            TabIndex        =   70
            ToolTipText     =   "Choose a font size that fits the text boxes"
            Top             =   2730
            Width           =   2400
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Font Size :"
            Height          =   330
            Index           =   3
            Left            =   690
            TabIndex        =   69
            Tag             =   "lblPrefsFontSize"
            Top             =   2730
            Width           =   885
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Config Window Font:"
            Height          =   300
            Index           =   2
            Left            =   0
            TabIndex        =   68
            Tag             =   "lblPrefsFont"
            ToolTipText     =   "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
            Top             =   1755
            Width           =   1635
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Choose a font to be used for the text in this preferences window alone"
            Height          =   480
            Index           =   6
            Left            =   1605
            TabIndex        =   67
            ToolTipText     =   "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
            Top             =   2115
            Width           =   4035
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "The chosen font size"
            Height          =   315
            Index           =   5
            Left            =   2310
            TabIndex        =   66
            ToolTipText     =   "Choose a font size that fits the text boxes"
            Top             =   1005
            Width           =   2385
         End
      End
   End
   Begin VB.Frame fraEmail 
      Caption         =   "Email"
      Height          =   8250
      Left            =   1080
      TabIndex        =   22
      Top             =   1950
      Width           =   8640
      Begin VB.Frame fraEmailInner 
         BorderStyle     =   0  'None
         Height          =   7665
         Left            =   510
         TabIndex        =   74
         Top             =   450
         Width           =   7470
         Begin VB.TextBox txtSmtpConfigName 
            Height          =   315
            Left            =   4680
            TabIndex        =   249
            ToolTipText     =   "Enter the configuration identifier here"
            Top             =   1305
            Width           =   1290
         End
         Begin VB.ComboBox cmbSmtpConfig 
            Height          =   315
            ItemData        =   "Form2.frx":41EB
            Left            =   2025
            List            =   "Form2.frx":41ED
            Style           =   2  'Dropdown List
            TabIndex        =   247
            ToolTipText     =   "Select which SMTP configuration slot you would like to operate."
            Top             =   1305
            Width           =   1860
         End
         Begin VB.Frame fraEmailfra 
            Height          =   4140
            Left            =   7215
            TabIndex        =   240
            Top             =   2070
            Visible         =   0   'False
            Width           =   5220
            Begin VB.CommandButton Command2 
               Caption         =   "Clear"
               Height          =   420
               Left            =   60
               TabIndex        =   246
               Top             =   3630
               Width           =   1080
            End
            Begin VB.CommandButton btnCloseEmailFra 
               Caption         =   "Close"
               Height          =   420
               Left            =   4065
               TabIndex        =   245
               Top             =   3630
               Width           =   1080
            End
            Begin VB.TextBox txtEmailLog 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   7.5
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3045
               Left            =   60
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   244
               Top             =   510
               Width           =   5070
            End
            Begin VB.PictureBox Picture 
               BackColor       =   &H00C0E0FF&
               BorderStyle     =   0  'None
               Height          =   210
               Index           =   1
               Left            =   4920
               Picture         =   "Form2.frx":41EF
               ScaleHeight     =   210
               ScaleWidth      =   225
               TabIndex        =   242
               ToolTipText     =   "Click to close the image"
               Top             =   195
               Width           =   225
            End
            Begin VB.PictureBox Picture 
               BackColor       =   &H00C0E0FF&
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   0
               Left            =   90
               Picture         =   "Form2.frx":441C
               ScaleHeight     =   240
               ScaleWidth      =   255
               TabIndex        =   241
               ToolTipText     =   "Click to close the image"
               Top             =   180
               Width           =   255
            End
            Begin VB.TextBox Text2 
               Height          =   300
               Left            =   60
               TabIndex        =   243
               Text            =   "                                               Email Log"
               Top             =   150
               Width           =   5115
            End
         End
         Begin VB.CheckBox chkAppendConfig 
            Caption         =   "Append the above details to test emails"
            Height          =   225
            Left            =   2025
            TabIndex        =   239
            ToolTipText     =   "This will make it easier to identify which settings belong to which test email"
            Top             =   4275
            Width           =   3900
         End
         Begin VB.ComboBox cmbSmtpSecurity 
            Height          =   315
            Left            =   2025
            Style           =   2  'Dropdown List
            TabIndex        =   236
            ToolTipText     =   "Choose the security level, none, SSL or TLS"
            Top             =   3060
            Width           =   1845
         End
         Begin VB.TextBox txtSmtpPort 
            Height          =   285
            Left            =   2025
            TabIndex        =   229
            Text            =   "Choose the SMTP port number, typically port 25, 465 or 587 for TLS"
            ToolTipText     =   "Enter your email server's SMTP port here, you will find this in your email client outgoing email configuration, eg. 25"
            Top             =   2220
            Width           =   645
         End
         Begin VB.ComboBox cmbSmtpAuthenticate 
            Height          =   315
            Left            =   2025
            Style           =   2  'Dropdown List
            TabIndex        =   228
            ToolTipText     =   "Select the authentication method"
            Top             =   2655
            Width           =   1845
         End
         Begin VB.TextBox txtSMTPNoPassword 
            Height          =   285
            Left            =   3570
            Locked          =   -1  'True
            TabIndex        =   224
            ToolTipText     =   "Enter your email server's SMTP details here, you will find those in your email client outgoing email configuration"
            Top             =   3555
            Width           =   3960
         End
         Begin VB.CommandButton btnSeePassword 
            Height          =   315
            Left            =   6045
            Picture         =   "Form2.frx":48A6
            Style           =   1  'Graphical
            TabIndex        =   223
            ToolTipText     =   "Click here to expose the password to prying eyes..."
            Top             =   3855
            Width           =   315
         End
         Begin VB.TextBox txtSmtpPassword 
            Height          =   285
            Left            =   2025
            TabIndex        =   221
            ToolTipText     =   "Enter your email server's SMTP password here, you will find those in your email client outgoing email configuration"
            Top             =   3870
            Width           =   3960
         End
         Begin VB.TextBox txtSmtpUsername 
            Height          =   285
            Left            =   2025
            TabIndex        =   219
            ToolTipText     =   "Enter your email server's SMTP username details here, you will find those in your email client outgoing email configuration"
            Top             =   3465
            Width           =   3960
         End
         Begin VB.CommandButton btnTestEmail 
            Caption         =   "Test"
            Enabled         =   0   'False
            Height          =   420
            Left            =   5145
            Style           =   1  'Graphical
            TabIndex        =   218
            Top             =   6825
            Width           =   1080
         End
         Begin VB.CheckBox chkSendErrorEmails 
            Caption         =   "Send Error Emails"
            Height          =   225
            Left            =   2025
            TabIndex        =   175
            Top             =   390
            Width           =   2025
         End
         Begin VB.TextBox txtRecipientEmail 
            Height          =   285
            Left            =   2025
            TabIndex        =   167
            ToolTipText     =   "Enter the email address where you wish to receive email updates on Fire Call Win's operational status."
            Top             =   4665
            Width           =   3960
         End
         Begin VB.TextBox txtEmailSubject 
            Height          =   285
            Left            =   2025
            TabIndex        =   166
            ToolTipText     =   "If you have a preference for a specific subject text, enter it here."
            Top             =   5085
            Width           =   3960
         End
         Begin VB.TextBox txtEmailMessage 
            Height          =   1035
            Left            =   2025
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   165
            ToolTipText     =   "If you have a preference for a specific email text, enter it here."
            Top             =   5490
            Width           =   4185
         End
         Begin VB.CheckBox chkSendEmails 
            Caption         =   "Send Advice Emails"
            Height          =   225
            Left            =   2025
            TabIndex        =   76
            Top             =   0
            Width           =   3105
         End
         Begin VB.ComboBox cmbAdviceInterval 
            Height          =   315
            ItemData        =   "Form2.frx":4B67
            Left            =   2025
            List            =   "Form2.frx":4B69
            Style           =   2  'Dropdown List
            TabIndex        =   75
            ToolTipText     =   "Advice messages are sent when new data is received but not more often than at the specified interval."
            Top             =   810
            Width           =   3960
         End
         Begin VB.Frame fraSMTPframe 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   765
            TabIndex        =   230
            Tag             =   "Do NOT delete me - I am here for balloon tooltip"
            Top             =   1785
            Width           =   6705
            Begin VB.TextBox txtSmtpServer 
               Height          =   285
               Left            =   1260
               TabIndex        =   231
               Text            =   "This is the SMTP server name as supplied by your email provider"
               ToolTipText     =   $"Form2.frx":4B6B
               Top             =   30
               Width           =   3960
            End
            Begin VB.Label lblEmailTab 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SMTP Server"
               Height          =   195
               Index           =   1
               Left            =   195
               TabIndex        =   232
               Tag             =   "lblServer"
               Top             =   45
               Width           =   960
            End
         End
         Begin VB.Label lblEmailTab 
            Caption         =   "Tag"
            Height          =   345
            Index           =   16
            Left            =   4245
            TabIndex        =   250
            Tag             =   "lblAdviceInterval"
            ToolTipText     =   "Give this configuration an identifier"
            Top             =   1350
            Width           =   705
         End
         Begin VB.Label lblEmailTab 
            Caption         =   "SMTP Configuration"
            Height          =   345
            Index           =   7
            Left            =   435
            TabIndex        =   248
            Tag             =   "lblAdviceInterval"
            ToolTipText     =   "The SMTP details will be saved to this chosen configuration slot"
            Top             =   1365
            Width           =   1725
         End
         Begin VB.Label lblEmailTab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(SSL is the default)"
            Height          =   195
            Index           =   14
            Left            =   4635
            TabIndex        =   238
            Tag             =   "lblMsg"
            Top             =   3075
            Width           =   1815
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblEmailTab 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SMTP Security"
            Height          =   195
            Index           =   15
            Left            =   840
            TabIndex        =   237
            Tag             =   "lblServer"
            Top             =   3105
            Width           =   1065
         End
         Begin VB.Label lblEmailTab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(Base 64 is the default)"
            Height          =   195
            Index           =   13
            Left            =   4350
            TabIndex        =   235
            Tag             =   "lblMsg"
            Top             =   2670
            Width           =   2070
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblEmailTab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(Typically port 25, 465 or 587)"
            Height          =   345
            Index           =   12
            Left            =   3855
            TabIndex        =   234
            Tag             =   "lblMsg"
            Top             =   2310
            Width           =   2370
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblEmailTab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "If you make any changes to the above, press SAVE before trying to test. Check  your email client to see if any email has arrived."
            Height          =   780
            Index           =   11
            Left            =   2040
            TabIndex        =   233
            Tag             =   "lblMsg"
            Top             =   6705
            Width           =   2925
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblEmailTab 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SMTP Authentication"
            Height          =   195
            Index           =   10
            Left            =   390
            TabIndex        =   227
            Tag             =   "lblServer"
            Top             =   2700
            Width           =   1515
         End
         Begin VB.Label lblEmailTab 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SMTP Port"
            Height          =   195
            Index           =   9
            Left            =   1110
            TabIndex        =   226
            Tag             =   "lblServer"
            Top             =   2250
            Width           =   780
         End
         Begin VB.Label lblEmailTab 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SMTP Password"
            Height          =   195
            Index           =   6
            Left            =   705
            TabIndex        =   222
            Tag             =   "lblServer"
            Top             =   3900
            Width           =   1185
         End
         Begin VB.Label lblEmailTab 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SMTP Username"
            Height          =   195
            Index           =   5
            Left            =   690
            TabIndex        =   220
            Tag             =   "lblServer"
            Top             =   3495
            Width           =   1215
         End
         Begin VB.Label lblEmailTab 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recipient Email"
            Height          =   195
            Index           =   2
            Left            =   780
            TabIndex        =   170
            Tag             =   "lblTo"
            Top             =   4710
            Width           =   1095
         End
         Begin VB.Label lblEmailTab 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Subject"
            Height          =   195
            Index           =   3
            Left            =   1350
            TabIndex        =   169
            Tag             =   "lblSubject"
            Top             =   5115
            Width           =   540
         End
         Begin VB.Label lblEmailTab 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Test Message"
            Height          =   195
            Index           =   4
            Left            =   900
            TabIndex        =   168
            Tag             =   "lblMsg"
            Top             =   5520
            Width           =   1005
         End
         Begin VB.Label lblEmailTab 
            Caption         =   "Advice Interval"
            Height          =   345
            Index           =   0
            Left            =   795
            TabIndex        =   77
            Tag             =   "lblAdviceInterval"
            ToolTipText     =   "Advice messages are sent when new data is received but not more often than at the specified interval."
            Top             =   870
            Width           =   1740
         End
      End
      Begin VB.Label lblEmailTab 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SMTP Server"
         Height          =   195
         Index           =   8
         Left            =   1320
         TabIndex        =   225
         Tag             =   "lblServer"
         Top             =   2745
         Width           =   960
      End
   End
   Begin VB.Frame fraEmojis 
      Caption         =   "Emojis"
      Height          =   3285
      Left            =   1140
      TabIndex        =   23
      Top             =   2490
      Width           =   8655
      Begin VB.Frame fraEmojisInner 
         BorderStyle     =   0  'None
         Height          =   2730
         Left            =   1410
         TabIndex        =   105
         Top             =   375
         Width           =   5565
         Begin VB.ComboBox cmbEmojiSet 
            Height          =   315
            ItemData        =   "Form2.frx":4BF8
            Left            =   1185
            List            =   "Form2.frx":4BFA
            Style           =   2  'Dropdown List
            TabIndex        =   107
            ToolTipText     =   "Choose the emoji set to use"
            Top             =   0
            Width           =   1710
         End
         Begin VB.CommandButton btnEmojiLocation 
            Caption         =   "..."
            Height          =   300
            Left            =   1185
            Style           =   1  'Graphical
            TabIndex        =   106
            ToolTipText     =   "Open a file explorer at the Emoji folder location."
            Top             =   1980
            Width           =   315
         End
         Begin VB.Label lblEmojiTab 
            Caption         =   "Emoji Set:"
            Height          =   375
            Index           =   0
            Left            =   405
            TabIndex        =   112
            Tag             =   "lblEmojiSet"
            Top             =   60
            Width           =   1740
         End
         Begin VB.Label lblEmojiTab 
            Caption         =   "Choose the Emoji set you wish to use."
            Height          =   300
            Index           =   2
            Left            =   1185
            TabIndex        =   111
            Top             =   450
            Width           =   3750
         End
         Begin VB.Label lblEmojiTab 
            Caption         =   $"Form2.frx":4BFC
            Height          =   825
            Index           =   4
            Left            =   1170
            TabIndex        =   110
            Top             =   870
            Width           =   3750
         End
         Begin VB.Label lblEmojiTab 
            Caption         =   "Open a file explorer at the Emoji folder location."
            Height          =   465
            Index           =   3
            Left            =   1665
            TabIndex        =   109
            Top             =   1965
            Width           =   2895
         End
         Begin VB.Label lblEmojiTab 
            Caption         =   "Emoji Location:"
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   108
            Tag             =   "lblEmojiLocation"
            Top             =   2010
            Width           =   1230
         End
      End
   End
   Begin VB.Frame fraSounds 
      Caption         =   "Sounds"
      Height          =   7095
      Left            =   90
      TabIndex        =   33
      Top             =   1110
      Visible         =   0   'False
      Width           =   8640
      Begin VB.Frame fraSoundsInner 
         BorderStyle     =   0  'None
         Height          =   6630
         Left            =   1125
         TabIndex        =   50
         Top             =   300
         Width           =   6480
         Begin VB.CheckBox chkEnableAlarmSound 
            Caption         =   "Enable Alarm Sound"
            Height          =   225
            Left            =   1365
            TabIndex        =   275
            ToolTipText     =   "Check this box to enable or disable the sounds played during any alarm raised."
            Top             =   870
            Width           =   3570
         End
         Begin vb6projectCCRSlider.Slider sliRecordingQuality 
            Height          =   450
            Left            =   1230
            TabIndex        =   214
            ToolTipText     =   "Quality of recording affects WAV file size."
            Top             =   5520
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   794
            Min             =   1
            Max             =   5
            Value           =   3
            SelStart        =   3
         End
         Begin VB.ComboBox cmbCaptureDevices 
            Height          =   315
            ItemData        =   "Form2.frx":4CC2
            Left            =   1365
            List            =   "Form2.frx":4CC4
            TabIndex        =   211
            Text            =   "cmbCaptureDevices"
            Top             =   4455
            Width           =   3420
         End
         Begin VB.CommandButton btnMute 
            Height          =   285
            Left            =   5385
            Picture         =   "Form2.frx":4CC6
            Style           =   1  'Graphical
            TabIndex        =   171
            TabStop         =   0   'False
            ToolTipText     =   "Mute the playing sound"
            Top             =   0
            Width           =   300
         End
         Begin VB.CheckBox chkPlayVolume 
            Caption         =   "Enable loud volume"
            Height          =   225
            Left            =   1365
            TabIndex        =   163
            ToolTipText     =   "When checked this box enables the louder versions of the sounds to be played"
            Top             =   2445
            Width           =   3405
         End
         Begin VB.CheckBox chkEnableSounds 
            Caption         =   "Enable Sounds for the Animations"
            Height          =   225
            Left            =   1365
            TabIndex        =   161
            ToolTipText     =   "Check this box to enable or disable all of the sounds used during any animation on the main screen."
            Top             =   1350
            Width           =   3405
         End
         Begin VB.ComboBox cmbAlarmSound 
            Height          =   315
            ItemData        =   "Form2.frx":4EF3
            Left            =   1365
            List            =   "Form2.frx":4EF5
            Style           =   2  'Dropdown List
            TabIndex        =   53
            ToolTipText     =   "Choose the alarm sound."
            Top             =   0
            Width           =   2160
         End
         Begin VB.CommandButton btnPlaySound 
            Height          =   285
            Left            =   5055
            Picture         =   "Form2.frx":4EF7
            Style           =   1  'Graphical
            TabIndex        =   52
            TabStop         =   0   'False
            ToolTipText     =   "Play this sound"
            Top             =   0
            Width           =   300
         End
         Begin VB.CommandButton btnSoundsLocation 
            Caption         =   "..."
            Height          =   300
            Left            =   1365
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Open a file explorer at the Sounds folder location."
            Top             =   3780
            Width           =   315
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Quality :"
            Height          =   375
            Index           =   10
            Left            =   645
            TabIndex        =   217
            Tag             =   "lblMicrophone"
            Top             =   5565
            Width           =   660
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "High"
            Height          =   195
            Index           =   9
            Left            =   4620
            TabIndex        =   216
            ToolTipText     =   "This records at 550khz and creates hight quality and large recordings that may fill your drive!"
            Top             =   6015
            Width           =   615
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Low"
            Height          =   195
            Index           =   8
            Left            =   1245
            TabIndex        =   215
            ToolTipText     =   "This captures at 5500khz and creates low quality but small recordings"
            Top             =   6015
            Width           =   615
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Select the audio input device and the recording quality option below."
            Height          =   525
            Index           =   7
            Left            =   1380
            TabIndex        =   213
            Top             =   4965
            Width           =   3615
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Recording :"
            Height          =   375
            Index           =   2
            Left            =   405
            TabIndex        =   212
            Tag             =   "lblMicrophone"
            Top             =   4470
            Width           =   1545
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "When checked this box enables the louder versions of the sounds to be played"
            Height          =   660
            Index           =   5
            Left            =   1410
            TabIndex        =   164
            Tag             =   "lblPlayVolumeDesc"
            Top             =   2805
            Width           =   3615
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "When checked this box enables all the other sounds used during any animation on the main screen."
            Height          =   660
            Index           =   4
            Left            =   1395
            TabIndex        =   162
            Tag             =   "lblEnableSoundsDesc"
            Top             =   1710
            Width           =   3615
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   " Alarm Sound :"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   58
            Tag             =   "lblAlarmSound"
            Top             =   45
            Width           =   1740
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Choose the alarm sound."
            Height          =   300
            Index           =   3
            Left            =   1350
            TabIndex        =   57
            Tag             =   "lblAlarmSoundDesc"
            Top             =   450
            Width           =   1890
         End
         Begin VB.Label lblPlaySound 
            Caption         =   "Play this sound"
            Height          =   300
            Left            =   3825
            TabIndex        =   56
            Top             =   45
            Width           =   1635
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Sounds Location:"
            Height          =   375
            Index           =   1
            Left            =   -15
            TabIndex        =   55
            Tag             =   "lblSoundsLocation"
            Top             =   3780
            Width           =   1425
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Open a file explorer at the sounds folder location."
            Height          =   465
            Index           =   6
            Left            =   1830
            TabIndex        =   54
            Tag             =   "lblSoundsLocationDesc"
            Top             =   3780
            Width           =   2895
         End
      End
   End
   Begin VB.Frame fraWindow 
      Caption         =   "Window"
      Height          =   8205
      Left            =   1170
      TabIndex        =   25
      Top             =   1380
      Width           =   8640
      Begin VB.Frame fraWindowInner 
         BorderStyle     =   0  'None
         Height          =   7500
         Left            =   1050
         TabIndex        =   34
         Top             =   345
         Width           =   6900
         Begin VB.Frame fraIconise 
            BorderStyle     =   0  'None
            Height          =   1470
            Left            =   1065
            TabIndex        =   276
            Top             =   3285
            Width           =   4005
            Begin VB.OptionButton optIconiseDesktop 
               Caption         =   "Iconise to Desktop"
               Height          =   330
               Left            =   270
               TabIndex        =   278
               ToolTipText     =   "Minimise to desktop"
               Top             =   60
               Width           =   2790
            End
            Begin VB.OptionButton optIconiseOpacity 
               Caption         =   "Iconise to Defined Opacity"
               Height          =   330
               Left            =   270
               TabIndex        =   277
               ToolTipText     =   "Fade to a defined opacity"
               Top             =   465
               Width           =   2790
            End
            Begin VB.Label lblOptIconiseOpacity 
               Caption         =   "Select whether to minimise to desktop or opacity."
               Height          =   600
               Left            =   270
               TabIndex        =   279
               Top             =   885
               Width           =   3900
            End
         End
         Begin VB.ComboBox cmbWindowLevel 
            Height          =   315
            Left            =   1305
            Style           =   2  'Dropdown List
            TabIndex        =   38
            ToolTipText     =   $"Form2.frx":5101
            Top             =   0
            Width           =   3960
         End
         Begin VB.CheckBox chkPreventDragging 
            Caption         =   "Ignore Mouse"
            Enabled         =   0   'False
            Height          =   225
            Left            =   1335
            TabIndex        =   36
            ToolTipText     =   "Checking this box turns off the ability to drag the program with the mouse. "
            Top             =   2250
            Width           =   225
         End
         Begin VB.CheckBox chkIgnoreMouse 
            Caption         =   "Ignore Mouse"
            Enabled         =   0   'False
            Height          =   225
            Left            =   1320
            TabIndex        =   35
            ToolTipText     =   "Checking this box causes the program to ignore all mouse events."
            Top             =   1215
            Width           =   225
         End
         Begin vb6projectCCRSlider.Slider sliOpacity 
            Height          =   390
            Left            =   1245
            TabIndex        =   37
            ToolTipText     =   "Set the transparency of the Program."
            Top             =   4635
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   688
            Min             =   20
            Max             =   100
            Value           =   100
            TickFrequency   =   6
            SelStart        =   20
         End
         Begin vb6projectCCRSlider.Slider sliIconiseDelay 
            Height          =   420
            Left            =   1260
            TabIndex        =   267
            ToolTipText     =   "Choose the delay (seconds) before auto-iconisation occurs. Set to 0 to disable,"
            Top             =   6060
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   741
            Max             =   600
            Value           =   100
            TickFrequency   =   30
            SelStart        =   20
         End
         Begin VB.Label lblIconiseDelay450 
            Caption         =   "450"
            Height          =   345
            Left            =   3930
            TabIndex        =   274
            Top             =   6570
            Width           =   495
         End
         Begin VB.Label lblIconiseDelay150 
            Caption         =   "150"
            Height          =   345
            Left            =   2205
            TabIndex        =   273
            Top             =   6570
            Width           =   345
         End
         Begin VB.Label lblIconiseDelay0 
            Caption         =   "0"
            Height          =   345
            Left            =   1395
            TabIndex        =   272
            Top             =   6570
            Width           =   345
         End
         Begin VB.Label lblIconiseDelay600 
            Caption         =   "600"
            Height          =   345
            Left            =   4740
            TabIndex        =   271
            Top             =   6570
            Width           =   405
         End
         Begin VB.Label lblIconiseDelay300 
            Caption         =   "300"
            Height          =   345
            Left            =   3075
            TabIndex        =   270
            Top             =   6570
            Width           =   840
         End
         Begin VB.Label lblConfigurationTab 
            Caption         =   "Iconise Delay :"
            Height          =   345
            Index           =   3
            Left            =   -15
            TabIndex        =   269
            Tag             =   "lblIconiseDelay"
            Top             =   6135
            Width           =   1800
         End
         Begin VB.Label lblConfigurationTab 
            Caption         =   "Choose the delay (seconds) before auto-iconisation occurs. Set to 0 to disable,"
            Height          =   360
            Index           =   9
            Left            =   1365
            TabIndex        =   268
            Top             =   6900
            Width           =   3810
         End
         Begin VB.Label lblIgnoreMouse 
            Caption         =   "Ignore Mouse"
            Enabled         =   0   'False
            Height          =   270
            Left            =   1650
            TabIndex        =   39
            Top             =   1215
            Width           =   1725
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Window Level"
            Height          =   345
            Left            =   0
            TabIndex        =   49
            Top             =   60
            Width           =   1740
         End
         Begin VB.Label lblOpacity20 
            Caption         =   "20%"
            Height          =   315
            Left            =   1335
            TabIndex        =   48
            Top             =   5145
            Width           =   345
         End
         Begin VB.Label lblOpacityLabel100 
            Caption         =   "100%"
            Height          =   315
            Left            =   4695
            TabIndex        =   47
            Top             =   5145
            Width           =   405
         End
         Begin VB.Label lblOpacityText 
            Caption         =   "Opacity"
            Height          =   315
            Left            =   2820
            TabIndex        =   46
            Top             =   5145
            Width           =   840
         End
         Begin VB.Label lblOpacityLabel 
            Caption         =   "Opacity:"
            Height          =   315
            Left            =   600
            TabIndex        =   45
            Top             =   4695
            Width           =   780
         End
         Begin VB.Label lblWindowLevelDescription 
            Caption         =   $"Form2.frx":519C
            Height          =   870
            Left            =   1365
            TabIndex        =   44
            Top             =   450
            Width           =   3930
         End
         Begin VB.Label lblOpacityLabelDesc 
            Caption         =   "Set the program transparency level."
            Height          =   330
            Left            =   1380
            TabIndex        =   43
            Top             =   5460
            Width           =   3810
         End
         Begin VB.Label lblPreventDraggingText 
            Caption         =   "Checking this box turns off the ability to drag the program with the mouse. "
            Enabled         =   0   'False
            Height          =   600
            Left            =   1335
            TabIndex        =   42
            Top             =   2625
            Width           =   3900
         End
         Begin VB.Label lblPreventDragging 
            Caption         =   "Prevent Dragging"
            Enabled         =   0   'False
            Height          =   270
            Left            =   1665
            TabIndex        =   41
            Top             =   2250
            Width           =   1725
         End
         Begin VB.Label lblIgnoreMouseText 
            Caption         =   "Checking this box causes the program to ignore all mouse events."
            Enabled         =   0   'False
            Height          =   660
            Left            =   1320
            TabIndex        =   40
            Top             =   1590
            Width           =   3810
         End
      End
   End
   Begin VB.Menu prefsMnuPopmenu 
      Caption         =   "The main menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAboutFireCallWin 
         Caption         =   "About Fire Call Win"
      End
      Begin VB.Menu blank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenSharedInputFile 
         Caption         =   "Open the Shared Input File"
      End
      Begin VB.Menu mnuOpenSharedOutputFile 
         Caption         =   "Open the Shared Output File"
      End
      Begin VB.Menu mnuOpenSharedExchangeFolder 
         Caption         =   "Open the Shared Exchange Folder"
      End
      Begin VB.Menu blank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCoffee 
         Caption         =   "Donate a coffee with paypal"
      End
      Begin VB.Menu mnuSweets 
         Caption         =   "Donate some sweets/candy with Amazon"
      End
      Begin VB.Menu mnuSupport 
         Caption         =   "Contact Support"
      End
      Begin VB.Menu blank3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuButton 
         Caption         =   "Theme Colours"
         Begin VB.Menu mnuLight 
            Caption         =   "Light Theme Enable"
         End
         Begin VB.Menu mnuDark 
            Caption         =   "High Contrast Theme Enable"
         End
         Begin VB.Menu mnuAuto 
            Caption         =   "Auto Theme Selection"
         End
      End
      Begin VB.Menu mnuLicenceA 
         Caption         =   "Display Licence Agreement"
      End
      Begin VB.Menu mnuClosePreferences 
         Caption         =   "Close Preferences"
      End
   End
End
Attribute VB_Name = "FireCallPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : FireCallPrefs
' Author    : beededea
' Date      : 17/08/2021
' Purpose   :
'---------------------------------------------------------------------------------------

'@ModuleAttribute VB_Name, "FireCallPrefs"
'@ModuleAttribute VB_Creatable, False
'@PredeclaredId
'@ModuleAttribute VB_GlobalNameSpace, False
'@ModuleAttribute VB_Exposed, False
Option Explicit

Private txtStringToAddFieldModified As Boolean

Private Const MODULE_NAME As String = "FireCallPrefs"

Private WithEvents m_oProxy As cSmtpProxy
Attribute m_oProxy.VB_VarHelpID = -1

'---------------------------------------------------------------------------------------
' Procedure : sendEmailPrefs
' Author    : beededea
' Date      : 29/01/2022
' Purpose   : This is a duplicate of sendEmailMain, the reason it is duplicated rather than dropped into
'             a shared module is due to the withEvents clause on m_oProxy. Events are only generated by forms.
'             I have yet to extract this code and make it operate through the use of a class but this is not yet done.
'
' STARTTLS is an email protocol command that tells an email server that an email client,
' including an email client running in a web browser, wants to turn an existing insecure connection
' into a secure one. We use a proxy to inject that command into the CDO stream by diverting the stream
' from the desired port to the LNG_PROXY_PORT where our proxy is ready to take over.
'
'---------------------------------------------------------------------------------------
'
Private Function sendEmailPrefs(ByVal strSender As String, _
                        ByVal strRecipient As String, _
                        ByVal strSubject As String, _
                        ByVal strBody As String, _
                        Optional ByVal strCc As String, _
                        Optional ByVal strBcc As String, _
                        Optional ByVal colAttachments As Collection _
                        ) As Boolean

    Dim cdoMsg As New CDO.Message
    Dim cdoConf As New CDO.Configuration
    Dim schema As String
    Dim Flds
    Dim attachment
    Dim securityStr As String
    Dim decryptedPassword As String

    On Error GoTo sendEmailPrefs_Error

    Const cdoSendUsingPort = 2
    securityStr = ""

    Set cdoMsg = CreateObject("CDO.Message")
    Set cdoConf = CreateObject("CDO.Configuration")

    Set Flds = cdoConf.Fields

    schema = "http://schemas.microsoft.com/cdo/configuration/"

    With Flds
        .Item(schema & "smtpconnectiontimeout") = 30
        .Item(schema & "sendusing") = 2 ' SMTP over the network = 2, set Local SMTP = 1
                
        If FCWSmtpSecurity = "0" Then
            .Item(schema & "smtpserverport") = Val(FCWSmtpPort) '465
            .Item(schema & "smtpserver") = FCWSmtpServer ' eg. smtp.gmail.com
            securityStr = " SMTP Security = NONE "
        End If
    
        If FCWSmtpSecurity = "1" Then
            '.Item(schema & "sendtls") = True ' I am sure this value does nothing
            .Item(schema & "smtpserver") = "127.0.0.1" 'localhost
            .Item(schema & "smtpserverport") = LNG_PROXY_PORT
            securityStr = " SMTP Security STARTTLS=true "
        End If

        If FCWSmtpSecurity = "2" Then
            .Item(schema & "smtpserverport") = Val(FCWSmtpPort) '25, 465 &c
            .Item(schema & "smtpusessl") = True
            .Item(schema & "smtpserver") = FCWSmtpServer
            securityStr = " SMTP Security SSL=true "
        End If
        
        .Item(schema & "smtpauthenticate") = Val(FCWSmtpAuthenticate) ' 0 - None  1 - Base 64 encoded (Normal)    2 - NTLM
        .Item(schema & "sendusername") = FCWSmtpUsername '"your email@gmail.com"
                
'        Dim a As String
'        a = decryptstr(FCWSmtpPassword)
        decryptedPassword = AesDecryptString(FCWSmtpPassword, emailTString)

        .Item(schema & "sendpassword") = decryptedPassword '"your password"
        .Update
    End With
    
    If FireCallPrefs.chkAppendConfig.Value = 1 Then
        securityStr = " SMTP server " & FCWSmtpServer & securityStr
        securityStr = securityStr & " Port:" & Val(FCWSmtpPort) & " Authentication Method:" & FireCallPrefs.cmbSmtpAuthenticate.List(Val(FCWSmtpAuthenticate))
        strSubject = strSubject & securityStr
        strBody = strBody & securityStr
    End If
    

    ' Apply the settings to the message.
    With cdoMsg
        Set .Configuration = cdoConf
        .To = strRecipient
        .from = FCWSmtpUsername
        .Subject = strSubject
        .TextBody = strBody
        If Not colAttachments Is Nothing Then
            For Each attachment In colAttachments
                .AddAttachment attachment
            Next
        End If
        If strCc <> "" Then .CC = strCc
        If strBcc <> "" Then .BCC = strBcc
        .Send
    End With

    Set cdoMsg = Nothing
    Set cdoConf = Nothing
    Set Flds = Nothing

    sendEmailPrefs = True

    On Error GoTo 0
    Exit Function

sendEmailPrefs_Error:
    sendEmailPrefs = False
    With err
         If .Number <> 0 Then
            debugLog "Error " & err.Number & " (" & err.Description & ") in procedure sendEmailPrefs of Form FireCallPrefs"
            Resume Next
          End If
    End With

End Function

'---------------------------------------------------------------------------------------
' Procedure : btnArchiveNow_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnArchiveNow_Click()
   On Error GoTo btnArchiveNow_Click_Error

    Call houseKeepingTimerLogic(True)

    MsgBox ("Archive command issued")
   
   On Error GoTo 0
   Exit Sub

btnArchiveNow_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnArchiveNow_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnBackupNow_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnBackupNow_Click()
   On Error GoTo btnBackupNow_Click_Error

    Call backupOutputFile(FCWSharedOutputFile, "manual")
    
    MsgBox ("Backup command issued")

   On Error GoTo 0
   Exit Sub

btnBackupNow_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnBackupNow_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkEnableAlarmSound_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkEnableAlarmSound_Click()
   On Error GoTo chkEnableAlarmSound_Click_Error

    btnSave.Enabled = True ' enable the save button
    FCWEnableAlarmSound = LTrim$(Str$(chkEnableAlarmSound.Value))

   On Error GoTo 0
   Exit Sub

chkEnableAlarmSound_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure chkEnableAlarmSound_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkGenStartup_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkGenStartup_Click()
   On Error GoTo chkGenStartup_Click_Error

    btnSave.Enabled = True ' enable the save button
    
    FCWStartup = LTrim$(Str$(chkGenStartup.Value))

   On Error GoTo 0
   Exit Sub

chkGenStartup_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure chkGenStartup_Click of Form FireCallPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkServiceProcesses_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkServiceProcesses_Click()
   On Error GoTo chkServiceProcesses_Click_Error

    btnSave.Enabled = True ' enable the save button
    
    FCWCheckServiceProcesses = LTrim$(Str$(chkServiceProcesses.Value))

   On Error GoTo 0
   Exit Sub

chkServiceProcesses_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure chkServiceProcesses_Click of Form FireCallPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : optIconiseDesktop_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub optIconiseDesktop_Click()
   On Error GoTo optIconiseDesktop_Click_Error

    btnSave.Enabled = True ' enable the save button
    Call checkIconiseOpacityLevel

   On Error GoTo 0
   Exit Sub

optIconiseDesktop_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure optIconiseDesktop_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : optIconiseOpacity_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub optIconiseOpacity_Click()
   On Error GoTo optIconiseOpacity_Click_Error

    btnSave.Enabled = True ' enable the save button
    
    Call checkIconiseOpacityLevel

   On Error GoTo 0
   Exit Sub

optIconiseOpacity_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure optIconiseOpacity_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkPreventDragging_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkPreventDragging_Click()
   On Error GoTo chkPreventDragging_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkPreventDragging_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure chkPreventDragging_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbArchiveDays_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbArchiveDays_Click()
   On Error GoTo cmbArchiveDays_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

cmbArchiveDays_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmbArchiveDays_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : m_oProxy_RecvFromClient
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub m_oProxy_RecvFromClient(Data() As Byte)
    Dim sText           As String
    
   On Error GoTo m_oProxy_RecvFromClient_Error

    sText = StrConv(Data, vbUnicode)
    If Right$(sText, 2) = vbCrLf Then
        sText = Left$(sText, Len(sText) - 2)
    End If
    pvLog "->" & Replace(sText, vbCrLf, vbCrLf & "  ")

   On Error GoTo 0
   Exit Sub

m_oProxy_RecvFromClient_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure m_oProxy_RecvFromClient of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : m_oProxy_RecvFromServer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub m_oProxy_RecvFromServer(Data() As Byte)
    Dim sText           As String
    
   On Error GoTo m_oProxy_RecvFromServer_Error

    sText = StrConv(Data, vbUnicode)
    If Right$(sText, 2) = vbCrLf Then
        sText = Left$(sText, Len(sText) - 2)
    End If
    pvLog "<-" & Replace(sText, vbCrLf, vbCrLf & "  ")

   On Error GoTo 0
   Exit Sub

m_oProxy_RecvFromServer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure m_oProxy_RecvFromServer of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : pvLog
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub pvLog(sText As String)
   On Error GoTo pvLog_Error

    txtEmailLog.SelStart = &H7FFF
    txtEmailLog.SelText = sText & vbCrLf
    txtEmailLog.SelStart = &H7FFF

   On Error GoTo 0
   Exit Sub

pvLog_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure pvLog of Form FireCallPrefs"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : cmbCaptureDevices_KeyPress
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbCaptureDevices_KeyPress(KeyAscii As Integer)
   On Error GoTo cmbCaptureDevices_KeyPress_Error

    KeyAscii = 0

   On Error GoTo 0
   Exit Sub

cmbCaptureDevices_KeyPress_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmbCaptureDevices_KeyPress of Form FireCallPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : btnBackupLocation_Click
' Author    : beededea
' Date      : 17/08/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnBackupLocation_Click()
    ' On Error GoTo btnBackupLocation_Click_Error

        If fDirExists(App.Path & "\Resources\sounds\") Then
            Call ShellExecute(Me.hwnd, "Open", FCWBackupFolder, vbNullString, App.Path, 1)
        End If

    On Error GoTo 0
    Exit Sub

btnBackupLocation_Click_Error:

    With err
         If .Number <> 0 Then
            debugLog "Error " & err.Number & " (" & err.Description & ") in procedure btnBackupLocation_Click of Form FireCallPrefs"
            Resume Next
          End If
    End With
End Sub





'---------------------------------------------------------------------------------------
' Procedure : btnSeePassword_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnSeePassword_Click()
   On Error GoTo btnSeePassword_Click_Error

    If txtSMTPNoPassword.Visible = False Then
        txtSMTPNoPassword.Visible = True
        txtSmtpPassword.Visible = False
    Else
        txtSMTPNoPassword.Visible = False
        txtSmtpPassword.Visible = True
    End If

   On Error GoTo 0
   Exit Sub

btnSeePassword_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnSeePassword_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnTestEmail_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnTestEmail_Click()
    Dim a As Boolean
    
   On Error GoTo btnTestEmail_Click_Error

    MsgBox "Test email message sent. Error from the server, if any, should appear within 30 seconds. Check your Email and press get new messages!"
    
    'if the starttls option is selected then do this
    If FCWSmtpSecurity = 1 Then ' STARTTLS
        Set m_oProxy = New cSmtpProxy
        If m_oProxy.Init(FCWSmtpServer, FCWSmtpPort, LNG_PROXY_PORT) Then
            pvLog "SMTP proxy listening on " & LNG_PROXY_PORT
        End If
'
        fraEmailfra.Visible = True
        fraEmailfra.Left = 1410
        fraEmailfra.Top = 3000
        btnCloseEmailFra.Enabled = True
    End If
    
    a = sendEmailPrefs(txtRecipientEmail.Text, _
                        txtRecipientEmail.Text, _
                        txtEmailSubject.Text, _
                        txtEmailMessage.Text)

   On Error GoTo 0
   Exit Sub

btnTestEmail_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnTestEmail_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnTestEmail_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnTestEmail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo btnTestEmail_MouseMove_Error

    CreateToolTip btnTestEmail.hwnd, "Error messages will only appear 30 secs after the button is pressed. A success can only be checked by viewing the email client to see if an email has arrived. Please note that STARTTLS on port 587 is not currently supported. Port 25 and SSL is tested and operates successfully on Hotmail.", _
                  TTIconInfo, "Help on Testing Email", , , , True

   On Error GoTo 0
   Exit Sub

btnTestEmail_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnTestEmail_MouseMove of Form FireCallPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkSendEmails_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkSendEmails_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo chkSendEmails_MouseMove_Error

    CreateToolTip chkSendEmails.hwnd, "Messages are sent by email using the SMTP details entered.  Extract these from your email client, Outlook or Thunderbird for example.", _
                  TTIconInfo, "Help on Advice Messages", , , , True

   On Error GoTo 0
   Exit Sub

chkSendEmails_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure chkSendEmails_MouseMove of Form FireCallPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkSendErrorEmails_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkSendErrorEmails_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo chkSendErrorEmails_MouseMove_Error

    CreateToolTip chkSendErrorEmails.hwnd, "Error messages are sent when an error is generated as long as FCW is still running. Messages are sent by email using the SMTP details entered below.", _
                  TTIconInfo, "Help on Error Messages", , , , True

   On Error GoTo 0
   Exit Sub

chkSendErrorEmails_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure chkSendErrorEmails_MouseMove of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbCaptureDevices_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbCaptureDevices_Click()
   On Error GoTo cmbCaptureDevices_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

cmbCaptureDevices_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmbCaptureDevices_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkAllowShutdowns_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkAllowShutdowns_Click()
   On Error GoTo chkAllowShutdowns_Click_Error

    btnSave.Enabled = True ' enable the save button
    FCWAllowShutdowns = LTrim$(Str$(chkAllowShutdowns.Value))

   On Error GoTo 0
   Exit Sub

chkAllowShutdowns_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure chkAllowShutdowns_Click of Form FireCallPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkAllowShutdowns_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkAllowShutdowns_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo chkAllowShutdowns_MouseMove_Error

    CreateToolTip chkAllowShutdowns.hwnd, "This determines whether your remote partner has the capability of issuing shutdown requests to your copy of Fire Call prior to him performing administration or housekeeping tasks such as reducing the size of the text files used to store the chat text. If these tasks are carried out whilst FCW is running it could cause the app some problems. Having the ability to indicate the need for a shutdown to your partner is a useful tool. This is really only needed if your chat partner performs the housekeeping tasks manually.", _
                  TTIconInfo, "Help on Remote Shutdown Requests", , , , True

   On Error GoTo 0
   Exit Sub

chkAllowShutdowns_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure chkAllowShutdowns_MouseMove of Form FireCallPrefs"
End Sub







'---------------------------------------------------------------------------------------
' Procedure : cmbSmtpAuthenticate_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbSmtpAuthenticate_Click()
    'smtpauthenticate Type of Authenthication
    '0 - None
    '1 - Base 64 encoded (Normal)
    '2 - NTLM
   On Error GoTo cmbSmtpAuthenticate_Click_Error

    btnTestEmail.Enabled = False
    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

cmbSmtpAuthenticate_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmbSmtpAuthenticate_Click of Form FireCallPrefs"
End Sub







'---------------------------------------------------------------------------------------
' Procedure : cmbSmtpConfig_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbSmtpConfig_Click()
   On Error GoTo cmbSmtpConfig_Click_Error

    btnSave.Enabled = True ' enable the save button
    ' read the listindex value
    ' open the settings file and read the specific settings chosen
    'cmbSmtpConfig.Text = cmbSmtpConfig.List(cmbSmtpConfig.ListIndex)

    Call readSmtpConfigDetails("Software\FireCallWin", cmbSmtpConfig.ListIndex)
    Call adjustPrefsSmtpControls
    
    btnTestEmail.Enabled = True

   On Error GoTo 0
   Exit Sub

cmbSmtpConfig_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmbSmtpConfig_Click of Form FireCallPrefs"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : adjustPrefsSmtpControls
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub adjustPrefsSmtpControls()
    
   On Error GoTo adjustPrefsSmtpControls_Error

    txtSmtpServer.Text = FCWSmtpServer
    txtSmtpUsername.Text = FCWSmtpUsername
    txtSmtpPassword.Text = AesDecryptString(FCWSmtpPassword, emailTString)
    
    If txtSmtpPassword.Text = "" Then
        txtSMTPNoPassword.Visible = False
        txtSmtpPassword.Visible = True
    End If
    
    txtSmtpConfigName.Text = FCWSmtpConfigName
    
    txtSmtpPort.Text = FCWSmtpPort
    cmbSmtpAuthenticate.ListIndex = Val(FCWSmtpAuthenticate) 'nnn
    cmbSmtpSecurity.ListIndex = Val(FCWSmtpSecurity) 'nnn
    
    txtRecipientEmail.Text = FCWRecipientEmail
    txtEmailSubject.Text = FCWEmailSubject
    txtEmailMessage.Text = FCWEmailMessage

   On Error GoTo 0
   Exit Sub

adjustPrefsSmtpControls_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure adjustPrefsSmtpControls of Form FireCallPrefs"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : cmbSmtpSecurity_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbSmtpSecurity_Click()
   On Error GoTo cmbSmtpSecurity_Click_Error

    btnTestEmail.Enabled = False
    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

cmbSmtpSecurity_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmbSmtpSecurity_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnCloseEmailFra_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnCloseEmailFra_Click()
   On Error GoTo btnCloseEmailFra_Click_Error

    fraEmailfra.Visible = False

   On Error GoTo 0
   Exit Sub

btnCloseEmailFra_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnCloseEmailFra_Click of Form FireCallPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : Command2_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Command2_Click()
   On Error GoTo Command2_Click_Error

    txtEmailLog.Text = ""

   On Error GoTo 0
   Exit Sub

Command2_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure Command2_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraAllowShutdowns_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraAllowShutdowns_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraAllowShutdowns_MouseMove_Error

    CreateToolTip fraAllowShutdowns.hwnd, "This determines whether your remote partner has the capability of issuing shutdown requests to your copy of Fire Call prior to him performing administration or housekeeping tasks such as reducing the size of the text files used to store the chat text. If these tasks are carried out whilst FCW is running it could cause the app some problems. Having the ability to indicate the need for a shutdown to your partner is a useful tool. This is really only needed if your chat partner performs the housekeeping tasks manually.", _
                  TTIconInfo, "Help on Remote Shutdown Requests", , , , True

   On Error GoTo 0
   Exit Sub

fraAllowShutdowns_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraAllowShutdowns_MouseMove of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkAutomaticBackups_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkAutomaticBackups_Click()
   On Error GoTo chkAutomaticBackups_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkAutomaticBackups_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure chkAutomaticBackups_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkAutomaticHousekeeping_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkAutomaticHousekeeping_Click()
   On Error GoTo chkAutomaticHousekeeping_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkAutomaticHousekeeping_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure chkAutomaticHousekeeping_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkBackupOnStart_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkBackupOnStart_Click()
   On Error GoTo chkBackupOnStart_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkBackupOnStart_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure chkBackupOnStart_Click of Form FireCallPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : chkEnableBalloonTooltips_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkEnableBalloonTooltips_Click()
   On Error GoTo chkEnableBalloonTooltips_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkEnableBalloonTooltips_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure chkEnableBalloonTooltips_Click of Form FireCallPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Command1_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Command1_Click()
   On Error GoTo Command1_Click_Error

        If fDirExists(FCWArchiveFolder) Then
            Call ShellExecute(Me.hwnd, "Open", FCWArchiveFolder, vbNullString, App.Path, 1)
        End If

   On Error GoTo 0
   Exit Sub

Command1_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure Command1_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

   On Error GoTo Form_Load_Error

    Set Bas64 = New Base64
    
    txtStringToAddFieldModified = False
    txtStringToAdd.Text = "Enter text here and click + button below"
          
    ' size and position the frames and buttons
    Call positionThings
    
    ' populate all the comboboxes in the prefs form
    Call populateComboBoxes
    
    ' adjust all the preferences and main program controls
    Call adjustPrefsControls
    
    ' check to see if the TEST button should be visible
    Call testEmailTestButton
    
    If FCWSkinTheme <> "" Then
        If FCWSkinTheme = "dark" Then
            Call setThemeShade(212, 208, 199)
        Else
            Call setThemeShade(240, 240, 240)
        End If
    Else
        If classicThemeCapable = True Then
            mnuAuto.Caption = "Auto Theme Enabled - Click to Disable"
            FireCallPrefs.themeTimer.Enabled = True
        Else
            Call setModernThemeColours
        End If
    End If
    
    ' make the last used tab appear on startup
    Call showLastTab
    
    btnSave.Enabled = False ' disable the save button

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure Form_Load of Form FireCallPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : showLastTab
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub showLastTab()

    ' make the last used tab appear on startup
   On Error GoTo showLastTab_Error

    If FCWLastSelectedTab = "general" Then Call picGeneral_Click
    If FCWLastSelectedTab = "config" Then Call picConfig_Click
    If FCWLastSelectedTab = "email" Then Call picEmail_Click
    If FCWLastSelectedTab = "emojis" Then Call picEmojis_Click
    If FCWLastSelectedTab = "fonts" Then Call picFonts_Click
    If FCWLastSelectedTab = "texts" Then Call picTexts_Click
    If FCWLastSelectedTab = "sounds" Then Call picSounds_Click
    If FCWLastSelectedTab = "housekeeping" Then Call picHouseKeeping_Click
    If FCWLastSelectedTab = "window" Then Call picWindow_Click

   On Error GoTo 0
   Exit Sub

showLastTab_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure showLastTab of Form FireCallPrefs"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : positionThings
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub positionThings()

    Dim frameWidth As Integer
    
    ' size and position the frames and buttons
    
   On Error GoTo positionThings_Error

    Me.Width = 9000
    
    fraGeneral.Top = 1050
    fraConfiguration.Top = 1050
    fraEmail.Top = 1050
    fraEmojis.Top = 1050
    fraFonts.Top = 1050
    fraWindow.Top = 1050
    fraTexts.Top = 1050
    fraHousekeeping.Top = 1050
    fraSounds.Top = 1050
    
    fraGeneral.Left = 120
    fraConfiguration.Left = 120
    fraEmail.Left = 120
    fraEmojis.Left = 120
    fraFonts.Left = 120
    fraWindow.Left = 120
    fraTexts.Left = 120
    fraHousekeeping.Left = 120
    fraSounds.Left = 120
    
    frameWidth = 8700
    fraGeneral.Width = frameWidth
    fraConfiguration.Width = frameWidth
    fraEmail.Width = frameWidth
    fraEmojis.Width = frameWidth
    fraFonts.Width = frameWidth
    fraWindow.Width = frameWidth
    fraTexts.Width = frameWidth
    fraHousekeeping.Width = frameWidth
    fraSounds.Width = frameWidth
    
    fraGeneralButtons.BorderStyle = 1
    
    fraGeneral.Visible = True
    fraConfiguration.Visible = False
    fraEmail.Visible = False
    fraEmojis.Visible = False
    fraFonts.Visible = False
    fraWindow.Visible = False
    fraTexts.Visible = False
    fraHousekeeping.Visible = False
    fraSounds.Visible = False
    
    FireCallPrefs.Height = fraGeneral.Height + 2000
    btnSave.Top = fraGeneral.Top + fraGeneral.Height + 100
    btnCancel.Top = fraGeneral.Top + fraGeneral.Height + 100
    btnHelp.Top = fraGeneral.Top + fraGeneral.Height + 100
    
    txtSMTPNoPassword.Left = txtSmtpPassword.Left
    txtSMTPNoPassword.Top = txtSmtpPassword.Top

   On Error GoTo 0
   Exit Sub

positionThings_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure positionThings of Form FireCallPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : toggleAllEmailControls
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function toggleAllEmailControls(setting As String) As String
Dim ctrl As Control
    Dim a As String
    
   On Error GoTo toggleAllEmailControls_Error

    If setting = "hide" Then
        toggleAllEmailControls = "hidden"
        
        For Each ctrl In Me.Controls
            a = ctrl.Name
            
            If (TypeOf ctrl Is CommandButton) Or (TypeOf ctrl Is TextBox) Or (TypeOf ctrl Is FileListBox) Or (TypeOf ctrl Is Label) Or (TypeOf ctrl Is ComboBox) Or (TypeOf ctrl Is CheckBox) Or (TypeOf ctrl Is OptionButton) Or (TypeOf ctrl Is Frame) Or (TypeOf ctrl Is ListBox) Then
                If ctrl.Container.Name = "fraEmailInner" Then
                    ctrl.Enabled = False
                End If
            End If
        Next
        
        txtSmtpServer.Enabled = False
        lblEmailTab(1).Enabled = False
    Else
        toggleAllEmailControls = "shown"
        
        For Each ctrl In Me.Controls
            a = ctrl.Name
            
            If (TypeOf ctrl Is CommandButton) Or (TypeOf ctrl Is TextBox) Or (TypeOf ctrl Is FileListBox) Or (TypeOf ctrl Is Label) Or (TypeOf ctrl Is ComboBox) Or (TypeOf ctrl Is CheckBox) Or (TypeOf ctrl Is OptionButton) Or (TypeOf ctrl Is Frame) Or (TypeOf ctrl Is ListBox) Then
                If ctrl.Container.Name = "fraEmailInner" Then
                    ctrl.Enabled = True
                End If
            End If
        Next
        txtSmtpServer.Enabled = True
        lblEmailTab(1).Enabled = True
    End If
    
    
    chkSendEmails.Enabled = True
    chkSendErrorEmails.Enabled = True
    'lblEmailsDesc.Enabled = True
    toggleAllEmailControls = True

   On Error GoTo 0
   Exit Function

toggleAllEmailControls_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure toggleAllEmailControls of Form FireCallPrefs"
    
End Function

' add new user defined text to the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnBusyAdd_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnBusyAdd_Click()
   On Error GoTo btnBusyAdd_Click_Error

    If txtStringToAddFieldModified = True And txtStringToAdd.Text <> vbNullString Then
        cmbBusy.AddItem txtStringToAdd.Text, 0
    Else
        MsgBox "No text to add - please add your text in the box above and then retry."
    End If
    cmbBusy.ListIndex = 0

   On Error GoTo 0
   Exit Sub

btnBusyAdd_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnBusyAdd_Click of Form FireCallPrefs"

End Sub
' remove user defined text from the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnBusyRemove_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnBusyRemove_Click()
   On Error GoTo btnBusyRemove_Click_Error

    btnSave.Enabled = True ' enable the save button
    If cmbBusy.ListIndex <> 0 Then
        cmbBusy.RemoveItem (cmbBusy.ListIndex)
    Else
        MsgBox "You cannot delete the first item in the list, try one of the others. Note: you can always add one at the top and replace the one you want to remove."
    End If
        
    cmbBusy.ListIndex = 0

   On Error GoTo 0
   Exit Sub

btnBusyRemove_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnBusyRemove_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnCancel_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnCancel_Click()
   On Error GoTo btnCancel_Click_Error

    btnSave.Enabled = False ' disable the save button
    FireCallPrefs.themeTimer.Enabled = False
    Call startThePollingTimers

    Unload Me

   On Error GoTo 0
   Exit Sub

btnCancel_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnCancel_Click of Form FireCallPrefs"
End Sub
' add new user defined text to the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnCodAdd_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnCodAdd_Click()
   On Error GoTo btnCodAdd_Click_Error

    If txtStringToAddFieldModified = True And txtStringToAdd.Text <> vbNullString Then
        cmbCod.AddItem txtStringToAdd.Text, 0
    Else
        MsgBox "No text to add - please add your text in the box above and then retry."
    End If
    cmbCod.ListIndex = 0

   On Error GoTo 0
   Exit Sub

btnCodAdd_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnCodAdd_Click of Form FireCallPrefs"

End Sub
' remove user defined text from the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnCodRemove_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnCodRemove_Click()
   On Error GoTo btnCodRemove_Click_Error

    btnSave.Enabled = True ' enable the save button ' enable the save button
    If cmbCod.ListIndex <> 0 Then
        cmbCod.RemoveItem (cmbCod.ListIndex)
    Else
        MsgBox "You cannot delete the first item in the list, try one of the others. Note: you can always add one at the top and replace the one you want to remove."
    End If
    cmbCod.ListIndex = 0

   On Error GoTo 0
   Exit Sub

btnCodRemove_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnCodRemove_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnDeleteText_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnDeleteText_Click()
   On Error GoTo btnDeleteText_Click_Error

    txtStringToAdd.Text = vbNullString

   On Error GoTo 0
   Exit Sub

btnDeleteText_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnDeleteText_Click of Form FireCallPrefs"
End Sub
' add new user defined text to the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnGdnAdd_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnGdnAdd_Click()
   On Error GoTo btnGdnAdd_Click_Error

    If txtStringToAddFieldModified = True And txtStringToAdd.Text <> vbNullString Then
        cmbGdn.AddItem txtStringToAdd.Text, 0
    Else
        MsgBox "No text to add - please add your text in the box above and then retry."
    End If
    cmbGdn.ListIndex = 0

   On Error GoTo 0
   Exit Sub

btnGdnAdd_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnGdnAdd_Click of Form FireCallPrefs"

End Sub
' remove user defined text from the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnGdnRemove_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnGdnRemove_Click()
   On Error GoTo btnGdnRemove_Click_Error

    btnSave.Enabled = True ' enable the save button ' enable the save button
    If cmbGdn.ListIndex <> 0 Then
        cmbGdn.RemoveItem (cmbGdn.ListIndex)
    Else
        MsgBox "You cannot delete the first item in the list, try one of the others. Note: you can always add one at the top and replace the one you want to remove."
    End If
    cmbGdn.ListIndex = 0

   On Error GoTo 0
   Exit Sub

btnGdnRemove_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnGdnRemove_Click of Form FireCallPrefs"
End Sub
' display the help file
'---------------------------------------------------------------------------------------
' Procedure : btnHelp_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnHelp_Click()
    ' variables declared
    Dim answer As VbMsgBoxResult
    
    'initialise the dimensioned variables
   On Error GoTo btnHelp_Click_Error

    answer = vbNo
    
    answer = MsgBox("This option opens a browser window and displays this tool's help. Proceed?", vbExclamation + vbYesNo)
    If answer = vbYes Then
        If fFExists(App.Path & "\help\FireCallWin Help.html") Then
            Call ShellExecute(Me.hwnd, "Open", App.Path & "\help\FireCallWin Help.html", vbNullString, App.Path, 1)
        Else
            MsgBox ("%Err-I-ErrorNumber 11 - The help file - FireCallWin Help.html - is missing from the help folder.")
        End If
    End If

   On Error GoTo 0
   Exit Sub

btnHelp_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnHelp_Click of Form FireCallPrefs"
End Sub
' remove user defined text from the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnMornRemove_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnMornRemove_Click()
   On Error GoTo btnMornRemove_Click_Error

    btnSave.Enabled = True ' enable the save button ' enable the save button
    If cmbMorn.ListIndex <> 0 Then
        cmbMorn.RemoveItem (cmbMorn.ListIndex)
    Else
        MsgBox "You cannot delete the first item in the list, try one of the others. Note: you can always add one at the top and replace the one you want to remove."
    End If
    cmbMorn.ListIndex = 0

   On Error GoTo 0
   Exit Sub

btnMornRemove_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnMornRemove_Click of Form FireCallPrefs"
End Sub
' remove user defined text from the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnNewsRemove_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnNewsRemove_Click()
   On Error GoTo btnNewsRemove_Click_Error

    btnSave.Enabled = True ' enable the save button ' enable the save button
    If cmbNews.ListIndex <> 0 Then
        cmbNews.RemoveItem (cmbNews.ListIndex)
    Else
        MsgBox "You cannot delete the first item in the list, try one of the others. Note: you can always add one at the top and replace the one you want to remove."
    End If
    cmbNews.ListIndex = 0

   On Error GoTo 0
   Exit Sub

btnNewsRemove_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnNewsRemove_Click of Form FireCallPrefs"
End Sub
' add new user defined text to the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnOutAdd_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnOutAdd_Click()
   On Error GoTo btnOutAdd_Click_Error

    If txtStringToAddFieldModified = True And txtStringToAdd.Text <> vbNullString Then
        cmbOut.AddItem txtStringToAdd.Text, 0
    Else
        MsgBox "No text to add - please add your text in the box above and then retry."
    End If
    cmbOut.ListIndex = 0

   On Error GoTo 0
   Exit Sub

btnOutAdd_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnOutAdd_Click of Form FireCallPrefs"
End Sub
' remove user defined text from the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnOutRemove_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnOutRemove_Click()
   On Error GoTo btnOutRemove_Click_Error

    btnSave.Enabled = True ' enable the save button ' enable the save button
    If cmbOut.ListIndex <> 0 Then
        cmbOut.RemoveItem (cmbOut.ListIndex)
    Else
        MsgBox "You cannot delete the first item in the list, try one of the others. Note: you can always add one at the top and replace the one you want to remove."
    End If
    cmbOut.ListIndex = 0

   On Error GoTo 0
   Exit Sub

btnOutRemove_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnOutRemove_Click of Form FireCallPrefs"
End Sub
' select a font for the chatbox areas alone on FireCallPrefs

' add new user defined text to the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnPrgAdd_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPrgAdd_Click()
   On Error GoTo btnPrgAdd_Click_Error

    If txtStringToAddFieldModified = True And txtStringToAdd.Text <> vbNullString Then
        cmbPrg.AddItem txtStringToAdd.Text, 0
    Else
        MsgBox "No text to add - please add your text in the box above and then retry."
    End If
    cmbPrg.ListIndex = 0

   On Error GoTo 0
   Exit Sub

btnPrgAdd_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPrgAdd_Click of Form FireCallPrefs"

End Sub
' remove user defined text from the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnPrgRemove_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPrgRemove_Click()
   On Error GoTo btnPrgRemove_Click_Error

    btnSave.Enabled = True ' enable the save button ' enable the save button
    If cmbPrg.ListIndex <> 0 Then
        cmbPrg.RemoveItem (cmbPrg.ListIndex)
    Else
        MsgBox "You cannot delete the first item in the list, try one of the others. Note: you can always add one at the top and replace the one you want to remove."
    End If
    cmbPrg.ListIndex = 0

   On Error GoTo 0
   Exit Sub

btnPrgRemove_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPrgRemove_Click of Form FireCallPrefs"
End Sub
' save the values from all the tabs
'---------------------------------------------------------------------------------------
' Procedure : btnSave_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnSave_Click()

    Dim foundMessage As Boolean
    Dim btnCnt As Integer
    Dim msgCnt As Integer
    Dim useloop As Integer
    Dim thisText As String
    

    'Dim smtpConfigValue As String
        
    ' save the values from the general tab
   On Error GoTo btnSave_Click_Error

    FCWSharedInputFile = txtSharedInputFile.Text
    FCWSharedOutputFile = txtSharedOutputFile.Text
    FCWExchangeFolder = txtExchangeFolder.Text
    
    FCWRefreshIntervalIndex = LTrim$(Str$(cmbRefreshInterval.ListIndex)) ' the index for the refresh
    FCWRefreshIntervalSecs = cmbRefreshInterval.ItemData(Val(FCWRefreshIntervalIndex)) ' the data

    FCWAlarmSoundIndex = LTrim$(Str$(cmbAlarmSound.ListIndex))
    FCWAlarmSound = cmbAlarmSound.List(cmbAlarmSound.ListIndex)

 
    
    ' save the values from the configuration tab
    FCWPrefixString = txtPrefixString.Text
    FCWLoadBottom = LTrim$(Str$(chkLoadBottom.Value))
    
    FCWMaxLineLengthIndex = LTrim$(Str$(cmbMaxLineLength.ListIndex))
    FCWMaxLineLength = cmbMaxLineLength.List(cmbMaxLineLength.ListIndex)

    FCWEnableScrollbars = LTrim$(Str$(chkEnableScrollbars.Value))
    FCWEnableTooltips = LTrim$(Str$(chkEnableTooltips.Value))
    FCWEnableBalloonTooltips = LTrim$(Str$(chkEnableBalloonTooltips.Value))
    
    FCWIconiseDelay = LTrim$(Str$(sliIconiseDelay.Value))
    
    ' save the values from the Emails tab
    FCWSendEmails = LTrim$(Str$(chkSendEmails.Value))
    FCWSendErrorEmails = LTrim$(Str$(chkSendErrorEmails.Value))
    
    'FCWEmailAddress = txtEmailAddress.Text
    FCWAdviceInterval = LTrim$(Str$(cmbAdviceInterval.ListIndex))
    FCWAdviceIntervalSecs = cmbAdviceInterval.ItemData(Val(FCWAdviceInterval)) ' the data
    
    
    'save the values from the Emojis Config Items
    FCWEmojiSetIndex = LTrim$(Str$(cmbEmojiSet.ListIndex))
    FCWEmojiSetDesc = cmbEmojiSet.List(cmbEmojiSet.ListIndex)
    
    'save the values from the Fonts Config Items
    FCWMainFont = txtTextFont.Text
    FCWMainFontSize = txtFontSize.Text
    'FCWMainFontItalics = txtFontSize.Text
    'FCWMainFontColour = txtFontSize.Text
    
    
    FCWPrefsFont = txtPrefsFont.Text
    FCWPrefsFontSize = txtPrefsFontSize.Text
    'FCWPrefsFontItalics = txtFontSize.Text
    
    
    'save the values from the Windows Config Items
    FCWWindowLevel = LTrim$(Str$(cmbWindowLevel.ListIndex))
    FCWOpacity = LTrim$(Str$(sliOpacity.Value))
    
    FCWEnableSounds = LTrim$(Str$(chkEnableSounds.Value))
    FCWEnableAlarmSound = LTrim$(Str$(chkEnableAlarmSound.Value))
    
    FCWPlayVolume = LTrim$(Str$(chkPlayVolume.Value))
    
    FCWSmtpConfig = LTrim$(Str$(cmbSmtpConfig.ListIndex))
    FCWSmtpConfigName = txtSmtpConfigName.Text
    
    FCWSmtpServer = txtSmtpServer.Text
    FCWSmtpUsername = txtSmtpUsername.Text
    FCWSmtpPassword = AesEncryptString(txtSmtpPassword.Text, emailTString)
    FCWSmtpPort = txtSmtpPort.Text
    FCWSmtpAuthenticate = LTrim$(Str$(cmbSmtpAuthenticate.ListIndex))
    FCWSmtpSecurity = LTrim$(Str$(cmbSmtpSecurity.ListIndex))
    
    FCWRecipientEmail = txtRecipientEmail.Text
    FCWEmailSubject = txtEmailSubject.Text
    FCWEmailMessage = txtEmailMessage.Text
    
    FCWSingleListBox = LTrim$(Str$(chkSingleListBox.Value))
    FCWAllowShutdowns = LTrim$(Str$(chkAllowShutdowns.Value))
    
    If optHandleData(0).Value = True Then FCWOptHandleData = "0"
    If optHandleData(1).Value = True Then FCWOptHandleData = "1"
'
'    If optWindowWidth(0).Value = True Then FCWOptWindowWidth = "10155"
'    If optWindowWidth(1).Value = True Then FCWOptWindowWidth = "12155"
'    If optWindowWidth(2).Value = True Then FCWOptWindowWidth = "14155"

    FCWAutomaticHousekeeping = LTrim$(Str$(chkAutomaticHousekeeping.Value))
    FCWStartup = LTrim$(Str$(chkGenStartup.Value))
    
    If FCWStartup = "1" Then
        Call savestring(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "FireCallWin", """" & App.Path & "\" & "FireCallWin.exe""")
    Else
        Call savestring(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "FireCallWin", "")
    End If
    
    FCWArchiveDays = cmbArchiveDays.List(cmbArchiveDays.ListIndex)
    FCWArchiveDaysIndex = cmbArchiveDays.ListIndex
    
    
    FCWBackupOnStart = LTrim$(Str$(chkBackupOnStart.Value))
    FCWAutomaticBackups = LTrim$(Str$(chkAutomaticBackups.Value))
    FCWAutomaticBackupInterval = LTrim$(Str$(sliAutomaticBackupInterval.Value))
    
    If optServiceProvider(0).Value = True Then FCWServiceProvider = "0"
    If optServiceProvider(1).Value = True Then FCWServiceProvider = "1"
    If optServiceProvider(2).Value = True Then FCWServiceProvider = "2"
    If optServiceProvider(3).Value = True Then FCWServiceProvider = "3"
    
    If chkServiceProcesses.Value = 1 Then FCWCheckServiceProcesses = "1"
    
    If recordingIsPossible = True Then
        FCWCaptureDevicesIndex = LTrim$(Str$(cmbCaptureDevices.ListIndex))
        FCWCaptureDevices = cmbCaptureDevices.List(cmbCaptureDevices.ListIndex)
        
        FireCallMain.cmbHiddenCaptureDevices.Clear
        If cmbCaptureDevices.ListCount > 0 Then
            For useloop = 0 To cmbCaptureDevices.ListCount - 1
                FireCallMain.cmbHiddenCaptureDevices.List(useloop) = cmbCaptureDevices.List(useloop)
            Next useloop
            
            FireCallMain.cmbHiddenCaptureDevices.ListIndex = cmbCaptureDevices.ListIndex
            FireCallMain.cmbHiddenCaptureDevices.Text = FireCallMain.cmbHiddenCaptureDevices.List(Val(cmbCaptureDevices.ListIndex))
        End If
    
    '    If optRecordingType(0).value = True Then FCWCaptureMethod = "0"
    '    If optRecordingType(1).value = True Then FCWCaptureMethod = "1"
    
        FCWRecordingQuality = LTrim$(Str$(sliRecordingQuality.Value))
    End If
    
    FCWIconiseOpacity = LTrim$(Str$(optIconiseOpacity.Value))
    FCWIconiseDesktop = LTrim$(Str$(optIconiseDesktop.Value))
    
    ' save the values from the general tab
    If fFExists(FCWSettingsFile) Then
        PutINISetting "Software\FireCallWin", "sharedInputFile", FCWSharedInputFile, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "sharedOutputFile", FCWSharedOutputFile, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "exchangeFolder", FCWExchangeFolder, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "refreshIntervalIndex", FCWRefreshIntervalIndex, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "refreshIntervalSecs", FCWRefreshIntervalSecs, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "alarmSoundIndex", FCWAlarmSoundIndex, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "alarmSound", FCWAlarmSound, FCWSettingsFile
        
        
        ' save the values from the configuration tab
        PutINISetting "Software\FireCallWin", "prefixString", FCWPrefixString, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "loadBottom", FCWLoadBottom, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "maxLineLengthIndex", FCWMaxLineLengthIndex, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "maxLineLength", FCWMaxLineLength, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "enableScrollbars", FCWEnableScrollbars, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "enableTooltips", FCWEnableTooltips, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "enableBalloonTooltips", FCWEnableBalloonTooltips, FCWSettingsFile
       
        PutINISetting "Software\FireCallWin", "iconiseDelay", FCWIconiseDelay, FCWSettingsFile
    
        ' save the values from the Emails tab
        PutINISetting "Software\FireCallWin", "sendEmails", FCWSendEmails, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "sendErrorEmails", FCWSendErrorEmails, FCWSettingsFile
        
        
        'PutINISetting "Software\FireCallWin", "emailAddress", FCWEmailAddress, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "adviceInterval", FCWAdviceInterval, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "adviceIntervalSecs", FCWAdviceIntervalSecs, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "lastEmail", FCWLastEmail, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "lastHouseKeep", FCWLastHouseKeep, FCWSettingsFile
        
        
        
        'save the values from the Emojis Config Items
        PutINISetting "Software\FireCallWin", "emojiSetIndex", FCWEmojiSetIndex, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "emojiSetDesc", FCWEmojiSetDesc, FCWSettingsFile
            
        'save the values from the Fonts Config Items
        PutINISetting "Software\FireCallWin", "mainFont", FCWMainFont, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "mainFontSize", FCWMainFontSize, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "mainFontItalics", FCWMainFontItalics, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "mainFontColour", FCWMainFontColour, FCWSettingsFile
        
        
        PutINISetting "Software\FireCallWin", "prefsFont", FCWPrefsFont, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "prefsFontSize", FCWPrefsFontSize, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "prefsFontItalics", FCWPrefsFontItalics, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "prefsFontColour", FCWPrefsFontColour, FCWSettingsFile
        
         
        'save the values from the Windows Config Items
        PutINISetting "Software\FireCallWin", "windowLevel", FCWWindowLevel, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "opacity", FCWOpacity, FCWSettingsFile

        PutINISetting "Software\FireCallWin", "minimiseFormX", FCWMinimiseFormX, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "minimiseFormY", FCWMinimiseFormY, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "lastSoundPlayed", FCWLastSoundPlayed, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "lastAwakeString", FCWLastAwakeString, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "lastShutdown", FCWLastShutdown, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "allowShutdowns", FCWAllowShutdowns, FCWSettingsFile
        
        PutINISetting "Software\FireCallWin", "optHandleData", FCWOptHandleData, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "optWindowWidth", FCWOptWindowWidth, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "automaticHousekeeping", FCWAutomaticHousekeeping, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "startup", FCWStartup, FCWSettingsFile

        PutINISetting "Software\FireCallWin", "archiveDays", FCWArchiveDays, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "archiveDaysIndex", FCWArchiveDaysIndex, FCWSettingsFile
        

        PutINISetting "Software\FireCallWin", "backupOnStart", FCWBackupOnStart, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "automaticBackups", FCWAutomaticBackups, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "automaticBackupInterval", FCWAutomaticBackupInterval, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "serviceProvider", FCWServiceProvider, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "checkServiceProcesses", FCWCheckServiceProcesses, FCWSettingsFile
        
        PutINISetting "Software\FireCallWin", "msgBox13Enabled", FCWMsgBox13Enabled, FCWSettingsFile
        
        PutINISetting "Software\FireCallWin", "captureDevices", FCWCaptureDevices, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "captureDevicesIndex", FCWCaptureDevicesIndex, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "recordingQuality", FCWRecordingQuality, FCWSettingsFile


        PutINISetting "Software\FireCallWin", "enableSounds", FCWEnableSounds, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "enableAlarmSound", FCWEnableAlarmSound, FCWSettingsFile
        
        PutINISetting "Software\FireCallWin", "playVolume", FCWPlayVolume, FCWSettingsFile
        
        ' find the currently selected SMTP config option
        
        PutINISetting "Software\FireCallWin", "smtpConfig", Trim$(Str$(cmbSmtpConfig.ListIndex)), FCWSettingsFile
        PutINISetting "Software\FireCallWin", "smtpConfigName" & cmbSmtpConfig.ListIndex, txtSmtpConfigName.Text, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "smtpServer" & cmbSmtpConfig.ListIndex, FCWSmtpServer, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "SMTPUsername" & cmbSmtpConfig.ListIndex, FCWSmtpUsername, FCWSettingsFile
             
        ' we no longer use WritePrivateProfileString in PutINISetting as it cannot write certain special chars
        ' generated by the encryption routine.
        
        Dim b64FCWSMTPPassword As String
        Bas64.sBuffer = FCWSmtpPassword
        Call Bas64.Base64Encode
        b64FCWSMTPPassword = Bas64.Base64Buf
        
        PutINISetting "Software\FireCallWin", "SMTPPassword" & cmbSmtpConfig.ListIndex, b64FCWSMTPPassword, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "smtpPort" & cmbSmtpConfig.ListIndex, FCWSmtpPort, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "smtpAuthenticate" & cmbSmtpConfig.ListIndex, FCWSmtpAuthenticate, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "smtpSecurity" & cmbSmtpConfig.ListIndex, FCWSmtpSecurity, FCWSettingsFile
        
        'Call altPutPrivateProfileString("Software\FireCallWin", "SMTPPassword", b64FCWSMTPPassword, FCWSettingsFile)


        PutINISetting "Software\FireCallWin", "recipientEmail" & cmbSmtpConfig.ListIndex, FCWRecipientEmail, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "emailSubject" & cmbSmtpConfig.ListIndex, FCWEmailSubject, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "emailMessage" & cmbSmtpConfig.ListIndex, FCWEmailMessage, FCWSettingsFile
        
        PutINISetting "Software\FireCallWin", "singleListBox", FCWSingleListBox, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "imageDisplay", FCWImageDisplay, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "lastSelectedTab", FCWLastSelectedTab, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "iconiseOpacity", FCWIconiseOpacity, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "iconiseDesktop", FCWIconiseDesktop, FCWSettingsFile
        
        PutINISetting "Software\FireCallWin", "archiveFolder", FCWArchiveFolder, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "backupFolder", FCWBackupFolder, FCWSettingsFile
        
        
        'save the values from the Text Items
          
        foundMessage = False
        btnCnt = 0
        msgCnt = 0
        
        For useloop = 1 To 11
            foundMessage = True
            msgCnt = 0
            Do Until foundMessage = False
                btnCnt = useloop
    
                foundMessage = False
                If btnCnt = 1 Then thisText = cmbTTFN.List(msgCnt)
                If btnCnt = 2 Then thisText = cmbWell.List(msgCnt)
                If btnCnt = 3 Then thisText = cmbNews.List(msgCnt)
                If btnCnt = 4 Then thisText = cmbMorn.List(msgCnt)
                If btnCnt = 5 Then thisText = cmbWot.List(msgCnt)
                If btnCnt = 6 Then thisText = cmbWth.List(msgCnt)
                If btnCnt = 7 Then thisText = cmbPrg.List(msgCnt)
                If btnCnt = 8 Then thisText = cmbGdn.List(msgCnt)
                If btnCnt = 9 Then thisText = cmbBusy.List(msgCnt)
                If btnCnt = 10 Then thisText = cmbCod.List(msgCnt)
                If btnCnt = 11 Then thisText = cmbOut.List(msgCnt)
                If thisText <> vbNullString Then foundMessage = True
                msgCnt = msgCnt + 1
                
                PutINISetting "Software\FireCallWin", "button" & btnCnt & "message" & msgCnt, thisText, FCWSettingsFile
            Loop
        Next useloop
    End If
    
    FireCallMain.lbxInputTextArea.Clear
    FireCallMain.lbxOutputTextArea.Clear
    btnSave.Enabled = False ' disable the save button showing it has successfully saved
    
    Call testEmailTestButton
    
    Call FireCallMain.formLoadTasks ' the only place where a routine is called in another form
    FireCallPrefs.SetFocus
    

   On Error GoTo 0
   Exit Sub

btnSave_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnSave_Click of Form FireCallPrefs"

End Sub



' add a file to act as the shared input file
'---------------------------------------------------------------------------------------
' Procedure : btnSharedInputFile_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnSharedInputFile_Click()
    Dim retFileName As String
    'Dim retfileTitle As String
    Dim answer As VbMsgBoxResult
    
    'retfileTitle = ""
   On Error GoTo btnSharedInputFile_Click_Error

    retFileName = vbNullString

    Call addTargetFile(txtSharedInputFile.Text, retFileName)
    
    If retFileName <> vbNullString Then
        txtSharedInputFile.Text = retFileName ' just assigning it to a text field strips the buffered bit, leaving just the filename
    End If
    
    If retFileName = vbNullString Then
        'txtSharedInputFile.Text = ""
        Exit Sub
    End If
    
    If Not fFExists(retFileName) Then
        answer = MsgBox("The file doesn't currently exist, do you want me to create the chosen file, " & "   -  are you sure?", vbYesNo)
        If answer = vbNo Then
            Exit Sub
        End If
    
        'create new
        Open retFileName For Output As #1
        Close #1
    End If
    
    

   On Error GoTo 0
   Exit Sub

btnSharedInputFile_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnSharedInputFile_Click of Form FireCallPrefs"

End Sub
' remove user defined text from the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnWellRemove_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnWellRemove_Click()
   On Error GoTo btnWellRemove_Click_Error

    btnSave.Enabled = True ' enable the save button
    If cmbWell.ListIndex <> 0 Then
        cmbWell.RemoveItem (cmbWell.ListIndex)
    Else
        MsgBox "You cannot delete the first item in the list, try one of the others. Note: you can always add one at the top and replace the one you want to remove."
    End If
    cmbWell.ListIndex = 0

   On Error GoTo 0
   Exit Sub

btnWellRemove_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnWellRemove_Click of Form FireCallPrefs"
End Sub
' remove user defined text from the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnWotRemove_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnWotRemove_Click()
   On Error GoTo btnWotRemove_Click_Error

    btnSave.Enabled = True ' enable the save button
    If cmbWot.ListIndex <> 0 Then
        cmbWot.RemoveItem (cmbWot.ListIndex)
    Else
        MsgBox "You cannot delete the first item in the list, try one of the others. Note: you can always add one at the top and replace the one you want to remove."
    End If
    cmbWot.ListIndex = 0

   On Error GoTo 0
   Exit Sub

btnWotRemove_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnWotRemove_Click of Form FireCallPrefs"
End Sub
' remove user defined text from the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnWthRemove_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnWthRemove_Click()
   On Error GoTo btnWthRemove_Click_Error

    btnSave.Enabled = True ' enable the save button
    If cmbWth.ListIndex <> 0 Then
        cmbWth.RemoveItem (cmbWth.ListIndex)
    Else
        MsgBox "You cannot delete the first item in the list, try one of the others. Note: you can always add one at the top and replace the one you want to remove."
    End If
    cmbWth.ListIndex = 0

   On Error GoTo 0
   Exit Sub

btnWthRemove_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnWthRemove_Click of Form FireCallPrefs"
End Sub
' scrollbar enable/disable
'---------------------------------------------------------------------------------------
' Procedure : chkEnableScrollbars_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkEnableScrollbars_Click()
   On Error GoTo chkEnableScrollbars_Click_Error

    btnSave.Enabled = True ' enable the save button
    
    FCWEnableScrollbars = LTrim$(Str$(chkEnableScrollbars.Value))

   On Error GoTo 0
   Exit Sub

chkEnableScrollbars_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure chkEnableScrollbars_Click of Form FireCallPrefs"
        
End Sub

' set a var on a checkbox tick
'---------------------------------------------------------------------------------------
' Procedure : chkEnableTooltips_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkEnableTooltips_Click()
   On Error GoTo chkEnableTooltips_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkEnableTooltips_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure chkEnableTooltips_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkEnableSounds_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkEnableSounds_Click()
   On Error GoTo chkEnableSounds_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkEnableSounds_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure chkEnableSounds_Click of Form FireCallPrefs"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : chkIgnoreMouse_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkIgnoreMouse_Click()
   On Error GoTo chkIgnoreMouse_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkIgnoreMouse_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure chkIgnoreMouse_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkLoadBottom_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkLoadBottom_Click()
   On Error GoTo chkLoadBottom_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkLoadBottom_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure chkLoadBottom_Click of Form FireCallPrefs"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : chkPlayVolume_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkPlayVolume_Click()
   On Error GoTo chkPlayVolume_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkPlayVolume_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure chkPlayVolume_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkSendEmails_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkSendEmails_Click()
   On Error GoTo chkSendEmails_Click_Error

    btnSave.Enabled = True ' enable the save button
    If chkSendEmails.Value = 1 Or chkSendErrorEmails.Value = 1 Then
        Call toggleAllEmailControls("show")
    Else
        Call toggleAllEmailControls("hide")
    End If

   On Error GoTo 0
   Exit Sub

chkSendEmails_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure chkSendEmails_Click of Form FireCallPrefs"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : chkSendErrorEmails_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkSendErrorEmails_Click()
   On Error GoTo chkSendErrorEmails_Click_Error

    btnSave.Enabled = True ' enable the save button
    If chkSendEmails.Value = 1 Or chkSendErrorEmails.Value = 1 Then
        Call toggleAllEmailControls("show")
    Else
        Call toggleAllEmailControls("hide")
    End If

   On Error GoTo 0
   Exit Sub

chkSendErrorEmails_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure chkSendErrorEmails_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkSingleListBox_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkSingleListBox_Click()
   On Error GoTo chkSingleListBox_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkSingleListBox_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure chkSingleListBox_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbAdviceInterval_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbAdviceInterval_Click()
   On Error GoTo cmbAdviceInterval_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

cmbAdviceInterval_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmbAdviceInterval_Click of Form FireCallPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbAlarmSound_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbAlarmSound_Click()
   On Error GoTo cmbAlarmSound_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

cmbAlarmSound_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmbAlarmSound_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbEmojiSet_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbEmojiSet_Click()
   On Error GoTo cmbEmojiSet_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

cmbEmojiSet_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmbEmojiSet_Click of Form FireCallPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbMaxLineLength_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbMaxLineLength_Click()
   On Error GoTo cmbMaxLineLength_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

cmbMaxLineLength_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmbMaxLineLength_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbRefreshInterval_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbRefreshInterval_Click()
   On Error GoTo cmbRefreshInterval_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

cmbRefreshInterval_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmbRefreshInterval_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbWindowLevel_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbWindowLevel_Click()
   On Error GoTo cmbWindowLevel_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

cmbWindowLevel_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmbWindowLevel_Click of Form FireCallPrefs"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : btnPrefsFont_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPrefsFont_Click()
   On Error GoTo btnPrefsFont_Click_Error

    btnSave.Enabled = True ' enable the save button

    Dim fntFont As String
    Dim fntSize As Integer
    Dim fntWeight As Integer
    Dim fntStyle As Boolean
    Dim fntColour As Long
    Dim fntItalics As Boolean
    Dim fntUnderline As Boolean
    Dim fntFontResult As Boolean

    fntFont = FCWPrefsFont
    fntSize = Val(FCWPrefsFontSize)
    fntItalics = CBool(FCWPrefsFontItalics)
    fntColour = CLng(FCWPrefsFontColour)
        
    Call changeFont(FireCallPrefs, True, fntFont, fntSize, fntWeight, fntStyle, fntColour, fntItalics, fntUnderline, fntFontResult)
                                'ByRef fntFont As String, ByRef fntSize As Integer, ByRef fntWeight As Integer, ByRef fntStyle As Boolean, ByRef fntColour As Long, ByRef fntItalics As Boolean, ByRef fntUnderline As Boolean, ByRef fntFontResult As Boolean)
    
    FCWPrefsFont = CStr(fntFont)
    FCWPrefsFontSize = CStr(fntSize)
    FCWPrefsFontItalics = CStr(fntItalics)
    FCWPrefsFontColour = CStr(fntColour)
    
    If fFExists(FCWSettingsFile) Then ' does the tool's own settings.ini exist?
        PutINISetting "Software\FireCallWin", "prefsTextFont", FCWPrefsFont, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "prefsFontSize", FCWPrefsFontSize, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "prefsFontItalics", FCWPrefsFontItalics, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "PrefsFontColour", FCWPrefsFontColour, FCWSettingsFile
    End If
    
    If fntFont = vbNullString Then fntFont = "arial"
    txtPrefsFont.Text = fntFont
    txtPrefsFont.Font.Name = fntFont
    txtPrefsFont.Font.Size = fntSize
    txtPrefsFont.Font.Italic = fntItalics
    txtPrefsFont.ForeColor = fntColour
    
    txtPrefsFontSize.Text = fntSize

   On Error GoTo 0
   Exit Sub

btnPrefsFont_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPrefsFont_Click of Form FireCallPrefs"

End Sub
' select a font for the chatbox areas alone on FireCallMain
'---------------------------------------------------------------------------------------
' Procedure : btnTextFont_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnTextFont_Click()
    Dim storedFont As String
    
    Dim fntFont As String
    Dim fntSize As Integer
    Dim fntWeight As Integer
    Dim fntStyle As Boolean
    Dim fntColour As Long
    Dim fntItalics As Boolean
    Dim fntUnderline As Boolean
    Dim fntFontResult As Boolean
    
   On Error GoTo btnTextFont_Click_Error

    btnSave.Enabled = True ' enable the save button
    
    storedFont = txtTextFont.Text
    
    fntFont = FCWMainFont
    fntSize = FCWMainFontSize
    fntItalics = CBool(FCWMainFontItalics)
    fntColour = CLng(FCWMainFontColour)
    
    Call changeFont(FireCallMain, False, fntFont, fntSize, fntWeight, fntStyle, fntColour, fntItalics, fntUnderline, fntFontResult)
    If fntFont = vbNullString Then
        fntFont = storedFont
        fntSize = "8"
    End If
    
    If fntSize = "0" Then
        fntSize = "8"
    End If
    
    FCWMainFont = CStr(fntFont)
    FCWMainFontSize = CStr(fntSize)
    FCWMainFontItalics = CStr(fntItalics)
    FCWMainFontColour = CStr(fntColour)
    
    If fFExists(FCWSettingsFile) Then ' does the tool's own settings.ini exist?
        PutINISetting "Software\FireCallWin", "mainFont", FCWMainFont, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "mainFontSize", FCWMainFontSize, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "mainFontItalics", FCWMainFontItalics, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "mainFontColour", FCWMainFontColour, FCWSettingsFile
        
    End If
    
    txtTextFont.Text = fntFont
    txtTextFont.Font.Name = fntFont
    txtTextFont.Font.Size = Val(fntSize)
    txtTextFont.Font.Italic = fntItalics
    txtTextFont.ForeColor = fntColour
    
    'txtFontSize.Text = fntSize
    FireCallMain.lbxInputTextArea.Height = 4300
    FireCallMain.lbxOutputTextArea.Height = 4300
    
    'FireCallMain.cmbEmojiSelection.SelLength = 0


   On Error GoTo 0
   Exit Sub

btnTextFont_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnTextFont_Click of Form FireCallPrefs"
    
End Sub

' add a file to act as the shared input file
'---------------------------------------------------------------------------------------
' Procedure : btnSharedOutputFile_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnSharedOutputFile_Click()
    Dim retFileName As String
    'Dim retfileTitle As String
    Dim answer As VbMsgBoxResult

   On Error GoTo btnSharedOutputFile_Click_Error

    Call addTargetFile(txtSharedOutputFile.Text, retFileName)
    
    If retFileName <> vbNullString Then
        txtSharedOutputFile.Text = retFileName ' strips the buffered bit, leaving just the filename
    End If
    
    If retFileName = vbNullString Then
        'txtSharedOutputFile.Text = ""
        Exit Sub
    End If
    
    If Not fFExists(retFileName) Then
        answer = MsgBox("The file doesn't currently exist, do you want me to create the chosen file, " & "   -  are you sure?", vbYesNo)
        If answer = vbNo Then
            Exit Sub
        End If
    
        'create new
        Open retFileName For Output As #1
        Close #1
    End If
    

   On Error GoTo 0
   Exit Sub

btnSharedOutputFile_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnSharedOutputFile_Click of Form FireCallPrefs"

End Sub
' add a file to act as the shared exchange folder
'---------------------------------------------------------------------------------------
' Procedure : btnExchangeFolder_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnExchangeFolder_Click()
    ' variables declared
    Dim getFolder As String
    Dim dialogInitDir As String
   
   On Error GoTo btnExchangeFolder_Click_Error

    If debugflg = 1 Then Debug.Print "%btnGeneralRdFolder_Click"
    
   'initialise the dimensioned variables
    getFolder = vbNullString
    dialogInitDir = vbNullString
    
    If txtExchangeFolder.Text <> vbNullString Then
        If fDirExists(txtExchangeFolder.Text) Then
            dialogInitDir = txtExchangeFolder.Text 'start dir, might be "C:\" or so also
        Else
            dialogInitDir = App.Path 'start dir, might be "C:\" or so also
        End If
    End If

    getFolder = fBrowseFolder(Me.hwnd, dialogInitDir) ' show the dialog box to select a folder
    If getFolder <> vbNullString Then txtExchangeFolder.Text = getFolder

   On Error GoTo 0
   Exit Sub

btnExchangeFolder_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnExchangeFolder_Click of Form FireCallPrefs"

End Sub


' open an explorer window and show the default emoji folder
'---------------------------------------------------------------------------------------
' Procedure : btnEmojiLocation_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnEmojiLocation_Click()
   On Error GoTo btnEmojiLocation_Click_Error

        If fDirExists(App.Path & "\Resources\emojis\") Then
            Call ShellExecute(Me.hwnd, "Open", App.Path & "\Resources\emojis\", vbNullString, App.Path, 1)
        End If

   On Error GoTo 0
   Exit Sub

btnEmojiLocation_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnEmojiLocation_Click of Form FireCallPrefs"
End Sub
' add new user defined text to the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnMornAdd_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnMornAdd_Click()
   On Error GoTo btnMornAdd_Click_Error

    If txtStringToAddFieldModified = True And txtStringToAdd.Text <> vbNullString Then
        If cmbMorn.ListCount <= 9 Then
            cmbMorn.AddItem txtStringToAdd.Text, 0
        Else
            MsgBox "A maximum of 10 messages per button - please remove some of the other assigned texts for this button, then retry."
        End If
    Else
        MsgBox "No text to add - please add your text in the box above and then retry."
    End If
    cmbMorn.ListIndex = 0

   On Error GoTo 0
   Exit Sub

btnMornAdd_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnMornAdd_Click of Form FireCallPrefs"

End Sub
' add new user defined text to the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnWotAdd_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnWotAdd_Click()
   On Error GoTo btnWotAdd_Click_Error

    If txtStringToAddFieldModified = True And txtStringToAdd.Text <> vbNullString Then
        If cmbWot.ListCount <= 9 Then
            cmbWot.AddItem txtStringToAdd.Text, 0
        Else
            MsgBox "A maximum of 10 messages per button - please remove some of the other assigned texts for this button, then retry."
        End If
    Else
        MsgBox "No text to add - please add your text in the box above and then retry."
    End If
    cmbWot.ListIndex = 0

   On Error GoTo 0
   Exit Sub

btnWotAdd_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnWotAdd_Click of Form FireCallPrefs"

End Sub
' add new user defined text to the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnWthAdd_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnWthAdd_Click()
   On Error GoTo btnWthAdd_Click_Error

    If txtStringToAddFieldModified = True And txtStringToAdd.Text <> vbNullString Then
        If cmbWth.ListCount <= 9 Then
            cmbWth.AddItem txtStringToAdd.Text, 0
        Else
            MsgBox "A maximum of 10 messages per button - please remove some of the other assigned texts for this button, then retry."
        End If
    Else
        MsgBox "No text to add - please add your text in the box above and then retry."
    End If
    cmbWth.ListIndex = 0

   On Error GoTo 0
   Exit Sub

btnWthAdd_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnWthAdd_Click of Form FireCallPrefs"

End Sub
' open an explorer window and show the default sounds folder
'---------------------------------------------------------------------------------------
' Procedure : btnSoundsLocation_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnSoundsLocation_Click()
   On Error GoTo btnSoundsLocation_Click_Error

        If fDirExists(App.Path & "\Resources\sounds\") Then
            Call ShellExecute(Me.hwnd, "Open", App.Path & "\Resources\sounds\", vbNullString, App.Path, 1)
        End If

   On Error GoTo 0
   Exit Sub

btnSoundsLocation_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnSoundsLocation_Click of Form FireCallPrefs"
End Sub
' add new user defined text to the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnTtfnAdd_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnTtfnAdd_Click()
   On Error GoTo btnTtfnAdd_Click_Error

    If txtStringToAddFieldModified = True And txtStringToAdd.Text <> vbNullString Then
        If cmbTTFN.ListCount <= 9 Then
            cmbTTFN.AddItem txtStringToAdd.Text, 0
        Else
            MsgBox "A maximum of 10 messages per button - please remove some of the other assigned texts for this button, then retry."
        End If
    Else
        MsgBox "No text to add - please add your text in the box above and then retry."
    End If
    cmbTTFN.ListIndex = 0

   On Error GoTo 0
   Exit Sub

btnTtfnAdd_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnTtfnAdd_Click of Form FireCallPrefs"

End Sub
' mute any test sound playing
'---------------------------------------------------------------------------------------
' Procedure : btnMute_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnMute_Click()
    Dim fileToPlay As String
    
   On Error GoTo btnMute_Click_Error

    fileToPlay = "nothing.wav"
    If fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    'btnSave.SetFocus

   On Error GoTo 0
   Exit Sub

btnMute_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnMute_Click of Form FireCallPrefs"
End Sub
' play a sound chosen from the sounds folder
'---------------------------------------------------------------------------------------
' Procedure : btnPlaySound_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPlaySound_Click()
    'Dim answer As VbMsgBoxResult
    Dim fileToPlay As String
        
'    If debugflg = 1 Then Debug.Print "%" & "mnuDelete_Click"
'
'    fileToKill = cmbAlarmSound.List(cmbAlarmSound.ListIndex)
'
'    ' delete the sound
'    If fileToKill = "G6AUC.wav" Then
'        MsgBox ("This is the default alarm and cannot be deleted")
'    Else
'        answer = MsgBox("This will delete the currently selected sound, " & fileToKill & " -  are you sure?", vbYesNo)
'        If answer = vbNo Then
'            Exit Sub
'        End If
'
'        Kill App.Path & "\resources\sounds\" & fileToKill
'        cmbAlarmSound.Clear
'        Call populateCmbAlarmSound
'        cmbAlarmSound.Text = cmbAlarmSound.List(0)
'
'        MsgBox (fileToKill & " file deleted")
'    End If
    
   On Error GoTo btnPlaySound_Click_Error

    fileToPlay = cmbAlarmSound.List(cmbAlarmSound.ListIndex)

    ' delete the sound
    'If fileToPlay = "G6AUC.wav" Then

        PlaySound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
'        cmbAlarmSound.Clear
'        Call populateCmbAlarmSound
'        cmbAlarmSound.Text = cmbAlarmSound.List(0)
'
'        MsgBox (fileToKill & " file deleted")
    'End If
    'btnSave.SetFocus

   On Error GoTo 0
   Exit Sub

btnPlaySound_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPlaySound_Click of Form FireCallPrefs"
    
End Sub
' add new user defined text to the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnWellAdd_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnWellAdd_Click()
   On Error GoTo btnWellAdd_Click_Error

    If txtStringToAddFieldModified = True And txtStringToAdd.Text <> vbNullString Then
        If cmbWell.ListCount <= 9 Then
            cmbWell.AddItem txtStringToAdd.Text, 0
        Else
            MsgBox "A maximum of 10 messages per button - please remove some of the other assigned texts for this button, then retry."
        End If
    Else
        MsgBox "No text to add - please add your text in the box above and then retry."
    End If
    cmbWell.ListIndex = 0


   On Error GoTo 0
   Exit Sub

btnWellAdd_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnWellAdd_Click of Form FireCallPrefs"
    
End Sub
' add new user defined text to the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnNewsAdd_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnNewsAdd_Click()
   On Error GoTo btnNewsAdd_Click_Error

    If txtStringToAddFieldModified = True And txtStringToAdd.Text <> vbNullString Then
        If cmbNews.ListCount <= 9 Then
            cmbNews.AddItem txtStringToAdd.Text, 0
        Else
            MsgBox "A maximum of 10 messages per button - please remove some of the other assigned texts for this button, then retry."
        End If
    Else
        MsgBox "No text to add - please add your text in the box above and then retry."
    End If
    cmbNews.ListIndex = 0

   On Error GoTo 0
   Exit Sub

btnNewsAdd_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnNewsAdd_Click of Form FireCallPrefs"

End Sub
' remove user defined text from the pre-defined button -
'---------------------------------------------------------------------------------------
' Procedure : btnTtfnRemove_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnTtfnRemove_Click()
   On Error GoTo btnTtfnRemove_Click_Error

    btnSave.Enabled = True ' enable the save button

    If cmbTTFN.ListIndex <> 0 Then
        cmbTTFN.RemoveItem (cmbTTFN.ListIndex)
    Else
        MsgBox "You cannot delete the first item in the list, try one of the others. Note: you can always add one at the top and replace the one you want to remove."
    End If
    cmbTTFN.ListIndex = 0

   On Error GoTo 0
   Exit Sub

btnTtfnRemove_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnTtfnRemove_Click of Form FireCallPrefs"

End Sub



 

'---------------------------------------------------------------------------------------
' Procedure : adjustPrefsControls
' Author    : beededea
' Date      : 12/05/2020
' Purpose   : adjust the controls so their startup position matches the last write of the config file
'---------------------------------------------------------------------------------------
'
Private Sub adjustPrefsControls()

    Dim fntFont As String
    Dim fntSize As Integer
    Dim fntWeight As Integer
    Dim fntStyle As Boolean
    Dim fntColour As Long
    Dim fntItalics As Boolean
    Dim fntUnderline As Boolean
    Dim fntFontResult As Boolean
    
    fntWeight = 0
    fntStyle = False

    ' save the values from the general tab
    txtSharedInputFile.Text = FCWSharedInputFile
    txtSharedOutputFile.Text = FCWSharedOutputFile
    txtExchangeFolder.Text = FCWExchangeFolder
    cmbRefreshInterval.ListIndex = Val(FCWRefreshIntervalIndex)
    
    ' the contents are already populated
    cmbAlarmSound.ListIndex = Val(FCWAlarmSoundIndex)
         
    ' save the values from the configuration tab
    txtPrefixString.Text = FCWPrefixString
    chkLoadBottom.Value = Val(FCWLoadBottom)
    
    cmbMaxLineLength.ListIndex = Val(FCWMaxLineLengthIndex)
    
    chkEnableScrollbars.Value = Val(FCWEnableScrollbars)
    
    chkEnableTooltips.Value = Val(FCWEnableTooltips)
    chkEnableBalloonTooltips.Value = Val(FCWEnableBalloonTooltips)
    
    
    'chkInputSelection.Value = Val(FCWInputSelection)
    sliIconiseDelay.Value = Val(FCWIconiseDelay)
    
    ' save the values from the Emails tab
    chkSendEmails.Value = Val(FCWSendEmails)
    chkSendErrorEmails.Value = Val(FCWSendErrorEmails)
        
    'txtEmailAddress.Text = FCWEmailAddress
    cmbAdviceInterval.ListIndex = Val(FCWAdviceInterval) 'nnn

    'save the values from the Emojis Config Items
    cmbEmojiSet.ListIndex = Val(FCWEmojiSetIndex)
    cmbEmojiSet.Text = cmbEmojiSet.List(Val(FCWEmojiSetIndex))
    
    'save the values from the Fonts Config Items
    txtTextFont.Text = FCWMainFont
    txtFontSize.Text = FCWMainFontSize
    txtTextFont.Font.Name = txtTextFont.Text
        
    txtPrefsFont.Text = FCWPrefsFont
    txtPrefsFontSize.Text = FCWPrefsFontSize
    
    If FCWPrefsFont <> vbNullString Then
        Call changeFormFont(FireCallPrefs, FCWPrefsFont, Val(FCWPrefsFontSize), fntWeight, fntStyle, FCWPrefsFontItalics, FCWPrefsFontColour)
    End If

    Call resetComboBoxHighlight
    
    'save the values from the Windows Items
    cmbWindowLevel.ListIndex = Val(FCWWindowLevel)
'    chkIgnoreMouse.Value = Val(FCWIgnoreMouse)
'    chkPreventDragging.Value = Val(FCWPreventDragging)
    sliOpacity.Value = Val(FCWOpacity)

    'forces the two listboxes to a specific height regardless of fonts chosen.
'    FireCallMain.lbxInputTextArea.Height = 4300
'    FireCallMain.lbxOutputTextArea.Height = 4300
    
'    If FCWClockStyle = "RedButton" Then
'        FireCallMain.picRedButton.Visible = True
'        FireCallMain.picClock.Visible = False
'    Else
'        FireCallMain.picRedButton.Visible = False
'        FireCallMain.picClock.Visible = True
'    End If
    
    chkEnableSounds.Value = Val(FCWEnableSounds)
    chkEnableAlarmSound.Value = Val(FCWEnableAlarmSound)
    
    chkPlayVolume.Value = Val(FCWPlayVolume)

    cmbSmtpConfig.ListIndex = Val(FCWSmtpConfig)
    ' we used to do     'Call adjustPrefsSmtpControls but a click sets and selects correctly
    btnSave.Enabled = False ' disable the save button after the click above

    
    txtRecipientEmail.Text = FCWRecipientEmail
    txtEmailSubject.Text = FCWEmailSubject
    txtEmailMessage.Text = FCWEmailMessage
    
    chkSingleListBox.Value = Val(FCWSingleListBox)
    chkAllowShutdowns.Value = Val(FCWAllowShutdowns)
    chkGenStartup.Value = Val(FCWStartup)
    
    cmbArchiveDays.ListIndex = Val(FCWArchiveDaysIndex)
    
    
    chkBackupOnStart.Value = Val(FCWBackupOnStart)
    chkAutomaticBackups.Value = Val(FCWAutomaticBackups)
    sliAutomaticBackupInterval.Value = Val(FCWAutomaticBackupInterval)
    
    chkAutomaticHousekeeping.Value = Val(FCWAutomaticHousekeeping)
    
'    If FCWSingleListBox = "1" Then
'        FireCallMain.lbxInputTextArea.Visible = False
'        FireCallMain.lbxOutputTextArea.Visible = False
'
'        FireCallMain.lbxCombinedTextArea.Height = 8300
'        FireCallMain.lbxCombinedTextArea.Visible = True
'    Else
'        FireCallMain.lbxInputTextArea.Visible = True
'        FireCallMain.lbxOutputTextArea.Visible = True
'
'        FireCallMain.lbxCombinedTextArea.Visible = False
'
'    End If
    
'    If FCWPlayVolume = "1" Then
'        FireCallMain.picSpeakerGrille.Visible = False
'        FireCallMain.picSpeakerGrilleOpen.Visible = True
'    Else
'        FireCallMain.picSpeakerGrille.Visible = True
'        FireCallMain.picSpeakerGrilleOpen.Visible = False
'    End If

    If FCWOptHandleData = "0" Then optHandleData(0).Value = True
    If FCWOptHandleData = "1" Then optHandleData(1).Value = True

    If FCWServiceProvider = "0" Then optServiceProvider(0).Value = True
    If FCWServiceProvider = "1" Then optServiceProvider(1).Value = True
    If FCWServiceProvider = "2" Then optServiceProvider(2).Value = True
    If FCWServiceProvider = "3" Then optServiceProvider(3).Value = True
    
    If FCWCheckServiceProcesses = "1" Then chkServiceProcesses.Value = 1
    
    If recordingIsPossible = True Then
        cmbCaptureDevices.ListIndex = Val(FCWCaptureDevicesIndex)
        cmbCaptureDevices.Text = cmbCaptureDevices.List(Val(FCWCaptureDevicesIndex))
    Else
        cmbCaptureDevices.Text = "No recording devices found" 'dean
    End If
    
'    If FCWCaptureMethod = "0" Then optRecordingType(0).value = True
'    If FCWCaptureMethod = "1" Then optRecordingType(1).value = True

    sliRecordingQuality.Value = Val(FCWRecordingQuality)
    optIconiseOpacity.Value = CBool(FCWIconiseOpacity)
    If optIconiseOpacity.Value = False Then optIconiseDesktop.Value = True
    
    Call checkIconiseOpacityLevel
   
   On Error GoTo 0
   Exit Sub

adjustPrefsControls_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure adjustPrefsControls of Form dockSettings on line " & Erl

End Sub
' all combo boxes in the prefs are populated here
'---------------------------------------------------------------------------------------
' Procedure : populateComboBoxes
' Author    : beededea
' Date      : 10/09/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub populateComboBoxes()
    Dim MyPath  As String
    Dim myName As String

    Dim buttonmessage As String
    Dim foundMessage As Boolean
    Dim btnCnt As Integer
    Dim msgCnt As Integer
    Dim useloop As Integer
    
    buttonmessage = vbNullString
    foundMessage = False
    btnCnt = 0
    msgCnt = 0
    
    ' populate comboboxes in the general tab
    ' On Error GoTo populateComboBoxes_Error

    cmbRefreshInterval.AddItem "No Timed Refresh", 0
    cmbRefreshInterval.ItemData(0) = 0
    cmbRefreshInterval.AddItem "15 seconds", 1
    cmbRefreshInterval.ItemData(1) = 15
    cmbRefreshInterval.AddItem "30 seconds", 2
    cmbRefreshInterval.ItemData(2) = 30
    cmbRefreshInterval.AddItem "1 minute", 3
    cmbRefreshInterval.ItemData(3) = 60
    cmbRefreshInterval.AddItem "5 minutes", 4
    cmbRefreshInterval.ItemData(4) = 300
    cmbRefreshInterval.AddItem "10 minutes", 5
    cmbRefreshInterval.ItemData(5) = 600
    cmbRefreshInterval.AddItem "30 minutes", 6
    cmbRefreshInterval.ItemData(6) = 1800
    cmbRefreshInterval.AddItem "1 hour", 7
    cmbRefreshInterval.ItemData(7) = 3600
    
    Call populateCmbAlarmSound
    
    ' populate comboboxes in the configuration tab
'    cmbButtonPositions.AddItem "automatic", 0
'    cmbButtonPositions.AddItem "left", 1
'    cmbButtonPositions.AddItem "right", 2

    cmbMaxLineLength.AddItem "20", 0
    cmbMaxLineLength.AddItem "30", 1
    cmbMaxLineLength.AddItem "60", 2
    cmbMaxLineLength.AddItem "72", 3
    cmbMaxLineLength.AddItem "84", 4
    cmbMaxLineLength.AddItem "96", 5
    cmbMaxLineLength.AddItem "108", 6
    cmbMaxLineLength.AddItem "120", 7
    cmbMaxLineLength.AddItem "144", 8
    cmbMaxLineLength.AddItem "168", 9
    cmbMaxLineLength.AddItem "192", 10
    cmbMaxLineLength.AddItem "216", 11
    cmbMaxLineLength.AddItem "240", 12
    
    
    
    ' populate comboboxes in the email tab
    cmbAdviceInterval.AddItem "No interval", 0
    cmbAdviceInterval.ItemData(0) = 0
    cmbAdviceInterval.AddItem "1 minute", 1
    cmbAdviceInterval.ItemData(1) = 60
    cmbAdviceInterval.AddItem "2 minutes", 2
    cmbAdviceInterval.ItemData(2) = 120
    cmbAdviceInterval.AddItem "5 minutes", 3
    cmbAdviceInterval.ItemData(3) = 300
    cmbAdviceInterval.AddItem "10 minutes", 4
    cmbAdviceInterval.ItemData(4) = 600
    cmbAdviceInterval.AddItem "15 minutes", 5
    cmbAdviceInterval.ItemData(5) = 900
    cmbAdviceInterval.AddItem "30 minutes", 6
    cmbAdviceInterval.ItemData(6) = 1800
    cmbAdviceInterval.AddItem "1 hour", 7
    cmbAdviceInterval.ItemData(7) = 3600
    cmbAdviceInterval.AddItem "2 hours", 8
    cmbAdviceInterval.ItemData(8) = 7200
    cmbAdviceInterval.AddItem "5 hours", 9
    cmbAdviceInterval.ItemData(9) = 18000
    cmbAdviceInterval.AddItem "10 hours", 10
    cmbAdviceInterval.ItemData(10) = 36000
    cmbAdviceInterval.AddItem "1 day", 11
    cmbAdviceInterval.ItemData(11) = 86400
    cmbAdviceInterval.AddItem "2 days", 12
    cmbAdviceInterval.ItemData(12) = 172800

    
    
    
    
    
    cmbSmtpAuthenticate.AddItem "None", 0
    cmbSmtpAuthenticate.AddItem "Base 64 encoded", 1
    cmbSmtpAuthenticate.AddItem "NTLM", 2
    
    cmbSmtpSecurity.AddItem "None", 0
    cmbSmtpSecurity.AddItem "STARTTLS", 1
    cmbSmtpSecurity.AddItem "SSL/TLS", 2
    
    ' populate comboboxes in the emojis tab
    If FCWEmojiSetDesc = vbNullString Then FCWEmojiSetDesc = "standard"
    MyPath = App.Path & "\Resources\emojis\"
        
    ' populate the emoji box with any folders that exist
    myName = Dir(MyPath, vbDirectory)   ' Retrieve the first entry.
    Do While myName <> vbNullString   ' Start the loop.
        If myName <> "." And myName <> ".." And myName <> vbNullString Then
            cmbEmojiSet.AddItem myName
        End If
        myName = Dir   ' Get next entry.
    Loop
    
    ' populate comboboxes in the windows tab
    cmbWindowLevel.AddItem "Keep on top of other windows", 0
    cmbWindowLevel.AddItem "Normal", 0
    cmbWindowLevel.AddItem "Keep below all other windows", 0
    
    
    'populate the SMTP configuration slots
    cmbSmtpConfig.AddItem "SMTP Config One", 0
    cmbSmtpConfig.AddItem "SMTP Config Two", 1
    cmbSmtpConfig.AddItem "SMTP Config Three", 2
    cmbSmtpConfig.AddItem "SMTP Config Four", 3
    cmbSmtpConfig.AddItem "SMTP Config Five", 4
    
    'cmbTTFN
    ' loop until none found
    ' read the settings file  button1message1
    ' if an error suppress the error
    ' if message found then add to the combobox

    
    For useloop = 1 To 11
        foundMessage = True
        msgCnt = 0
        Do Until foundMessage = False
            foundMessage = False
            btnCnt = useloop
            msgCnt = msgCnt + 1
            buttonmessage = fGetINISetting("Software\FireCallWin", "button" & btnCnt & "message" & msgCnt, FCWSettingsFile)
            If buttonmessage <> vbNullString Then
                foundMessage = True
                If btnCnt = 1 Then
                    If buttonmessage <> "" Then
                        cmbTTFN.AddItem buttonmessage, 0
                    Else
                        cmbTTFN.AddItem "TTFN, Cheerio!", 0
                    End If
                End If
                If btnCnt = 2 Then
                    If buttonmessage <> "" Then
                        cmbWell.AddItem buttonmessage, 0
                    Else
                        cmbWell.AddItem "How are you today?", 0
                    End If
                End If
                If btnCnt = 3 Then
                    If buttonmessage <> "" Then
                        cmbNews.AddItem buttonmessage, 0
                    Else
                        cmbNews.AddItem "Anything new to tell me?", 0
                    End If
                End If
                If btnCnt = 4 Then
                    If buttonmessage <> "" Then
                        cmbMorn.AddItem buttonmessage, 0
                    Else
                        cmbMorn.AddItem "Morning!", 0
                    End If
                End If
                If btnCnt = 5 Then
                    If buttonmessage <> "" Then
                        cmbWot.AddItem buttonmessage, 0
                    Else
                        cmbWot.AddItem "What's going on in your life?", 0
                    End If
                End If
                If btnCnt = 6 Then cmbWth.AddItem buttonmessage, 0
                If btnCnt = 7 Then cmbPrg.AddItem buttonmessage, 0
                If btnCnt = 8 Then cmbGdn.AddItem buttonmessage, 0
                If btnCnt = 9 Then cmbBusy.AddItem buttonmessage, 0
                If btnCnt = 10 Then cmbCod.AddItem buttonmessage, 0
                If btnCnt = 11 Then cmbOut.AddItem buttonmessage, 0
            End If
        Loop
    Next useloop
    If cmbTTFN.ListCount > 0 Then cmbTTFN.ListIndex = 0
    If cmbWell.ListCount > 0 Then cmbWell.ListIndex = 0
    If cmbNews.ListCount > 0 Then cmbNews.ListIndex = 0
    If cmbMorn.ListCount > 0 Then cmbMorn.ListIndex = 0
    If cmbWot.ListCount > 0 Then cmbWot.ListIndex = 0
    If cmbWth.ListCount > 0 Then cmbWth.ListIndex = 0
    If cmbPrg.ListCount > 0 Then cmbPrg.ListIndex = 0
    If cmbGdn.ListCount > 0 Then cmbGdn.ListIndex = 0
    If cmbBusy.ListCount > 0 Then cmbBusy.ListIndex = 0
    If cmbCod.ListCount > 0 Then cmbCod.ListIndex = 0
    If cmbOut.ListCount > 0 Then cmbOut.ListIndex = 0
    
    cmbCaptureDevices.Text = "No recording devices found" 'dean
    cmbCaptureDevices.Enabled = False
        
    If recordingIsPossible = True Then
        cmbCaptureDevices.Enabled = True
        If FireCallMain.cmbHiddenCaptureDevices.ListCount > 0 Then
            For useloop = 0 To FireCallMain.cmbHiddenCaptureDevices.ListCount - 1
                cmbCaptureDevices.List(useloop) = FireCallMain.cmbHiddenCaptureDevices.List(useloop)
                'MsgBox cmbCaptureDevices.List(useloop)
            Next useloop
        End If
    End If


    cmbArchiveDays.AddItem "10 days", 0
    cmbArchiveDays.ItemData(0) = 864000
    cmbArchiveDays.AddItem "15 days", 1
    cmbArchiveDays.ItemData(1) = 1296000
    cmbArchiveDays.AddItem "30 days", 2
    cmbArchiveDays.ItemData(2) = 2592000
    cmbArchiveDays.AddItem "45 days", 3
    cmbArchiveDays.ItemData(3) = 3888000
    cmbArchiveDays.AddItem "60 days", 4
    cmbArchiveDays.ItemData(4) = 5184000
    cmbArchiveDays.AddItem "75 days", 5
    cmbArchiveDays.ItemData(5) = 6480000
    cmbArchiveDays.AddItem "90 days", 6
    cmbArchiveDays.ItemData(6) = 7776000
    cmbArchiveDays.AddItem "100 days", 7
    cmbArchiveDays.ItemData(7) = 8640000


    On Error GoTo 0
    Exit Sub

populateComboBoxes_Error:

    With err
         If .Number <> 0 Then
            debugLog "Error " & err.Number & " (" & err.Description & ") in procedure populateComboBoxes of Form FireCallPrefs"
            Resume Next
          End If
    End With
                
End Sub




'read the sounds folder and add each WAV file to the combo box
'---------------------------------------------------------------------------------------
' Procedure : populateCmbAlarmSound
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub populateCmbAlarmSound()
    Dim MyPath  As String
    Dim myName As String
    
   On Error GoTo populateCmbAlarmSound_Error

    MyPath = App.Path & "\Resources\sounds\"
    
    ' populate the alarm box with any .wav files that exist
    myName = Dir(MyPath, vbNormal)   ' Retrieve the first entry.
    Do While myName <> vbNullString   ' Start the loop.
       If myName <> "." And myName <> ".." And myName <> vbNullString And fExtractSuffixWithDot(myName) = ".wav" Then
            cmbAlarmSound.AddItem myName
       End If
       myName = Dir   ' Retrieve the next entry.
    Loop

   On Error GoTo 0
   Exit Sub

populateCmbAlarmSound_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure populateCmbAlarmSound of Form FireCallPrefs"

End Sub

' Clicking on the icon inner frame
'---------------------------------------------------------------------------------------
' Procedure : fraConfigInner_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraConfigInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraConfigInner_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

fraConfigInner_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraConfigInner_MouseDown of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraConfigInner_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraConfigInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraConfigInner_MouseMove_Error

    CreateToolTip fraConfigInner.hwnd, "The configuration panel is the location for optional configuration items. These items change how FireCall operates, configure them to suit your needs and your mode of operation.", _
                  TTIconInfo, "Help on Configuration", , , , True

   On Error GoTo 0
   Exit Sub

fraConfigInner_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraConfigInner_MouseMove of Form FireCallPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraConfiguration_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraConfiguration_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraConfiguration_MouseMove_Error

    CreateToolTip fraConfiguration.hwnd, "The configuration panel is the location for optional configuration items. These items change how FireCall operates, configure them to suit your needs and your mode of operation.", _
                  TTIconInfo, "Help on Configuration", , , , True

   On Error GoTo 0
   Exit Sub

fraConfiguration_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraConfiguration_MouseMove of Form FireCallPrefs"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : fraDropbox_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraDropbox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraDropbox_MouseMove_Error

    CreateToolTip fraDropbox.hwnd, "Selecting Dropbox here means that FireCall will look for the Dropbox processes and report an error if they are missing. Uncheck the check box below to suppress the alarm.", _
                  TTIconInfo, "Help on Dropbox Selection", , , , True

   On Error GoTo 0
   Exit Sub

fraDropbox_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraDropbox_MouseMove of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraEmail_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraEmail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraEmail_MouseMove_Error

    CreateToolTip fraEmail.hwnd, "The email panel is where you will configure FCW to work with your email client in order to send email messages containing status and advice.", _
                  TTIconInfo, "Help on Email", , , , True

   On Error GoTo 0
   Exit Sub

fraEmail_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraEmail_MouseMove of Form FireCallPrefs"
End Sub

' Clicking on the icon inner frame
'---------------------------------------------------------------------------------------
' Procedure : fraEmailInner_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraEmailInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraEmailInner_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

fraEmailInner_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraEmailInner_MouseDown of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraEmailInner_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraEmailInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraEmailInner_MouseMove_Error

    CreateToolTip fraEmailInner.hwnd, "The email panel is where you will configure FCW to work with your email client in order to send email messages containing status and advice.", _
                  TTIconInfo, "Help on Email", , , , True

   On Error GoTo 0
   Exit Sub

fraEmailInner_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraEmailInner_MouseMove of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraEmojis_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraEmojis_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraEmojis_MouseMove_Error

    CreateToolTip fraEmojis.hwnd, "Your own Emoji sets can be copied to a folder alongside the standard folder and must have two forms of the emojis within two subfolders, base and telly, both containing emojis of the size, 96x96 pixels.", _
                  TTIconInfo, "Help on Emoji Sets", , , , True

   On Error GoTo 0
   Exit Sub

fraEmojis_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraEmojis_MouseMove of Form FireCallPrefs"
                  
End Sub

' Clicking on the icon inner frame
'---------------------------------------------------------------------------------------
' Procedure : fraEmojisInner_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraEmojisInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraEmojisInner_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

fraEmojisInner_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraEmojisInner_MouseDown of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraEmojisInner_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraEmojisInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraEmojisInner_MouseMove_Error

    CreateToolTip fraEmojisInner.hwnd, "Your own Emoji sets can be copied to a folder alongside the standard folder and must have two forms of the emojis within two subfolders, base and telly, both containing emojis of the size, 96x96 pixels.", _
                  TTIconInfo, "Help on Emoji Sets", , , , True

   On Error GoTo 0
   Exit Sub

fraEmojisInner_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraEmojisInner_MouseMove of Form FireCallPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraFonts_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraFonts_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraFonts_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip fraFonts.hwnd, "For the chat window we suggest Linux Biolinum G at 8pt and Centurion Light SF at 8pt for the config. screen, both of which you will find bundled in the FCW program folder. Choose a small 8pt font for each.", _
                  TTIconInfo, "Help on Setting the Fonts", , , , True

   On Error GoTo 0
   Exit Sub

fraFonts_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraFonts_MouseMove of Form FireCallPrefs"

End Sub

' Clicking on the icon inner frame
'---------------------------------------------------------------------------------------
' Procedure : fraFontsInner_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraFontsInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraFontsInner_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

fraFontsInner_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraFontsInner_MouseDown of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraFontsInner_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraFontsInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraFontsInner_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip fraFontsInner.hwnd, "For the chat window we suggest Linux Biolinum G at 8pt and Centurion Light SF at 8pt for the config. screen, both of which you will find bundled in the FCW program folder. Choose a small 8pt font for each.", _
                  TTIconInfo, "Help on Setting the Fonts", , , , True

   On Error GoTo 0
   Exit Sub

fraFontsInner_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraFontsInner_MouseMove of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraGeneral_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraGeneral_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraGeneral_MouseMove_Error

    CreateToolTip fraGeneral.hwnd, "The General Panel contains the most important user-configurable items required for the program to operate correctly. If these items are not filled in then FireCall will not operate at all.", _
                  TTIconInfo, "Help on Essential Configuration", , , , True

   On Error GoTo 0
   Exit Sub

fraGeneral_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraGeneral_MouseMove of Form FireCallPrefs"
End Sub

' Clicking on the icon inner frame
'---------------------------------------------------------------------------------------
' Procedure : fraGeneralInner_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraGeneralInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraGeneralInner_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

fraGeneralInner_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraGeneralInner_MouseDown of Form FireCallPrefs"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : fraGeneralInner_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraGeneralInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraGeneralInner_MouseMove_Error

    CreateToolTip fraGeneralInner.hwnd, "The General Panel contains the most important user-configurable items required for the program to operate correctly. If these items are not filled in then FireCall will not operate at all.", _
                  TTIconInfo, "Help on Essential Configuration", , , , True

   On Error GoTo 0
   Exit Sub

fraGeneralInner_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraGeneralInner_MouseMove of Form FireCallPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : fraGoogleDrive_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraGoogleDrive_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraGoogleDrive_MouseMove_Error

    CreateToolTip fraGoogleDrive.hwnd, "Selecting Google Drive here means that FireCall will look for the Google Drive processes and report an error if they are missing.", _
                  TTIconInfo, "Help on Google Drive Selection", , , , True

   On Error GoTo 0
   Exit Sub

fraGoogleDrive_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraGoogleDrive_MouseMove of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraHiddenHousekeeping_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraHiddenHousekeeping_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraHiddenHousekeeping_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

fraHiddenHousekeeping_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraHiddenHousekeeping_MouseDown of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraHiddenHousekeeping_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraHiddenHousekeeping_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraHiddenHousekeeping_MouseMove_Error

    CreateToolTip fraHiddenHousekeeping.hwnd, "The housekeeping panel is where you can configure backups and the archiving of old data. The backup functionality is working well but the archiving has not yet been implemented.", _
                  TTIconInfo, "Help on Housekeeping", , , , True

   On Error GoTo 0
   Exit Sub

fraHiddenHousekeeping_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraHiddenHousekeeping_MouseMove of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraHousekeeping_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraHousekeeping_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraHousekeeping_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

fraHousekeeping_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraHousekeeping_MouseDown of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraHousekeeping_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraHousekeeping_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraHousekeeping_MouseMove_Error

    CreateToolTip fraHousekeeping.hwnd, "The housekeeping panel is where you can configure backups and the archiving of old data. The backup functionality is working well but the archiving has not yet been implemented.", _
                  TTIconInfo, "Help on Housekeeping", , , , True

   On Error GoTo 0
   Exit Sub

fraHousekeeping_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraHousekeeping_MouseMove of Form FireCallPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : fraNone_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraNone_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraNone_MouseMove_Error

    CreateToolTip fraNone.hwnd, "Selecting None - FireCall will not look for any processes. This implies you are using your own network for internal file sharing.", _
                  TTIconInfo, "Help on OneDrive Selection", , , , True

   On Error GoTo 0
   Exit Sub

fraNone_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraNone_MouseMove of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraOneDrive_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraOneDrive_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraOneDrive_MouseMove_Error

    CreateToolTip fraOneDrive.hwnd, "Selecting OneDrive here means that FireCall will look for the OneDrive processes and report an error if they are missing.", _
                  TTIconInfo, "Help on OneDrive Selection", , , , True

   On Error GoTo 0
   Exit Sub

fraOneDrive_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraOneDrive_MouseMove of Form FireCallPrefs"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : fraSMTPframe_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraSMTPframe_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraSMTPframe_MouseMove_Error

    CreateToolTip fraSMTPframe.hwnd, "Messages are sent by email using the SMTP details entered.  Extract these from your email client, Outlook or Thunderbird for example.", _
                  TTIconInfo, "Help on SMTP Server", , , , True

   On Error GoTo 0
   Exit Sub

fraSMTPframe_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraSMTPframe_MouseMove of Form FireCallPrefs"
End Sub

' Clicking on the icon inner frame
'---------------------------------------------------------------------------------------
' Procedure : fraSounds_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraSounds_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraSounds_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

fraSounds_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraSounds_MouseDown of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraSounds_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraSounds_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

   On Error GoTo fraSounds_MouseMove_Error

    CreateToolTip fraSounds.hwnd, "The sound panel allows you to configure the sounds that occur within FCW. Some of the animations have associated sounds, you can control these here..", _
                  TTIconInfo, "Help on Configuring Sounds", , , , True

   On Error GoTo 0
   Exit Sub

fraSounds_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraSounds_MouseMove of Form FireCallPrefs"

End Sub

' Clicking on the icon inner frame
'---------------------------------------------------------------------------------------
' Procedure : fraSoundsInner_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraSoundsInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraSoundsInner_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

fraSoundsInner_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraSoundsInner_MouseDown of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraSoundsInner_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraSoundsInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraSoundsInner_MouseMove_Error

    CreateToolTip fraSoundsInner.hwnd, "The sound panel allows you to configure the sounds that occur within FCW. Some of the animations have associated sounds, you can control these here..", _
                  TTIconInfo, "Help on Configuring Sounds", , , , True

   On Error GoTo 0
   Exit Sub

fraSoundsInner_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraSoundsInner_MouseMove of Form FireCallPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraTargetClient_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraTargetClient_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraTargetClient_MouseMove_Error

    CreateToolTip fraTargetClient.hwnd, "The target client is an option that you may not need to change unless you are communicating with the javascript version of the FireCall app that runs on Mac OS X. That version requires UTF8 support to display and handle unicode characters. If you are a Windows user communicating with FireCall for Windows you do not need to select the UTF8 option. However, the code we use to handle UTF8 files may be faster for reading and writing the input/output data files, so by all means try it out.", _
                  TTIconInfo, "Help on Selecting ANSI or UTF8", , , , True

   On Error GoTo 0
   Exit Sub

fraTargetClient_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraTargetClient_MouseMove of Form FireCallPrefs"

End Sub

' Clicking on the icon inner frame
'---------------------------------------------------------------------------------------
' Procedure : fraTexts_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraTexts_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraTexts_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

fraTexts_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraTexts_MouseDown of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraTexts_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraTexts_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraTexts_MouseMove_Error

    CreateToolTip fraTexts.hwnd, "The texts panel is where you can configure the pre-programmed messages that FCW can send using the buttons at the bottom of the utility. This panel allows you to change or add to the pre-defined texts that appear on the buttons.", _
                  TTIconInfo, "Help on PreDefined Texts", , , , True

   On Error GoTo 0
   Exit Sub

fraTexts_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraTexts_MouseMove of Form FireCallPrefs"
End Sub

' Clicking on the icon inner frame
'---------------------------------------------------------------------------------------
' Procedure : fraTextsInner_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraTextsInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraTextsInner_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

fraTextsInner_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraTextsInner_MouseDown of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraTextsInner_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraTextsInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraTextsInner_MouseMove_Error

    CreateToolTip fraTextsInner.hwnd, "The texts panel is where you can configure the pre-programmed messages that FCW can send using the buttons at the bottom of the utility. This panel allows you to change or add to the pre-defined texts that appear on the buttons.", _
                  TTIconInfo, "Help on PreDefined Texts", , , , True

   On Error GoTo 0
   Exit Sub

fraTextsInner_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraTextsInner_MouseMove of Form FireCallPrefs"

End Sub

' Clicking on the icon inner frame
'---------------------------------------------------------------------------------------
' Procedure : fraWindow_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraWindow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraWindow_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

fraWindow_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraWindow_MouseDown of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraWindow_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraWindow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraWindow_MouseMove_Error

    CreateToolTip fraWindow.hwnd, "The Opacity and Window Level of the program are rather strange characteristics to change in a Windows program, however Fire Call Win is a copy of a Yahoo Widget of the same name. All widgets have similar window tab options including the capability to change the opacity and window level. Whether these options are useful to you or anyone is a moot point but as this tool aims to replicate the YWE version functionality it has been reproduced here. It is here as more of an experiment as to how to implement a feature, one carried over from the Yahoo Widget (javascript) version of this program.", _
                  TTIconInfo, "Help on YWE Quirk Mode Options", , , , True

   On Error GoTo 0
   Exit Sub

fraWindow_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraWindow_MouseMove of Form FireCallPrefs"
End Sub

' Clicking on the icon inner frame
'---------------------------------------------------------------------------------------
' Procedure : fraWindowInner_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraWindowInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraWindowInner_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

fraWindowInner_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraWindowInner_MouseDown of Form FireCallPrefs"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : lblAllowShutdowns_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblAllowShutdowns_Click()
   On Error GoTo lblAllowShutdowns_Click_Error

    If chkAllowShutdowns.Value = 1 Then
        chkAllowShutdowns.Value = 0
    Else
        chkAllowShutdowns.Value = 1
    End If

   On Error GoTo 0
   Exit Sub

lblAllowShutdowns_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lblAllowShutdowns_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraWindowInner_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraWindowInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraWindowInner_MouseMove_Error

    CreateToolTip fraWindowInner.hwnd, "The Opacity and Window Level of the program are rather strange characteristics to change in a Windows program, however Fire Call Win is a copy of a Yahoo Widget of the same name. All widgets have similar window tab options including the capability to change the opacity and window level. Whether these options are useful to you or anyone is a moot point but as this tool aims to replicate the YWE version functionality it has been reproduced here. It is here as more of an experiment as to how to implement a feature, one carried over from the Yahoo Widget (javascript) version of this program.", _
                  TTIconInfo, "Help on YWE Quirk Mode Options", , , , True

   On Error GoTo 0
   Exit Sub

fraWindowInner_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraWindowInner_MouseMove of Form FireCallPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : lblAutomaticHousekeeping_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblAutomaticHousekeeping_Click()
   On Error GoTo lblAutomaticHousekeeping_Click_Error

    If chkAutomaticHousekeeping.Value = 1 Then
        chkAutomaticHousekeeping.Value = 0
    Else
        chkAutomaticHousekeeping.Value = 1
    End If

   On Error GoTo 0
   Exit Sub

lblAutomaticHousekeeping_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lblAutomaticHousekeeping_Click of Form FireCallPrefs"
End Sub












' clicking upon the labels below the main prefs icons
'---------------------------------------------------------------------------------------
' Procedure : lblEmojis_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblEmojis_Click()
   On Error GoTo lblEmojis_Click_Error

    Call picEmojis_Click

   On Error GoTo 0
   Exit Sub

lblEmojis_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lblEmojis_Click of Form FireCallPrefs"
End Sub



' clicking upon the labels below the main prefs icons
'---------------------------------------------------------------------------------------
' Procedure : lblFonts_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblFonts_Click()
   On Error GoTo lblFonts_Click_Error

    Call picFonts_Click

   On Error GoTo 0
   Exit Sub

lblFonts_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lblFonts_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : lbloptServiceProvider_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbloptServiceProvider_Click(Index As Integer)
   On Error GoTo lbloptServiceProvider_Click_Error

    optServiceProvider(Index).Value = True

   On Error GoTo 0
   Exit Sub

lbloptServiceProvider_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbloptServiceProvider_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fraServiceProvider_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraServiceProvider_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo fraServiceProvider_MouseMove_Error

    CreateToolTip fraServiceProvider.hwnd, "Select which utility you are using to share the files and folders. Dependant upon which selection you choose, Fire Call for Windows will check if the processes associated with the chosen utility are running. This setting will not change the shared folder location, you'll have to do that yourself using the text fields above. If you are not using a service provider and instead just sharing files over a network then select - none", _
                  TTIconInfo, "Help on Selecting a Service Provider", , , , True

   On Error GoTo 0
   Exit Sub

fraServiceProvider_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraServiceProvider_MouseMove of Form FireCallPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : lblPlayVolume_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblPlayVolume_Click()
   On Error GoTo lblPlayVolume_Click_Error

    If chkPlayVolume.Value = 1 Then
        chkPlayVolume.Value = 0
    Else
        chkPlayVolume.Value = 1
    End If

   On Error GoTo 0
   Exit Sub

lblPlayVolume_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lblPlayVolume_Click of Form FireCallPrefs"
End Sub





' clicking upon the labels below the main prefs icons
'---------------------------------------------------------------------------------------
' Procedure : lblTexts_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblTexts_Click()
   On Error GoTo lblTexts_Click_Error

    Call picTexts_Click

   On Error GoTo 0
   Exit Sub

lblTexts_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lblTexts_Click of Form FireCallPrefs"
End Sub

' clicking upon the labels below the main prefs icons
'---------------------------------------------------------------------------------------
' Procedure : lblWindow_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblWindow_Click()
   On Error GoTo lblWindow_Click_Error

    Call picWindow_Click

   On Error GoTo 0
   Exit Sub

lblWindow_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lblWindow_Click of Form FireCallPrefs"
End Sub

' clicking upon the labels below the main prefs icons
'---------------------------------------------------------------------------------------
' Procedure : lblSounds_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblSounds_Click()
   On Error GoTo lblSounds_Click_Error

    Call picSounds_Click

   On Error GoTo 0
   Exit Sub

lblSounds_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lblSounds_Click of Form FireCallPrefs"
End Sub

' clicking upon the labels below the main prefs icons
'---------------------------------------------------------------------------------------
' Procedure : lblConfig_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblConfig_Click()
   On Error GoTo lblConfig_Click_Error

    Call picConfig_Click

   On Error GoTo 0
   Exit Sub

lblConfig_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lblConfig_Click of Form FireCallPrefs"
End Sub

' clicking upon the labels below the main prefs icons
'---------------------------------------------------------------------------------------
' Procedure : lblEmail_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblEmail_Click()
   On Error GoTo lblEmail_Click_Error

    Call picEmail_Click

   On Error GoTo 0
   Exit Sub

lblEmail_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lblEmail_Click of Form FireCallPrefs"
End Sub
' clicking upon the labels below the main prefs icons
'---------------------------------------------------------------------------------------
' Procedure : lblGeneral_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblGeneral_Click()
   On Error GoTo lblGeneral_Click_Error

    Call picGeneral_Click

   On Error GoTo 0
   Exit Sub

lblGeneral_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lblGeneral_Click of Form FireCallPrefs"
End Sub



' removes all styling from the icon frames and makes the major frames below invisible too
'---------------------------------------------------------------------------------------
' Procedure : clearBorderStyle
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub clearBorderStyle()

   On Error GoTo clearBorderStyle_Error

    fraGeneral.Visible = False
    fraConfiguration.Visible = False
    fraEmail.Visible = False
    fraEmojis.Visible = False
    fraFonts.Visible = False
    fraWindow.Visible = False
    fraTexts.Visible = False
    fraHousekeeping.Visible = False
    fraSounds.Visible = False

    fraConfigurationButtons.BorderStyle = 0
    fraGeneralButtons.BorderStyle = 0
    fraEmailButtons.BorderStyle = 0
    fraEmojisButtons.BorderStyle = 0
    fraFontsButtons.BorderStyle = 0
    fraWindowButton.BorderStyle = 0
    fraTextsButton.BorderStyle = 0
    fraHousekeepingButton.BorderStyle = 0
    fraSoundsButton.BorderStyle = 0

   On Error GoTo 0
   Exit Sub

clearBorderStyle_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure clearBorderStyle of Form FireCallPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : optHandleData_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub optHandleData_Click(Index As Integer)
   On Error GoTo optHandleData_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

optHandleData_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure optHandleData_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : optWindowWidth_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub optWindowWidth_Click(Index As Integer)
   On Error GoTo optWindowWidth_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

optWindowWidth_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure optWindowWidth_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : optHandleData_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub optHandleData_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
   On Error GoTo optHandleData_MouseMove_Error

    CreateToolTip optHandleData(Index).hwnd, "The target client is an option that you may not need to change unless you are communicating with the javascript version of the FireCall app that runs on Mac OS X. That version requires UTF8 support to display and handle unicode characters. If you are a Windows user communicating with FireCall for Windows you do not need to select the UTF8 option. However, the code we use to handle UTF8 files may be faster for reading and writing the input/output data files, so by all means try it out. The first uses the File System Object to read and write text, whereas the second uses an ADO record stream to write UTF8 compatible files.", _
                  TTIconInfo, "Help on Selecting ANSI or UTF8", , , , True

   On Error GoTo 0
   Exit Sub

optHandleData_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure optHandleData_MouseMove of Form FireCallPrefs"
End Sub
'
'Private Sub optRecordingType_Click(Index As Integer)
'    btnSave.Enabled = True ' enable the save button
'    If Index = 0 Then cmbCaptureDevices.Enabled = False
'    If Index = 1 Then cmbCaptureDevices.Enabled = True
'
'
'
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : optServiceProvider_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub optServiceProvider_Click(Index As Integer)
   On Error GoTo optServiceProvider_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

optServiceProvider_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure optServiceProvider_Click of Form FireCallPrefs"
End Sub

'Private Sub optServiceProvider_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    CreateToolTip Me.hWnd, "If you are going to be sharing files over Dropbox's network then select Dropbox, FCW will then check for the existence of the Dropbox processes and will report an error if they are missing.", _
'                  TTIconInfo, "Help on Dropbox as a Service Provider", , , , True
'End Sub



'---------------------------------------------------------------------------------------
' Procedure : picHouseKeeping_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picHouseKeeping_Click()
   On Error GoTo picHouseKeeping_Click_Error

    Call clearBorderStyle

    fraHousekeeping.Visible = True
    fraHousekeepingButton.BorderStyle = 1
    FCWLastSelectedTab = "housekeeping"
    PutINISetting "Software\FireCallWin", "lastSelectedTab", FCWLastSelectedTab, FCWSettingsFile

    FireCallPrefs.Height = fraHousekeeping.Height + 2000
    btnSave.Top = fraHousekeeping.Top + fraHousekeeping.Height + 100
    btnCancel.Top = fraHousekeeping.Top + fraHousekeeping.Height + 100
    btnHelp.Top = fraHousekeeping.Top + fraHousekeeping.Height + 100

   On Error GoTo 0
   Exit Sub

picHouseKeeping_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picHouseKeeping_Click of Form FireCallPrefs"
End Sub

' clicking on the config icon
'---------------------------------------------------------------------------------------
' Procedure : picConfig_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picConfig_Click()

   On Error GoTo picConfig_Click_Error

    Call clearBorderStyle
    fraConfiguration.Visible = True
    FCWLastSelectedTab = "config"
    PutINISetting "Software\FireCallWin", "lastSelectedTab", FCWLastSelectedTab, FCWSettingsFile
    
    fraConfigurationButtons.BorderStyle = 1
    
    FireCallPrefs.Height = fraConfiguration.Height + 2000
    btnSave.Top = fraConfiguration.Top + fraConfiguration.Height + 100
    btnCancel.Top = fraConfiguration.Top + fraConfiguration.Height + 100
    btnHelp.Top = fraConfiguration.Top + fraConfiguration.Height + 100

   On Error GoTo 0
   Exit Sub

picConfig_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picConfig_Click of Form FireCallPrefs"
    
End Sub

' clicking on the email icon
'---------------------------------------------------------------------------------------
' Procedure : picEmail_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picEmail_Click()

   On Error GoTo picEmail_Click_Error

    Call clearBorderStyle
    
    fraEmail.Visible = True
    fraEmailButtons.BorderStyle = 1
    FCWLastSelectedTab = "email"
    PutINISetting "Software\FireCallWin", "lastSelectedTab", FCWLastSelectedTab, FCWSettingsFile
    
    FireCallPrefs.Height = fraEmail.Height + 2000
    btnSave.Top = fraEmail.Top + fraEmail.Height + 100
    btnCancel.Top = fraEmail.Top + fraEmail.Height + 100
    btnHelp.Top = fraEmail.Top + fraEmail.Height + 100

   On Error GoTo 0
   Exit Sub

picEmail_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picEmail_Click of Form FireCallPrefs"
        
End Sub
' clicking on the emojis icon
'---------------------------------------------------------------------------------------
' Procedure : picEmojis_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picEmojis_Click()

   On Error GoTo picEmojis_Click_Error

    Call clearBorderStyle
    fraEmojis.Visible = True
    fraEmojisButtons.BorderStyle = 1
    FCWLastSelectedTab = "emojis"
    PutINISetting "Software\FireCallWin", "lastSelectedTab", FCWLastSelectedTab, FCWSettingsFile
    
    FireCallPrefs.Height = fraEmojis.Height + 2000
    btnSave.Top = fraEmojis.Top + fraEmojis.Height + 100
    btnCancel.Top = fraEmojis.Top + fraEmojis.Height + 100
    btnHelp.Top = fraEmojis.Top + fraEmojis.Height + 100

   On Error GoTo 0
   Exit Sub

picEmojis_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picEmojis_Click of Form FireCallPrefs"
    
End Sub

' clicking on the fonts icon
'---------------------------------------------------------------------------------------
' Procedure : picFonts_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picFonts_Click()
   On Error GoTo picFonts_Click_Error

    Call clearBorderStyle
    
    fraFonts.Visible = True
    fraFontsButtons.BorderStyle = 1
    FCWLastSelectedTab = "fonts"
    PutINISetting "Software\FireCallWin", "lastSelectedTab", FCWLastSelectedTab, FCWSettingsFile
    
    FireCallPrefs.Height = fraFonts.Height + 2000
    btnSave.Top = fraFonts.Top + fraFonts.Height + 100
    btnCancel.Top = fraFonts.Top + fraFonts.Height + 100
    btnHelp.Top = fraFonts.Top + fraFonts.Height + 100

   On Error GoTo 0
   Exit Sub

picFonts_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picFonts_Click of Form FireCallPrefs"
    
End Sub
' clicking on the general icon
'---------------------------------------------------------------------------------------
' Procedure : picGeneral_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picGeneral_Click()

   On Error GoTo picGeneral_Click_Error

    Call clearBorderStyle
    fraGeneralButtons.BorderStyle = 1
    
    fraGeneral.Visible = True
    FCWLastSelectedTab = "general"
    PutINISetting "Software\FireCallWin", "lastSelectedTab", FCWLastSelectedTab, FCWSettingsFile
    
    FireCallPrefs.Height = fraGeneral.Height + 2000
    btnSave.Top = fraGeneral.Top + fraGeneral.Height + 100
    btnCancel.Top = fraGeneral.Top + fraGeneral.Height + 100
    btnHelp.Top = fraGeneral.Top + fraGeneral.Height + 100
    

   On Error GoTo 0
   Exit Sub

picGeneral_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picGeneral_Click of Form FireCallPrefs"
    
End Sub
' clicking on the sounds icon
'---------------------------------------------------------------------------------------
' Procedure : picSounds_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picSounds_Click()
   On Error GoTo picSounds_Click_Error

    Call clearBorderStyle

    fraSounds.Visible = True
    fraSoundsButton.BorderStyle = 1
    FCWLastSelectedTab = "sounds"
    PutINISetting "Software\FireCallWin", "lastSelectedTab", FCWLastSelectedTab, FCWSettingsFile
    
    FireCallPrefs.Height = fraSounds.Height + 2000
    btnSave.Top = fraSounds.Top + fraSounds.Height + 100
    btnCancel.Top = fraSounds.Top + fraSounds.Height + 100
    btnHelp.Top = fraSounds.Top + fraSounds.Height + 100

   On Error GoTo 0
   Exit Sub

picSounds_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picSounds_Click of Form FireCallPrefs"
    
End Sub
' clicking on the texts icon
'---------------------------------------------------------------------------------------
' Procedure : picTexts_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picTexts_Click()
    
   On Error GoTo picTexts_Click_Error

    Call clearBorderStyle

    fraTexts.Visible = True
    fraTextsButton.BorderStyle = 1
    FCWLastSelectedTab = "texts"
    PutINISetting "Software\FireCallWin", "lastSelectedTab", FCWLastSelectedTab, FCWSettingsFile
    
    FireCallPrefs.Height = fraTexts.Height + 2000
    btnSave.Top = fraTexts.Top + fraTexts.Height + 100
    btnCancel.Top = fraTexts.Top + fraTexts.Height + 100
    btnHelp.Top = fraTexts.Top + fraTexts.Height + 100

   On Error GoTo 0
   Exit Sub

picTexts_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picTexts_Click of Form FireCallPrefs"
    
End Sub





'---------------------------------------------------------------------------------------
' Procedure : Picture_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Picture_Click(Index As Integer)
   On Error GoTo Picture_Click_Error

    fraEmailfra.Visible = False

   On Error GoTo 0
   Exit Sub

Picture_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure Picture_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Picture_DblClick
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Picture_DblClick(Index As Integer)
   On Error GoTo Picture_DblClick_Error

    fraEmailfra.Visible = False

   On Error GoTo 0
   Exit Sub

Picture_DblClick_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure Picture_DblClick of Form FireCallPrefs"
End Sub

' clicking on the windows icon
'---------------------------------------------------------------------------------------
' Procedure : picWindow_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picWindow_Click()
    
   On Error GoTo picWindow_Click_Error

    Call clearBorderStyle
    fraWindow.Visible = True
    fraWindowButton.BorderStyle = 1
    FCWLastSelectedTab = "window"
    PutINISetting "Software\FireCallWin", "lastSelectedTab", FCWLastSelectedTab, FCWSettingsFile
    
    FireCallPrefs.Height = fraWindow.Height + 2000
    btnSave.Top = fraWindow.Top + fraWindow.Height + 100
    btnCancel.Top = fraWindow.Top + fraWindow.Height + 100
    btnHelp.Top = fraWindow.Top + fraWindow.Height + 100

   On Error GoTo 0
   Exit Sub

picWindow_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picWindow_Click of Form FireCallPrefs"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliAutomaticBackupInterval_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliAutomaticBackupInterval_Click()
   On Error GoTo sliAutomaticBackupInterval_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

sliAutomaticBackupInterval_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure sliAutomaticBackupInterval_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliIconiseDelay_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliIconiseDelay_Click()
   On Error GoTo sliIconiseDelay_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

sliIconiseDelay_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure sliIconiseDelay_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliOpacity_Change
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliOpacity_Change()
   On Error GoTo sliOpacity_Change_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

sliOpacity_Change_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure sliOpacity_Change of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliOpacity_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliOpacity_Click()
   On Error GoTo sliOpacity_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

sliOpacity_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure sliOpacity_Click of Form FireCallPrefs"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : sliRecordingQuality_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliRecordingQuality_Click()
   On Error GoTo sliRecordingQuality_Click_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

sliRecordingQuality_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure sliRecordingQuality_Click of Form FireCallPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : txtEmailMessage_Change
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtEmailMessage_Change()
   On Error GoTo txtEmailMessage_Change_Error

    btnTestEmail.Enabled = False
    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

txtEmailMessage_Change_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtEmailMessage_Change of Form FireCallPrefs"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : txtEmailSubject_Change
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtEmailSubject_Change()
   On Error GoTo txtEmailSubject_Change_Error

    btnTestEmail.Enabled = False
    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

txtEmailSubject_Change_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtEmailSubject_Change of Form FireCallPrefs"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : txtExchangeFolder_Change
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtExchangeFolder_Change()
   On Error GoTo txtExchangeFolder_Change_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

txtExchangeFolder_Change_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtExchangeFolder_Change of Form FireCallPrefs"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : txtFontSize_Change
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtFontSize_Change()
   On Error GoTo txtFontSize_Change_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

txtFontSize_Change_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtFontSize_Change of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtSmtpConfigName_Change
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtSmtpConfigName_Change()
   On Error GoTo txtSmtpConfigName_Change_Error

    btnTestEmail.Enabled = False
    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

txtSmtpConfigName_Change_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtSmtpConfigName_Change of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtSMTPNoPassword_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtSMTPNoPassword_Click()
   On Error GoTo txtSMTPNoPassword_Click_Error

    MsgBox "Please press the 'show password' button to amend the password details."

   On Error GoTo 0
   Exit Sub

txtSMTPNoPassword_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtSMTPNoPassword_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtSMTPPassword_Change
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtSMTPPassword_Change()
    Dim i As Integer
   On Error GoTo txtSMTPPassword_Change_Error

    btnSave.Enabled = True ' enable the save button
    btnTestEmail.Enabled = False
    txtSMTPNoPassword.Text = String$(Len(txtSmtpPassword.Text), "*")

   On Error GoTo 0
   Exit Sub

txtSMTPPassword_Change_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtSMTPPassword_Change of Form FireCallPrefs"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtPop3Server_Change
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtPop3Server_Change()
   On Error GoTo txtPop3Server_Change_Error

    btnTestEmail.Enabled = False
    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

txtPop3Server_Change_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtPop3Server_Change of Form FireCallPrefs"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtSMTPUsername_Change
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtSMTPUsername_Change()
   On Error GoTo txtSMTPUsername_Change_Error

    btnTestEmail.Enabled = False
    btnSave.Enabled = True ' enable the save button
    

   On Error GoTo 0
   Exit Sub

txtSMTPUsername_Change_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtSMTPUsername_Change of Form FireCallPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtPrefixString_Change
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtPrefixString_Change()
   On Error GoTo txtPrefixString_Change_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

txtPrefixString_Change_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtPrefixString_Change of Form FireCallPrefs"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : txtPrefsFont_Change
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtPrefsFont_Change()
   On Error GoTo txtPrefsFont_Change_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

txtPrefsFont_Change_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtPrefsFont_Change of Form FireCallPrefs"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : txtRecipientEmail_Change
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtRecipientEmail_Change()
   On Error GoTo txtRecipientEmail_Change_Error

    btnTestEmail.Enabled = False
    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

txtRecipientEmail_Change_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtRecipientEmail_Change of Form FireCallPrefs"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : txtSharedInputFile_Change
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtSharedInputFile_Change()
   On Error GoTo txtSharedInputFile_Change_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

txtSharedInputFile_Change_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtSharedInputFile_Change of Form FireCallPrefs"
End Sub
' check file existence when the user presses carriage return after manually typing a filename
'---------------------------------------------------------------------------------------
' Procedure : txtSharedInputFile_KeyPress
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtSharedInputFile_KeyPress(ByRef KeyAscii As Integer)
    Dim answer As VbMsgBoxResult

    ' check for a CR, set the keyascii to 0 to prevent the beeps
   On Error GoTo txtSharedInputFile_KeyPress_Error

    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not fFExists(txtSharedInputFile.Text) Then
            answer = MsgBox("The file doesn't currently exist, do you want me to create the chosen file, " & "   -  are you sure?", vbYesNo)
            If answer = vbNo Then
                Exit Sub
            End If
        
            'create new
            Open txtSharedInputFile.Text For Output As #1
            Close #1
        End If
    End If

   On Error GoTo 0
   Exit Sub

txtSharedInputFile_KeyPress_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtSharedInputFile_KeyPress of Form FireCallPrefs"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : txtSharedOutputFile_Change
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtSharedOutputFile_Change()
   On Error GoTo txtSharedOutputFile_Change_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

txtSharedOutputFile_Change_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtSharedOutputFile_Change of Form FireCallPrefs"
End Sub
' check file existence when the user presses carriage return after manually typing a filename
'---------------------------------------------------------------------------------------
' Procedure : txtSharedOutputFile_KeyPress
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtSharedOutputFile_KeyPress(ByRef KeyAscii As Integer)
    Dim answer As VbMsgBoxResult

    ' check for a CR, set the keyascii to 0 to prevent the beeps
   On Error GoTo txtSharedOutputFile_KeyPress_Error

    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not fFExists(txtSharedOutputFile.Text) Then
            answer = MsgBox("The file doesn't currently exist, do you want me to create the chosen file, " & "   -  are you sure?", vbYesNo)
            If answer = vbNo Then
                Exit Sub
            End If
        
            'create new
            Open txtSharedOutputFile.Text For Output As #1
            Close #1
        End If
    End If

   On Error GoTo 0
   Exit Sub

txtSharedOutputFile_KeyPress_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtSharedOutputFile_KeyPress of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtSmtpPort_Change
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtSmtpPort_Change()
   On Error GoTo txtSmtpPort_Change_Error

    btnTestEmail.Enabled = False
    btnSave.Enabled = True ' enable the save button
    

   On Error GoTo 0
   Exit Sub

txtSmtpPort_Change_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtSmtpPort_Change of Form FireCallPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtSmtpServer_Change
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtSmtpServer_Change()
   On Error GoTo txtSmtpServer_Change_Error

    btnSave.Enabled = True ' enable the save button
    btnTestEmail.Enabled = False

   On Error GoTo 0
   Exit Sub

txtSmtpServer_Change_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtSmtpServer_Change of Form FireCallPrefs"
End Sub



' add new user defined text to the pre-defined buttons
'---------------------------------------------------------------------------------------
' Procedure : txtStringToAdd_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtStringToAdd_Click()
   On Error GoTo txtStringToAdd_Click_Error

    btnSave.Enabled = True ' enable the save button

    If txtStringToAdd.Text = "Enter text here and click + button below" Then txtStringToAdd.Text = vbNullString
    txtStringToAddFieldModified = True

   On Error GoTo 0
   Exit Sub

txtStringToAdd_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtStringToAdd_Click of Form FireCallPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtTextFont_Change
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtTextFont_Change()
   On Error GoTo txtTextFont_Change_Error

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

txtTextFont_Change_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtTextFont_Change of Form FireCallPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : themeTimer_Timer
' Author    : beededea
' Date      : 13/06/2020
' Purpose   : a timer to apply a theme automatically
'---------------------------------------------------------------------------------------
'
Private Sub themeTimer_Timer()
        
    ' variables declared
    Dim SysClr As Long
        
    'initialise the dimensioned variables
    SysClr = 0

   ' On Error GoTo themeTimer_Timer_Error

    SysClr = GetSysColor(COLOR_BTNFACE)
    If debugflg = 1 Then Debug.Print "COLOR_BTNFACE = " & SysClr ' generates too many debug statements in the log
    If SysClr <> storeThemeColour Then
    
        Call setThemeColour

    End If

   On Error GoTo 0
   Exit Sub

themeTimer_Timer_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure themeTimer_Timer of Form dockSettings"

End Sub
' show the about us form
'---------------------------------------------------------------------------------------
' Procedure : mnuAboutFireCallWin_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAboutFireCallWin_Click()
   On Error GoTo mnuAboutFireCallWin_Click_Error

    about.Show

   On Error GoTo 0
   Exit Sub

mnuAboutFireCallWin_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuAboutFireCallWin_Click of Form FireCallPrefs"
End Sub

' The menu options are replicated on the prefs form as well, it seems we cannot easily share menu options
' between forms.

' open the shared input file using the default application
'---------------------------------------------------------------------------------------
' Procedure : mnuOpenSharedInputFile_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuOpenSharedInputFile_Click()
   On Error GoTo mnuOpenSharedInputFile_Click_Error

            Call ShellExecute(Me.hwnd, "Open", FCWSharedInputFile, vbNullString, App.Path, 1)

   On Error GoTo 0
   Exit Sub

mnuOpenSharedInputFile_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuOpenSharedInputFile_Click of Form FireCallPrefs"
End Sub

' open the shared output file using the default application
'---------------------------------------------------------------------------------------
' Procedure : mnuOpenSharedOutputFile_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuOpenSharedOutputFile_Click()
   On Error GoTo mnuOpenSharedOutputFile_Click_Error

            Call ShellExecute(Me.hwnd, "Open", FCWSharedOutputFile, vbNullString, App.Path, 1)

   On Error GoTo 0
   Exit Sub

mnuOpenSharedOutputFile_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuOpenSharedOutputFile_Click of Form FireCallPrefs"
End Sub
' open the shared folder using the file explorer
'---------------------------------------------------------------------------------------
' Procedure : mnuOpenSharedExchangeFolder_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuOpenSharedExchangeFolder_Click()
   On Error GoTo mnuOpenSharedExchangeFolder_Click_Error

            Call ShellExecute(Me.hwnd, "Open", FCWExchangeFolder, vbNullString, App.Path, 1)

   On Error GoTo 0
   Exit Sub

mnuOpenSharedExchangeFolder_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuOpenSharedExchangeFolder_Click of Form FireCallPrefs"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuCoffee_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   : main menu item to buy the developer a coffee
'---------------------------------------------------------------------------------------
'
Private Sub mnuCoffee_Click()
        
    ' variables declared
    Dim answer As VbMsgBoxResult
    
    'initialise the dimensioned variables
    answer = vbNo
    
    ' On Error GoTo mnuCoffee_Click_Error
    If debugflg = 1 Then Debug.Print "%" & "mnuCoffee_Click"
    
    answer = MsgBox(" Help support the creation of more widgets like this, send us a beer! This button opens a browser window and connects to the Paypal donate page for this widget). Will you be kind and proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "https://www.paypal.com/cgi-bin/webscr?cmd=_xclick&business=info@lightquick.co.uk&currency_code=GBP&amount=2.50&return=&item_name=Donate%20a%20Beer", vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub
mnuCoffee_Click_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure mnuCoffee_Click of Form quartermaster"
End Sub


' Error handling
'    On Error GoTo err:
    
'    Exit Function
'err:
'    With err
'         If .Number <> 0 Then
'            'create .bas named [ErrHandler]  see http://vb6.info/h764u
'            ErrHandler.ReportError Date & ": Strings.bMultiInstr." & err.Number & "." & err.Description
'            Resume Next
'          End If
'    End With


' menu option to show licence
'---------------------------------------------------------------------------------------
' Procedure : mnuLicenceA_Click
' Author    : beededea
' Date      : 17/08/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLicenceA_Click()
    ' On Error GoTo mnuLicenceA_Click_Error

    Call LoadFileToTB(licence.txtLicenceTextBox, App.Path & "\licence.txt", False)
    licence.Show

    On Error GoTo 0
    Exit Sub

mnuLicenceA_Click_Error:

    With err
         If .Number <> 0 Then
            debugLog "Error " & err.Number & " (" & err.Description & ") in procedure mnuLicenceA_Click of Form FireCallPrefs"
            Resume Next
          End If
    End With

End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuSupport_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   : menu option to open support page
'---------------------------------------------------------------------------------------
'
Private Sub mnuSupport_Click()
        
    ' variables declared
    Dim answer As VbMsgBoxResult
    
    'initialise the dimensioned variables
    answer = vbNo
    
    ' On Error GoTo mnuSupport_Click_Error
    If debugflg = 1 Then Debug.Print "%" & "mnuSupport_Click"

    answer = MsgBox("Visiting the support page - this button opens a browser window and connects to our contact us page where you can send us a support query or just have a chat). Proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "https://www.deviantart.com/yereverluvinuncleber/art/Quartermaster-VB6-Desktop-784624943", vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub

mnuSupport_Click_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure mnuSupport_Click of Form quartermaster"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuSweets_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuSweets_Click()
        
    ' variables declared
    Dim answer As VbMsgBoxResult
    
    'initialise the dimensioned variables
    answer = vbNo
    

    ' On Error GoTo mnuSweets_Click_Error
       If debugflg = 1 Then Debug.Print "%" & "mnuSweets_Click"
    
    
    answer = MsgBox(" Help support the creation of more widgets like this. Buy me a small item on my Amazon wishlist! This button opens a browser window and connects to my Amazon wish list page). Will you be kind and proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "http://www.amazon.co.uk/gp/registry/registry.html?ie=UTF8&id=A3OBFB6ZN4F7&type=wishlist", vbNullString, App.Path, 1)
    End If
    
    On Error GoTo 0
    Exit Sub

mnuSweets_Click_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure mnuSweets_Click of Form quartermaster"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuClosePreferences_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuClosePreferences_Click()
   On Error GoTo mnuClosePreferences_Click_Error

    Call btnCancel_Click

   On Error GoTo 0
   Exit Sub

mnuClosePreferences_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuClosePreferences_Click of Form FireCallPrefs"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : mnuAuto_Click
' Author    : beededea
' Date      : 19/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub mnuAuto_Click()
    ' set the menu checks
    
   ' On Error GoTo mnuAuto_Click_Error

    If FireCallPrefs.themeTimer.Enabled = True Then
            MsgBox "Automatic Theme Selection is now Disabled"
            FireCallPrefs.mnuAuto.Caption = "Auto Theme Disabled - Click to Enable"
            FireCallPrefs.mnuAuto.Checked = False
            
            FireCallPrefs.themeTimer.Enabled = False
    Else
            MsgBox "Auto Theme Selection Enabled. If the o/s theme changes the utility should automatically skin the utility to suit the theme."
            FireCallPrefs.mnuAuto.Caption = "Auto Theme Enabled - Click to Disable"
            FireCallPrefs.mnuAuto.Checked = True
            
            FireCallPrefs.themeTimer.Enabled = True
            Call setThemeColour
    End If

   On Error GoTo 0
   Exit Sub

mnuAuto_Click_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure mnuAuto_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuDark_Click
' Author    : beededea
' Date      : 19/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub mnuDark_Click()
   ' On Error GoTo mnuDark_Click_Error

    FireCallPrefs.mnuAuto.Caption = "Auto Theme Disabled - Click to Enable"
    FireCallPrefs.mnuAuto.Checked = False
    FireCallPrefs.mnuDark.Caption = "Dark Theme Enabled"
    FireCallPrefs.mnuLight.Caption = "Light Theme Enable"
    FireCallPrefs.themeTimer.Enabled = False
    
    FCWSkinTheme = "dark"

    Call setThemeShade(212, 208, 199)

   On Error GoTo 0
   Exit Sub

mnuDark_Click_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure mnuDark_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuLight_Click
' Author    : beededea
' Date      : 19/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub mnuLight_Click()
    'MsgBox "Auto Theme Selection Manually Disabled"
   On Error GoTo mnuLight_Click_Error
    
    FireCallPrefs.mnuAuto.Caption = "Auto Theme Disabled - Click to Enable"
    FireCallPrefs.mnuAuto.Checked = False
    FireCallPrefs.mnuDark.Caption = "Dark Theme Enable"
    FireCallPrefs.mnuLight.Caption = "Light Theme Enabled"
    FireCallPrefs.themeTimer.Enabled = False
    FCWSkinTheme = "light"

    Call setThemeShade(240, 240, 240)

   On Error GoTo 0
   Exit Sub

mnuLight_Click_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure mnuLight_Click of Form dockSettings"
End Sub


' right click menu display
'---------------------------------------------------------------------------------------
' Procedure : Form_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo Form_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

Form_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure Form_MouseDown of Form FireCallPrefs"
End Sub

' right click menu display
'---------------------------------------------------------------------------------------
' Procedure : fraConfiguration_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraConfiguration_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo fraConfiguration_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

fraConfiguration_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraConfiguration_MouseDown of Form FireCallPrefs"
End Sub

' right click menu display
'---------------------------------------------------------------------------------------
' Procedure : fraEmail_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraEmail_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo fraEmail_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

fraEmail_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraEmail_MouseDown of Form FireCallPrefs"
End Sub


' right click menu display
'---------------------------------------------------------------------------------------
' Procedure : fraEmojis_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraEmojis_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo fraEmojis_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

fraEmojis_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraEmojis_MouseDown of Form FireCallPrefs"
End Sub


' right click menu display
'---------------------------------------------------------------------------------------
' Procedure : fraFonts_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraFonts_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo fraFonts_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

fraFonts_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraFonts_MouseDown of Form FireCallPrefs"
End Sub

' right click menu display
'---------------------------------------------------------------------------------------
' Procedure : fraGeneral_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fraGeneral_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo fraGeneral_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

fraGeneral_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fraGeneral_MouseDown of Form FireCallPrefs"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : changePrefsFont
' Author    : beededea
' Date      : 12/07/2019
' Purpose   : change the font throughout the whole form
'---------------------------------------------------------------------------------------
'
Public Sub changePrefsFont(ByRef formName As Object, ByVal suppliedFont As String, ByVal suppliedSize As Integer, ByVal suppliedWeight As Integer, ByVal suppliedStyle As Boolean)
        
    ' variables declared
    'Dim useloop As Integer
    Dim ctrl As Control
        
    'initialise the dimensioned variables
    'useloop = 0
    'Ctrl
    
    ' On Error GoTo changePrefsFont_Error
    
    If debugflg = 1 Then Debug.Print "%" & "changePrefsFont"
      
    ' a method of looping through all the controls and identifying the labels and text boxes
    For Each ctrl In formName.Controls
'      If formName.Name = "FireCallPrefs" And Ctrl = "txtTextFont" Then

         If (TypeOf ctrl Is CommandButton) Or (TypeOf ctrl Is TextBox) Or (TypeOf ctrl Is FileListBox) Or (TypeOf ctrl Is Label) Or (TypeOf ctrl Is ComboBox) Or (TypeOf ctrl Is CheckBox) Or (TypeOf ctrl Is OptionButton) Or (TypeOf ctrl Is Frame) Or (TypeOf ctrl Is ListBox) Then
           If suppliedFont <> vbNullString Then ctrl.Font.Name = suppliedFont
           If suppliedSize > 0 Then ctrl.Font.Size = suppliedSize
           'If suppliedStyle <> "" Then Ctrl.Font.Style = suppliedStyle
        End If

    Next
    
    FireCallPrefs.txtTextFont.Font.Name = FireCallPrefs.txtTextFont.Text
         


       
   On Error GoTo 0
   Exit Sub

changePrefsFont_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure changePrefsFont of Form dockSettings"
    
End Sub





    
'---------------------------------------------------------------------------------------
' Procedure : testEmailTestButton
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub testEmailTestButton()
    
   On Error GoTo testEmailTestButton_Error

    If FCWSmtpServer <> "" And _
        FCWSmtpUsername <> "" And _
        FCWSmtpPassword <> "" And _
        FCWRecipientEmail <> "" And _
        FCWEmailSubject <> "" And _
        FCWSmtpPort <> "" And _
        FCWSmtpConfigName <> "" And _
        FCWEmailMessage <> "" Then
        
        btnTestEmail.Enabled = True
    Else
        btnTestEmail.Enabled = False
    End If

    If chkSendEmails.Value = 0 And chkSendErrorEmails.Value = 0 Then
        Call toggleAllEmailControls("hide")
    End If

   On Error GoTo 0
   Exit Sub

testEmailTestButton_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure testEmailTestButton of Form FireCallPrefs"

End Sub






'---------------------------------------------------------------------------------------
' Procedure : checkIconiseOpacityLevel
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub checkIconiseOpacityLevel()

   On Error GoTo checkIconiseOpacityLevel_Error

    If optIconiseOpacity.Value = True Then
        lblOpacityLabel.Enabled = True
        sliOpacity.Enabled = True
        lblOpacity20.Enabled = True
        lblOpacityText.Enabled = True
        lblOpacityLabel100.Enabled = True
        lblOpacityLabelDesc.Enabled = True
        lblOptIconiseOpacity.Enabled = True
    Else
        lblOpacityLabel.Enabled = False
        sliOpacity.Enabled = False
        lblOpacity20.Enabled = False
        lblOpacityText.Enabled = False
        lblOpacityLabel100.Enabled = False
        lblOpacityLabelDesc.Enabled = False
        lblOptIconiseOpacity.Enabled = False
    End If

   On Error GoTo 0
   Exit Sub

checkIconiseOpacityLevel_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure checkIconiseOpacityLevel of Form FireCallPrefs"
End Sub

