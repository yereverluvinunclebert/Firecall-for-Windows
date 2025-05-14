VERSION 5.00
Begin VB.Form FireCallMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Fire Call Win"
   ClientHeight    =   10185
   ClientLeft      =   3120
   ClientTop       =   2070
   ClientWidth     =   10065
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":058A
   ScaleHeight     =   10185
   ScaleWidth      =   10065
   Begin VB.ComboBox cmbHiddenCaptureDevices 
      Height          =   315
      Left            =   60
      TabIndex        =   89
      Text            =   "cmbHiddenCaptureDevices"
      Top             =   1050
      Visible         =   0   'False
      Width           =   3225
   End
   Begin VB.ComboBox cmbEmojiSelection 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   465
      Style           =   2  'Dropdown List
      TabIndex        =   87
      ToolTipText     =   "Select from a list of JPG Emojis"
      Top             =   255
      Width           =   6015
   End
   Begin VB.ListBox lbxCombinedTextArea 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   1395
      Left            =   135
      MultiSelect     =   2  'Extended
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   36
      Top             =   675
      Visible         =   0   'False
      Width           =   7245
   End
   Begin VB.Frame hiddenFrame 
      Height          =   5355
      Left            =   90
      TabIndex        =   18
      Top             =   2565
      Visible         =   0   'False
      Width           =   7275
      Begin VB.Timer opacityToTimer 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   2625
         Top             =   3930
      End
      Begin VB.Timer configBusyTimer 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   5010
         Top             =   3630
      End
      Begin VB.Timer houseKeepingTimer 
         Enabled         =   0   'False
         Interval        =   65000
         Left            =   300
         Tag             =   "runs once a minute and promptly exits"
         Top             =   3885
      End
      Begin VB.Timer emailIconTimer 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   2625
         Top             =   3465
      End
      Begin VB.Timer emailTimer 
         Enabled         =   0   'False
         Left            =   300
         Top             =   3450
      End
      Begin VB.Timer backupTimer 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   5010
         Tag             =   "This is the 60 second timer for the backups "
         Top             =   3150
      End
      Begin VB.Timer PlayTimer 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   300
         Top             =   675
      End
      Begin VB.Timer recordTimer 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   300
         Top             =   210
      End
      Begin VB.Timer shutdownTimer 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   5010
         Top             =   2670
      End
      Begin VB.Timer combinedScrollBarTimer 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   5010
         Tag             =   "When the scrollbars are set to hidden, causes the vertical scrollbar to disappear 2 seconds after the last keypress"
         Top             =   2190
      End
      Begin VB.Timer sendCommandTimer 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   4995
         Top             =   765
      End
      Begin VB.Timer buzzerTimer 
         Enabled         =   0   'False
         Interval        =   1250
         Left            =   4995
         Top             =   270
      End
      Begin VB.Timer clockTimer 
         Interval        =   1000
         Left            =   2625
         Tag             =   "This is the timer for the analogue clock"
         Top             =   3015
      End
      Begin VB.Timer brightTimer 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   300
         Top             =   2985
      End
      Begin VB.Timer pausePrinterTimer 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2625
         Top             =   2550
      End
      Begin VB.Timer dropTimer 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   300
         Top             =   2520
      End
      Begin VB.Timer shredderTimer 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2625
         Top             =   2085
      End
      Begin VB.Timer printerTimer 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   300
         Top             =   2055
      End
      Begin VB.Timer pollingTimer 
         Enabled         =   0   'False
         Left            =   2625
         Top             =   1620
      End
      Begin VB.Timer zOrderTimer 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   2625
         Top             =   1155
      End
      Begin VB.Timer opacityFadeInTimer 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   2625
         Top             =   690
      End
      Begin VB.Timer opacityFadeOutTimer 
         Enabled         =   0   'False
         Interval        =   20
         Left            =   2625
         Top             =   225
      End
      Begin VB.Timer iconiseTimer 
         Enabled         =   0   'False
         Left            =   300
         Tag             =   "iconise the main form to the stamp icon"
         Top             =   1605
      End
      Begin VB.Timer lampTimer 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   300
         Tag             =   "turns a lit lamp off"
         Top             =   1140
      End
      Begin VB.Timer outputScrollBarTimer 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   5010
         Tag             =   "When the scrollbars are set to hidden, causes the vertical scrollbar to disappear 2 seconds after the last keypress"
         Top             =   1725
      End
      Begin VB.Timer inputScrollBarTimer 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   5010
         Tag             =   "When the scrollbars are set to hidden, causes the vertical scrollbar to disappear 2 seconds after the last keypress"
         Top             =   1245
      End
      Begin VB.Label lblopacityFadeOutTimer 
         Caption         =   "opacityFadeToTimer"
         Height          =   285
         Index           =   1
         Left            =   3210
         TabIndex        =   96
         Tag             =   "When the scrollbars are set to hidden, causes the vertical scrollbar to disappear 2 seconds after the last keypress"
         Top             =   4005
         Width           =   1785
      End
      Begin VB.Label lblTimer 
         Caption         =   "configBusyTimer"
         Height          =   285
         Index           =   5
         Left            =   5520
         TabIndex        =   95
         Tag             =   "When the scrollbars are set to hidden, causes the vertical scrollbar to disappear 2 seconds after the last keypress"
         Top             =   3675
         Width           =   1710
      End
      Begin VB.Label lblTimer 
         Caption         =   "HouseKeepingTimer"
         Height          =   285
         Index           =   4
         Left            =   900
         TabIndex        =   94
         Tag             =   "turns a lit lamp off"
         Top             =   3990
         Width           =   1500
      End
      Begin VB.Label lblTimer 
         Caption         =   "emailIconTimer"
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   93
         Tag             =   "turns a lit lamp off"
         Top             =   3570
         Width           =   1140
      End
      Begin VB.Label lblTimer 
         Caption         =   "EmailTimer"
         Height          =   285
         Index           =   0
         Left            =   885
         TabIndex        =   90
         Tag             =   "turns a lit lamp off"
         Top             =   3540
         Width           =   1140
      End
      Begin VB.Label lblTimer 
         Caption         =   "backupTimer"
         Height          =   285
         Index           =   2
         Left            =   5520
         TabIndex        =   86
         Tag             =   "When the scrollbars are set to hidden, causes the vertical scrollbar to disappear 2 seconds after the last keypress"
         Top             =   3240
         Width           =   1710
      End
      Begin VB.Label lblPlayTimer 
         Caption         =   "playTimer"
         Height          =   285
         Left            =   855
         TabIndex        =   85
         Tag             =   "When the scrollbars are set to hidden, causes the vertical scrollbar to disappear 2 seconds after the last keypress"
         Top             =   765
         Width           =   1785
      End
      Begin VB.Label lblEmailTimer 
         Caption         =   "recordTimer"
         Height          =   285
         Index           =   2
         Left            =   855
         TabIndex        =   81
         Tag             =   "When the scrollbars are set to hidden, causes the vertical scrollbar to disappear 2 seconds after the last keypress"
         Top             =   330
         Width           =   1785
      End
      Begin VB.Label lblShutdownTimer 
         Caption         =   "shutdownTimer"
         Height          =   285
         Left            =   5520
         TabIndex        =   80
         Tag             =   "When the scrollbars are set to hidden, causes the vertical scrollbar to disappear 2 seconds after the last keypress"
         Top             =   2745
         Width           =   1785
      End
      Begin VB.Label lblcombinedScrollBarTimer 
         Caption         =   "combinedScrollBarTimer"
         Height          =   285
         Left            =   5520
         TabIndex        =   37
         Tag             =   "When the scrollbars are set to hidden, causes the vertical scrollbar to disappear 2 seconds after the last keypress"
         Top             =   2280
         Width           =   1785
      End
      Begin VB.Label lblsendCommandTimer 
         Caption         =   "sendCommandTimer"
         Height          =   285
         Left            =   5475
         TabIndex        =   35
         Tag             =   "When the scrollbars are set to hidden, causes the vertical scrollbar to disappear 2 seconds after the last keypress"
         Top             =   855
         Width           =   1455
      End
      Begin VB.Label lblbuzzerTimer 
         Caption         =   "buzzerTimer"
         Height          =   285
         Left            =   5475
         TabIndex        =   34
         Tag             =   "When the scrollbars are set to hidden, causes the vertical scrollbar to disappear 2 seconds after the last keypress"
         Top             =   360
         Width           =   1290
      End
      Begin VB.Label lblclockTimer 
         Caption         =   "clockTimer"
         Height          =   285
         Left            =   3240
         TabIndex        =   33
         Tag             =   "When the scrollbars are set to hidden, causes the vertical scrollbar to disappear 2 seconds after the last keypress"
         Top             =   3105
         Width           =   1710
      End
      Begin VB.Label lblpausePrinterTimer 
         Caption         =   "pausePrinterTimer"
         Height          =   285
         Left            =   3240
         TabIndex        =   32
         Tag             =   "When the scrollbars are set to hidden, causes the vertical scrollbar to disappear 2 seconds after the last keypress"
         Top             =   2625
         Width           =   1710
      End
      Begin VB.Label lblshredderTimer 
         Caption         =   "shredderTimer"
         Height          =   285
         Left            =   3255
         TabIndex        =   31
         Tag             =   "When the scrollbars are set to hidden, causes the vertical scrollbar to disappear 2 seconds after the last keypress"
         Top             =   2175
         Width           =   1260
      End
      Begin VB.Label lblTimer 
         Caption         =   "brightTimer"
         Height          =   285
         Index           =   3
         Left            =   885
         TabIndex        =   30
         Tag             =   "turns a lit lamp off"
         Top             =   3090
         Width           =   1140
      End
      Begin VB.Label lbldropTimer 
         Caption         =   "dropTimer"
         Height          =   285
         Left            =   885
         TabIndex        =   29
         Tag             =   "turns a lit lamp off"
         Top             =   2595
         Width           =   1140
      End
      Begin VB.Label lblPrinterTimer 
         Caption         =   "printerTimer"
         Height          =   285
         Left            =   870
         TabIndex        =   28
         Tag             =   "turns a lit lamp off"
         Top             =   2130
         Width           =   1140
      End
      Begin VB.Label lblPollingTimer 
         Caption         =   "pollingTimer *"
         Height          =   285
         Left            =   3255
         TabIndex        =   27
         Tag             =   "When the scrollbars are set to hidden, causes the vertical scrollbar to disappear 2 seconds after the last keypress"
         Top             =   1755
         Width           =   1260
      End
      Begin VB.Label lblTimerDesc 
         Caption         =   $"Form1.frx":72854
         Height          =   585
         Left            =   345
         TabIndex        =   26
         Top             =   4650
         Width           =   6330
      End
      Begin VB.Label lblzOrderTimer 
         Caption         =   "zOrderTimer"
         Height          =   285
         Left            =   3255
         TabIndex        =   25
         Tag             =   "When the scrollbars are set to hidden, causes the vertical scrollbar to disappear 2 seconds after the last keypress"
         Top             =   1260
         Width           =   1260
      End
      Begin VB.Label lblopacityFadeInTimer 
         Caption         =   "opacityFadeInTimer"
         Height          =   285
         Left            =   3240
         TabIndex        =   24
         Tag             =   "When the scrollbars are set to hidden, causes the vertical scrollbar to disappear 2 seconds after the last keypress"
         Top             =   810
         Width           =   1785
      End
      Begin VB.Label lblopacityFadeOutTimer 
         Caption         =   "opacityFadeOutTimer"
         Height          =   285
         Index           =   0
         Left            =   3195
         TabIndex        =   23
         Tag             =   "When the scrollbars are set to hidden, causes the vertical scrollbar to disappear 2 seconds after the last keypress"
         Top             =   315
         Width           =   1785
      End
      Begin VB.Label lbliconiseTimer 
         Caption         =   "iconiseTimer *"
         Height          =   285
         Left            =   885
         TabIndex        =   22
         Tag             =   "iconises the main application to a small stamp image"
         Top             =   1710
         Width           =   1140
      End
      Begin VB.Label lblinputScrollBarTimer 
         Caption         =   "inputScrollBarTimer"
         Height          =   285
         Left            =   5535
         TabIndex        =   21
         Tag             =   "When the scrollbars are set to hidden, causes the vertical scrollbar to disappear 2 seconds after the last keypress"
         Top             =   1365
         Width           =   1785
      End
      Begin VB.Label lbloutputScrollBarTimer 
         Caption         =   "outputScrollBarTimer"
         Height          =   285
         Left            =   5520
         TabIndex        =   20
         Tag             =   "When the scrollbars are set to hidden, causes the vertical scrollbar to disappear 2 seconds after the last keypress"
         Top             =   1815
         Width           =   1785
      End
      Begin VB.Label lbllampTimer 
         Caption         =   "lampTimer"
         Height          =   285
         Left            =   885
         TabIndex        =   19
         Tag             =   "turns a lit lamp off"
         Top             =   1260
         Width           =   1140
      End
   End
   Begin VB.PictureBox picEmojiSmall 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   17
      Top             =   240
      Width           =   345
   End
   Begin VB.PictureBox btnPicOut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   6780
      Picture         =   "Form1.frx":728FD
      ScaleHeight     =   360
      ScaleWidth      =   630
      TabIndex        =   16
      ToolTipText     =   "Send - Just going out for a while, back later!"
      Top             =   9660
      Width           =   630
   End
   Begin VB.PictureBox btnPicCod 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   6135
      Picture         =   "Form1.frx":72F83
      ScaleHeight     =   360
      ScaleWidth      =   630
      TabIndex        =   15
      ToolTipText     =   "Send - busy coding here, and you?"
      Top             =   9660
      Width           =   630
   End
   Begin VB.PictureBox btnPicBusy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   5490
      Picture         =   "Form1.frx":73BC5
      ScaleHeight     =   360
      ScaleWidth      =   630
      TabIndex        =   14
      ToolTipText     =   "Send - Very busy at the moment."
      Top             =   9660
      Width           =   630
   End
   Begin VB.PictureBox btnPicGdn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   4845
      Picture         =   "Form1.frx":74288
      ScaleHeight     =   360
      ScaleWidth      =   630
      TabIndex        =   13
      ToolTipText     =   "Send - Out in the garden, doing stuff."
      Top             =   9660
      Width           =   630
   End
   Begin VB.PictureBox btnPicPrg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   4185
      Picture         =   "Form1.frx":74945
      ScaleHeight     =   360
      ScaleWidth      =   630
      TabIndex        =   12
      ToolTipText     =   "Send -  Doing a bit of programming today..."
      Top             =   9660
      Width           =   630
   End
   Begin VB.PictureBox BtnPicWth 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3525
      Picture         =   "Form1.frx":74FF2
      ScaleHeight     =   360
      ScaleWidth      =   630
      TabIndex        =   11
      ToolTipText     =   "Send - What's the weather like today?"
      Top             =   9660
      Width           =   630
   End
   Begin VB.PictureBox btnPicWot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   2865
      Picture         =   "Form1.frx":756B5
      ScaleHeight     =   360
      ScaleWidth      =   630
      TabIndex        =   10
      ToolTipText     =   "Send - What's happening?"
      Top             =   9660
      Width           =   630
   End
   Begin VB.PictureBox btnPicMorn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   2205
      Picture         =   "Form1.frx":75D80
      ScaleHeight     =   360
      ScaleWidth      =   630
      TabIndex        =   9
      ToolTipText     =   "Send Good Morning!"
      Top             =   9660
      Width           =   630
   End
   Begin VB.PictureBox btnPicNews 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1545
      Picture         =   "Form1.frx":76458
      ScaleHeight     =   360
      ScaleWidth      =   630
      TabIndex        =   8
      ToolTipText     =   "Send - What news?"
      Top             =   9660
      Width           =   630
   End
   Begin VB.PictureBox btnPicWell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   885
      Picture         =   "Form1.frx":76B37
      ScaleHeight     =   360
      ScaleWidth      =   630
      TabIndex        =   7
      ToolTipText     =   "Send - Are you well?"
      Top             =   9660
      Width           =   630
   End
   Begin VB.PictureBox btnPicTtfn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   225
      Picture         =   "Form1.frx":77204
      ScaleHeight     =   360
      ScaleWidth      =   630
      TabIndex        =   6
      ToolTipText     =   "Send - TTFN!"
      Top             =   9660
      Width           =   630
   End
   Begin VB.TextBox txtHiddenRetFileName 
      Height          =   360
      Left            =   3015
      TabIndex        =   5
      Text            =   "hidden"
      Top             =   5370
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton btnSendText 
      Height          =   375
      Left            =   6615
      Picture         =   "Form1.frx":778A8
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Click here to send your text"
      Top             =   9180
      Width           =   795
   End
   Begin VB.ListBox lbxOutputTextArea 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   4125
      Left            =   120
      MultiSelect     =   2  'Extended
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Top             =   4905
      Width           =   7245
   End
   Begin VB.ListBox lbxInputTextArea 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   4125
      Left            =   135
      MultiSelect     =   2  'Extended
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   675
      Width           =   7245
   End
   Begin VB.CommandButton btnEmojiSet 
      Height          =   375
      Left            =   6645
      Picture         =   "Form1.frx":77CF8
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "When you have chosen an Emoji then click here to send."
      Top             =   240
      Width           =   795
   End
   Begin VB.TextBox txtTextEntry 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   9180
      Width           =   6360
   End
   Begin VB.PictureBox picSideBar 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   10560
      Left            =   7410
      Picture         =   "Form1.frx":78148
      ScaleHeight     =   10560
      ScaleWidth      =   2715
      TabIndex        =   38
      Top             =   0
      Width           =   2715
      Begin VB.PictureBox picTextChangeBright 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1980
         Picture         =   "Form1.frx":7CF71
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   47
         ToolTipText     =   "This lamp will glow when there has been a recent update"
         Top             =   150
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox picTimerLampBright 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1605
         Picture         =   "Form1.frx":7D212
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   45
         ToolTipText     =   "When this lamp glows it is polling!"
         Top             =   150
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox picWEmailIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   60
         Picture         =   "Form1.frx":7D4B3
         ScaleHeight     =   375
         ScaleWidth      =   450
         TabIndex        =   92
         Top             =   2055
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.PictureBox picLidOpen 
         BorderStyle     =   0  'None
         Height          =   450
         Left            =   -15
         Picture         =   "Form1.frx":7DAF3
         ScaleHeight     =   450
         ScaleWidth      =   2655
         TabIndex        =   91
         Top             =   5355
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.PictureBox picThermometer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   120
         Picture         =   "Form1.frx":7F276
         ScaleHeight     =   405
         ScaleWidth      =   2460
         TabIndex        =   88
         Top             =   5055
         Width           =   2460
         Begin VB.Line linRed 
            BorderColor     =   &H000000C0&
            X1              =   540
            X2              =   1500
            Y1              =   210
            Y2              =   210
         End
      End
      Begin VB.PictureBox btnLid 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3480
         Left            =   -120
         Picture         =   "Form1.frx":80263
         ScaleHeight     =   3480
         ScaleWidth      =   2400
         TabIndex        =   74
         Top             =   6030
         Width           =   2400
         Begin VB.PictureBox picRecordLampBright 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   180
            Picture         =   "Form1.frx":87EC7
            ScaleHeight     =   360
            ScaleWidth      =   360
            TabIndex        =   78
            ToolTipText     =   "Speech is being recorded now..."
            Top             =   2085
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.PictureBox picBtnPlaySound 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   150
            Picture         =   "Form1.frx":88388
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   82
            Top             =   2535
            Width           =   480
         End
         Begin VB.PictureBox btnStop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   1815
            Picture         =   "Form1.frx":8894A
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   77
            ToolTipText     =   "End Recording"
            Top             =   345
            Width           =   495
         End
         Begin VB.PictureBox btnStartRecord 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   465
            Left            =   150
            Picture         =   "Form1.frx":88C19
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   76
            ToolTipText     =   "Record Speech"
            Top             =   345
            Width           =   465
         End
         Begin VB.PictureBox picBtnLidCatch 
            BorderStyle     =   0  'None
            Height          =   795
            Left            =   900
            Picture         =   "Form1.frx":891F7
            ScaleHeight     =   795
            ScaleWidth      =   585
            TabIndex        =   75
            ToolTipText     =   "This lid covers the emoji display"
            Top             =   2550
            Width           =   585
         End
         Begin VB.PictureBox picRecordLampDull 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   180
            Picture         =   "Form1.frx":896DA
            ScaleHeight     =   360
            ScaleWidth      =   360
            TabIndex        =   79
            ToolTipText     =   "This lamp glows red when recording"
            Top             =   2085
            Width           =   360
         End
         Begin VB.PictureBox picPlayLampDull 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   1860
            Picture         =   "Form1.frx":89B4D
            ScaleHeight     =   360
            ScaleWidth      =   360
            TabIndex        =   84
            Top             =   2070
            Width           =   360
         End
         Begin VB.PictureBox picPlayLampBright 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   1875
            Picture         =   "Form1.frx":8A046
            ScaleHeight     =   360
            ScaleWidth      =   360
            TabIndex        =   83
            ToolTipText     =   "Speech is being recorded now..."
            Top             =   2085
            Visible         =   0   'False
            Width           =   360
         End
      End
      Begin VB.PictureBox picFsoLid 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   2280
         Picture         =   "Form1.frx":8A553
         ScaleHeight     =   645
         ScaleWidth      =   345
         TabIndex        =   73
         ToolTipText     =   "Click this cover to reveal the FSO/UTF8 lamps"
         Top             =   1800
         Width           =   345
      End
      Begin VB.PictureBox picBtnLidShadow 
         BorderStyle     =   0  'None
         Height          =   3435
         Left            =   -1995
         Picture         =   "Form1.frx":8A770
         ScaleHeight     =   3435
         ScaleWidth      =   2385
         TabIndex        =   72
         Top             =   6615
         Width           =   2385
      End
      Begin VB.PictureBox picClockSwitch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   150
         Picture         =   "Form1.frx":8B28A
         ScaleHeight     =   345
         ScaleWidth      =   330
         TabIndex        =   71
         Top             =   2880
         Width           =   330
      End
      Begin VB.PictureBox picBuzzerDullLamp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1140
         Picture         =   "Form1.frx":8B80B
         ScaleHeight     =   330
         ScaleWidth      =   345
         TabIndex        =   70
         Top             =   2535
         Width           =   345
      End
      Begin VB.PictureBox picBuzzerBrightLamp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1140
         Picture         =   "Form1.frx":8BBDE
         ScaleHeight     =   330
         ScaleWidth      =   345
         TabIndex        =   69
         Top             =   2535
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton btnClose 
         Height          =   450
         Left            =   1080
         Picture         =   "Form1.frx":8C017
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Click to close FireCall"
         Top             =   9630
         Width           =   1350
      End
      Begin VB.CommandButton btnPing 
         Height          =   450
         Left            =   1080
         Picture         =   "Form1.frx":8C740
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Click to send a ping"
         Top             =   9135
         Width           =   1350
      End
      Begin VB.PictureBox btnPicHelp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   540
         Picture         =   "Form1.frx":8CD32
         ScaleHeight     =   570
         ScaleWidth      =   1725
         TabIndex        =   43
         ToolTipText     =   "Click to open the help for this utility"
         Top             =   1845
         Width           =   1725
      End
      Begin VB.PictureBox btnPicConfig 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   525
         Picture         =   "Form1.frx":8E42E
         ScaleHeight     =   585
         ScaleWidth      =   1725
         TabIndex        =   42
         ToolTipText     =   "Click to open the configuration/preferences for this program"
         Top             =   1200
         Width           =   1725
      End
      Begin VB.PictureBox btnPicAttach 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   540
         Picture         =   "Form1.frx":8FCB2
         ScaleHeight     =   570
         ScaleWidth      =   1725
         TabIndex        =   41
         ToolTipText     =   "Click to attach a single file for transmission"
         Top             =   600
         Width           =   1725
      End
      Begin VB.PictureBox picLidBackground 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3435
         Left            =   135
         Picture         =   "Form1.frx":914AB
         ScaleHeight     =   3435
         ScaleWidth      =   2415
         TabIndex        =   56
         Top             =   5625
         Width           =   2415
         Begin VB.PictureBox picImagePrintOut 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0E0FF&
            ForeColor       =   &H80000008&
            Height          =   2475
            Left            =   195
            ScaleHeight     =   163
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   142
            TabIndex        =   57
            Top             =   2250
            Width           =   2160
            Begin VB.PictureBox picCloseMe 
               BackColor       =   &H00C0E0FF&
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   1875
               Picture         =   "Form1.frx":92F25
               ScaleHeight     =   210
               ScaleWidth      =   225
               TabIndex        =   58
               ToolTipText     =   "Click to close the image"
               Top             =   45
               Width           =   225
            End
         End
         Begin VB.PictureBox picEmoji 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1455
            Left            =   495
            Picture         =   "Form1.frx":93152
            ScaleHeight     =   1455
            ScaleWidth      =   1500
            TabIndex        =   59
            ToolTipText     =   "Click on me to show partner's Emoji status"
            Top             =   -1200
            Width           =   1500
         End
         Begin VB.PictureBox picPrintOutShadow 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2475
            Left            =   1530
            ScaleHeight     =   165
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   144
            TabIndex        =   60
            Top             =   2115
            Visible         =   0   'False
            Width           =   2160
         End
         Begin VB.PictureBox picImageButton 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   465
            Left            =   1965
            Picture         =   "Form1.frx":951A1
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   66
            Top             =   330
            Width           =   465
         End
         Begin VB.PictureBox picSpeakerGrille 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   690
            Left            =   765
            Picture         =   "Form1.frx":95426
            ScaleHeight     =   690
            ScaleWidth      =   930
            TabIndex        =   64
            Top             =   1935
            Width           =   930
         End
         Begin VB.PictureBox picOutputEmoji 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1440
            Left            =   510
            Picture         =   "Form1.frx":9592E
            ScaleHeight     =   1440
            ScaleWidth      =   1455
            TabIndex        =   63
            Top             =   375
            Width           =   1455
         End
         Begin VB.PictureBox picEmojiKnobRight 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   720
            Left            =   1680
            Picture         =   "Form1.frx":96644
            ScaleHeight     =   720
            ScaleWidth      =   720
            TabIndex        =   62
            Top             =   1935
            Width           =   720
         End
         Begin VB.PictureBox picEmojiKnobLeft 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   720
            Left            =   0
            Picture         =   "Form1.frx":96A62
            ScaleHeight     =   720
            ScaleWidth      =   675
            TabIndex        =   61
            Top             =   1950
            Width           =   675
         End
         Begin VB.PictureBox picSpeakerGrilleOpen 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   690
            Left            =   765
            Picture         =   "Form1.frx":96E5C
            ScaleHeight     =   690
            ScaleWidth      =   930
            TabIndex        =   65
            Top             =   1935
            Width           =   930
         End
      End
      Begin VB.PictureBox picUtf8LampDull 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   2310
         Picture         =   "Form1.frx":972ED
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   51
         ToolTipText     =   "This lamp will glow when writing files as UTF8"
         Top             =   2130
         Width           =   315
      End
      Begin VB.PictureBox picFsoLampDull 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   2310
         Picture         =   "Form1.frx":9752F
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   49
         ToolTipText     =   "This lamp will glow when we are writing files as ANSI using FSO"
         Top             =   1815
         Width           =   315
      End
      Begin VB.PictureBox picFsoLampBright 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   2310
         Picture         =   "Form1.frx":97766
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   48
         ToolTipText     =   "We are currently writing files as ANSI using FSO"
         Top             =   1815
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox picUtf8LampBright 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   2310
         Picture         =   "Form1.frx":979BF
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   50
         ToolTipText     =   "We are currently writing files as UTF8"
         Top             =   2130
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox picClock 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2955
         Left            =   120
         Picture         =   "Form1.frx":97C1F
         ScaleHeight     =   2955
         ScaleWidth      =   2460
         TabIndex        =   53
         Top             =   2535
         Width           =   2460
         Begin VB.Shape shpCentreBoss 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   180
            Left            =   1125
            Shape           =   3  'Circle
            Top             =   1320
            Width           =   135
         End
         Begin VB.Line MinuteHand 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   3
            Index           =   1
            Visible         =   0   'False
            X1              =   945
            X2              =   735
            Y1              =   795
            Y2              =   1395
         End
         Begin VB.Line MinuteHand 
            BorderColor     =   &H00404040&
            BorderWidth     =   5
            Index           =   0
            Visible         =   0   'False
            X1              =   615
            X2              =   870
            Y1              =   990
            Y2              =   1320
         End
         Begin VB.Line HourHand 
            BorderColor     =   &H00000040&
            BorderWidth     =   5
            Index           =   1
            Visible         =   0   'False
            X1              =   1485
            X2              =   1680
            Y1              =   1440
            Y2              =   900
         End
         Begin VB.Line HourHand 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   6
            Index           =   0
            Visible         =   0   'False
            X1              =   1410
            X2              =   1605
            Y1              =   1425
            Y2              =   840
         End
         Begin VB.Line SecondHand 
            BorderColor     =   &H00C0E0FF&
            Index           =   1
            Visible         =   0   'False
            X1              =   1185
            X2              =   1185
            Y1              =   1425
            Y2              =   450
         End
         Begin VB.Line SecondHand 
            BorderColor     =   &H00000000&
            BorderWidth     =   3
            Index           =   0
            X1              =   1185
            X2              =   1170
            Y1              =   1350
            Y2              =   525
         End
         Begin VB.Line SecondHandStub 
            BorderColor     =   &H00C0E0FF&
            Index           =   1
            Visible         =   0   'False
            X1              =   1185
            X2              =   1185
            Y1              =   1650
            Y2              =   1395
         End
         Begin VB.Line SecondHandStub 
            BorderColor     =   &H00000000&
            BorderWidth     =   3
            Index           =   0
            X1              =   1185
            X2              =   1185
            Y1              =   1650
            Y2              =   1395
         End
         Begin VB.Label lblDate 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   6
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   180
            Left            =   1515
            TabIndex        =   55
            Top             =   1635
            Width           =   195
         End
         Begin VB.Label lblSeconds 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   6
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   180
            Left            =   780
            TabIndex        =   54
            Top             =   870
            Width           =   195
         End
      End
      Begin VB.PictureBox picRedButton 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2955
         Left            =   120
         Picture         =   "Form1.frx":9FE48
         ScaleHeight     =   2955
         ScaleWidth      =   2460
         TabIndex        =   52
         Top             =   2535
         Width           =   2460
      End
      Begin VB.PictureBox picTextChangeDull 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1980
         Picture         =   "Form1.frx":A2D4D
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   46
         ToolTipText     =   "This lamp will glow when there has been a recent update"
         Top             =   135
         Width           =   315
      End
      Begin VB.PictureBox picTimerLampDull 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1605
         Picture         =   "Form1.frx":A2FC1
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   44
         ToolTipText     =   "This lamp will glow when the program is polling for new data"
         Top             =   135
         Width           =   315
      End
      Begin VB.CommandButton btnMinimise 
         Height          =   405
         Left            =   555
         Picture         =   "Form1.frx":A3235
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Minimise the program to a desktop icon"
         Top             =   105
         Width           =   390
      End
      Begin VB.CommandButton btnCloseIt 
         Height          =   405
         Left            =   990
         Picture         =   "Form1.frx":A3757
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Close the program"
         Top             =   105
         Width           =   390
      End
   End
   Begin VB.Menu mainMnuPopmenu 
      Caption         =   "The main menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAboutFireCallWin 
         Caption         =   "About Fire Call Win"
      End
      Begin VB.Menu mnuBlankLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProgramPreferences 
         Caption         =   "Program Preferences"
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
      Begin VB.Menu mnuBlankLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTogglePolling 
         Caption         =   "Disable Polling Temporarily"
      End
      Begin VB.Menu mnuBlankLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSendPingRequest 
         Caption         =   "Send a Ping Request"
      End
      Begin VB.Menu mnuSendAwakeCall 
         Caption         =   "Send an Awake Call"
      End
      Begin VB.Menu mnuSendShutdownRequest 
         Caption         =   "Send a Shutdown Request"
      End
      Begin VB.Menu mnuBlankLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowEmojiState 
         Caption         =   "Show/Hide the Emoji State"
      End
      Begin VB.Menu mnuShowClock 
         Caption         =   "Show/Hide the Clock"
      End
      Begin VB.Menu mnuBlankLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Font selection for this utility"
         Visible         =   0   'False
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
      Begin VB.Menu mnuBlankLine6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBringToCentre 
         Caption         =   "Centre Program on Main Monitor"
      End
      Begin VB.Menu mnuLicenceA 
         Caption         =   "Display Licence Agreement"
      End
      Begin VB.Menu mnuBlankLine7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh the Chat boxes (F5)"
      End
      Begin VB.Menu mnuHideProgram 
         Caption         =   "Hide Program"
      End
      Begin VB.Menu mnuCloseProgram 
         Caption         =   "Close Program"
      End
   End
   Begin VB.Menu listBoxMnuPopmenu 
      Caption         =   "List Box Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuLBoxSendPingRequest 
         Caption         =   "Send a Ping Request"
      End
      Begin VB.Menu mnuLBoxSendAwakeCall 
         Caption         =   "Send an Awake Call"
      End
      Begin VB.Menu mnuLBoxSendShutdownRequest 
         Caption         =   "Send Shutdown Request"
      End
      Begin VB.Menu mnuBlankLine8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCombinedEditLine 
         Caption         =   "Edit This Line"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOutputEditLine 
         Caption         =   "Edit This Line"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCombinedDeleteLine 
         Caption         =   "Delete This Line"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOutputDeleteLine 
         Caption         =   "Delete This Line"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOutputCopyLine 
         Caption         =   "Copy Selected Line(s) to Clipboard (Ctrl+C)"
      End
      Begin VB.Menu mnuInputCopyLine 
         Caption         =   "Copy Selected Line(s) to Clipboard (Ctrl+C)"
      End
      Begin VB.Menu mnuInputQuoteLine 
         Caption         =   "Copy and Quote Line"
      End
      Begin VB.Menu mnuCombinedQuoteLine 
         Caption         =   "Copy and Quote Line"
      End
      Begin VB.Menu mnuCombinedCopyLine 
         Caption         =   "Copy Selected Line(s) to Clipboard (Ctrl+C)"
      End
      Begin VB.Menu mnuCombinedPasteLine 
         Caption         =   "Paste From Clipboard (Ctrl+V)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOutputPasteLine 
         Caption         =   "Paste From Clipboard (Ctrl+V)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCombinedPasteAndGo 
         Caption         =   "Paste && Go"
      End
      Begin VB.Menu mnuOutputPasteAndGo 
         Caption         =   "Paste && Go"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBlankLine9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSwitchChatBoxes 
         Caption         =   "Switch to Single Chat Box"
      End
      Begin VB.Menu mnuBlankLine10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindInput 
         Caption         =   "Find (Ctrl+F)"
      End
      Begin VB.Menu mnuFindOutput 
         Caption         =   "Find (Ctrl+F)"
      End
      Begin VB.Menu mnuFindCombined 
         Caption         =   "Find (Ctrl+F)"
      End
      Begin VB.Menu mnuLBOpenSharedInputFile 
         Caption         =   "Open the Shared Input File"
      End
      Begin VB.Menu mnuLBOpenSharedOutputFile 
         Caption         =   "Open the Shared Output File"
      End
      Begin VB.Menu mnuLBOpenSharedExchangeFolder 
         Caption         =   "Open the Shared Exchange Folder"
      End
      Begin VB.Menu mnuLBRefresh 
         Caption         =   "Refresh the Chat boxes (F5)"
      End
   End
   Begin VB.Menu ClockMnuPopmenu 
      Caption         =   "Clock Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuSynchWindowsTime 
         Caption         =   "Synchronise Windows Time"
      End
      Begin VB.Menu mnuHandsCode 
         Caption         =   "Select hands drawn using just lines and code"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHandsGdip 
         Caption         =   "Select hands animated using GDI+"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu picMnuPopmenu 
      Caption         =   "Picture Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuFindFile 
         Caption         =   "Open Folder for this Attachment"
      End
      Begin VB.Menu mnuOpenFile 
         Caption         =   "Open this file using Default App."
      End
   End
   Begin VB.Menu textMnuPopmenu 
      Caption         =   "Text Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuText1 
         Caption         =   "Text 1"
      End
      Begin VB.Menu mnuText2 
         Caption         =   "Text 2"
      End
      Begin VB.Menu mnuText3 
         Caption         =   "Text 3"
      End
      Begin VB.Menu mnuText4 
         Caption         =   "Text 5"
      End
      Begin VB.Menu mnuText5 
         Caption         =   "Text 5"
      End
      Begin VB.Menu mnuText6 
         Caption         =   "Text 6"
      End
      Begin VB.Menu mnuText7 
         Caption         =   "Text 7"
      End
      Begin VB.Menu mnuText8 
         Caption         =   "Text 8"
      End
      Begin VB.Menu mnuText9 
         Caption         =   "Text 9"
      End
      Begin VB.Menu mnuText10 
         Caption         =   "Text 10"
      End
   End
End
Attribute VB_Name = "FireCallMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Node: The buglist is buglist.txt in the project's Related Documents.
'---------------------------------------------------------------------------------------
' Module    : FireCallMain
' Author    : beededea
' Date      : 17/08/2021
' Purpose   :
'---------------------------------------------------------------------------------------

'@PredeclaredId
'@ModuleAttribute VB_Creatable, False
'@ModuleAttribute VB_Exposed, False
'@ModuleAttribute VB_GlobalNameSpace, False
'@ModuleAttribute VB_Name, "FireCallMain"
'@IgnoreModule ProcedureNotUsed

' For those that don't know the above are Rubberduck annotations that assist when RD is doing its code quality
' analysis.

'---------------------------------------------------------------------------------
' Thanks:   LA Volpe (VB Forums) for his transparent picture handling.
'           Shuja Ali (codeguru.com) for his settings.ini code.
'           Registry reading code from ALLAPI.COM.
'           Rxbagain on codeguru for his Open File common dialog code without dependent OCX.
'           Krool on the VBForums for his impressive common control replacements, slider and textboxW.
'           si_the_geek for his special folder code.
'           theTrick for his sound recording and saving to a WAV file.
'           Elroy for his kind help with subclassing and balloon tooltips and all his other kindness.
'           Wqweto for his innovative email injection work and help.
'
'           That's all as far as I know. There may be others but it is not my intention to hide their efforts.
'
' Built using: VB6, MZ-TOOLS 3.0, CodeHelp Core IDE Extender Framework 2.2 & Rubberduck 2.4.1
'
' Credits : MZ-TOOLS https://www.mztools.com/
'           vBAdvance
'           CodeHelp http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=62468&lngWId=1
'           Rubberduck http://rubberduckvba.com/
'           Registry code ALLAPI.COM
'           La Volpe superb VB6 non-native image types  http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1
'           Open File common dialog code without dependent OCX - http://forums.codeguru.com/member.php?92278-rxbagain
'           Open font dialog code without dependent OCX - unknown URL
'           Krool's superb replacement Controls http://www.vbforums.com/showthread.php?698563-CommonControls-%28Replacement-of-the-MS-common-controls%29
'           Chris Fannin (AbbydonKrafts) Copying a folder  http://vbcity.com/forums/t/129391.aspx
'           Austin Hickl fnGetDateInUniversalFormat  http://computer-programming-forum.com/66-vb-controls/6dff1bae05df0a6e.htm
'           Ellis Dee VB6 quicksort https://www.vbforums.com/showthread.php?473677-VB6-Sorting-algorithms-(sort-array-sorting-arrays)
'           KayJay fnIsGoodURL that utilises the isValidURL API  https://www.vbforums.com/showthread.php?231061-Validate-URL&p=1361958&viewfull=1#post1361958
'           JCI's resize image https://www.vbforums.com/member.php?40893-jcis
'           qvb6 vb6 date from epoch  https://www.vbforums.com/member.php?291519-qvb6
'           Elroy's superb balloon Tooltips.
'           Wqweto's superb TLS/STARTTLS code to enable email from VB6 using STARTTLS.
'           theTrick's superb sound code allowing recording of high quality sound.
'           Keith Lacelle for the alternative FSO code to read a value from an INI file when GetPrivateProfileString fails.
'               https://gist.github.com/Grimthorr/d17810f34cd361769ed0
'           Olaf Schmidt and his Date to Epoch code
'
' Tested on :
'           ReactOS 0.4.14 32bit on virtualBox
'           Windows 7 Professional 32bit on Intel
'           Windows 7 Ultimate 64bit on Intel
'           Windows 7 Professional 64bit on Intel
'           Windows XP SP3 32bit on Intel - not yet!
'           Windows 10 Home 64bit on Intel
'           Windows 10 Home 64bit on AMD
'
' Dependencies:
'           Krool's replacement for the Microsoft Windows Common Controls found in mscomctl.ocx (slider) is replicated
'           by the addition of one dedicated OCX file that is shipped with this package - CCRSlider.ocx
'
'           Microsoft ActiveX Data Objects 2.8 Library msador28.tlb as shipped with Windows XP +
'
'           You also need a reference to the Microsoft CDO for windows 2000 library component cdosys.dll
'               as ticked entry in the list of Project References. c:\windows\sysWoW64\cdosys.dll
'
'           requires a FireCallWin folder in C:\Users\<user>\AppData\Roaming\ eg: C:\Users\<user>\AppData\Roaming\FireCallWin
'           requires a Recordings folder in C:\Users\<user>\AppData\Roaming\FireCallWin eg: C:\Users\<user>\AppData\Roaming\FireCallWin\Recordings
'           requires a settings.ini file to exist in C:\Users\<user>\AppData\Roaming\FireCallWin
'           requires CCRSlider.ocx to exist in the program folder
'           requires an archive folder in app.path
'           requires a backup folder in app.path
'
' Notes:
'
' The VB6 non native images (PNGs &c) are displayed using Lavolpe's transparent DIB image code,
' except for the .ico files which use his earlier StdPictureEx class.
' Lavolpe's later ico code caused many strange visual artifacts and complete failures to show .ico files.
' especially when other image types were displayed on screen simultaneously.
'
' The sound is recorded using theTrick's sound code. It previously used MCISendString to record but Cortana on Win10+
' hijacks the sound device so it does not work on those oses.
'
' We have two comboboxes to store the audio input devices. The main combobox on the main form is used on form
' startup, reason this is done this way is because the enumeration must be done on form_load for the recording
' button to operate in HQ mode. Although we normally store the config. data in the prefs form, if we read that
' construct on startup it will try to load the whole prefs form and the prefs program variables are not ready
' for that to occur. Basically, we cannot have the combobox on another form and instead we keep the two in synch.

' The email is achieved using a tool from Microsoft called CDO, Firecall uses this to make the email point-to-point
' connection. Microsoft have failed to update CDO for a while so STARTTLS is not supported by default. In order to
' make STARTTLS function we have a proxy on port 10025 that takes any STARTTLS connection and manually injects the
' STARTTLS command into the stream just at the right time, for correct negotiation of a secure connection. The proxy
' forwards on the connection to the user's chosen port. This is the only way to make CDO negotiate a STARTTLS
' connection.
'
'---------------------------------------------------------------------------------

Option Explicit

'------------------------------------------------------ STARTS
' for SetWindowPos z-ordering
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOP As Long = 0 ' for SetWindowPos z-ordering
Private Const HWND_TOPMOST As Long = -1
Private Const HWND_BOTTOM As Long = 1
Private Const SWP_NOMOVE  As Long = &H2
Private Const SWP_NOSIZE  As Long = &H1
Private Const OnTopFlags  As Long = SWP_NOMOVE Or SWP_NOSIZE
Private Const SWP_NOACTIVATE  As Long = &H20
Private Const SWP_SHOWWINDOW  As Long = &H40



'------------------------------------------------------ ENDS



'------------------------------------------------------ STARTS
' alternative to comdlg32
Private Declare Function SHFileOperation Lib "Shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long 'alternative to comdlg32

Private Type SHFILEOPSTRUCT  'alternative to comdlg32
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long ' only used if FOF_SIMPLEPROGRESS, sets dialog title
End Type

Private Const FO_COPY  As Long = &H2 ' Copy File/Folder
Private Const FOF_SIMPLEPROGRESS As Long = &H100 ' Does not display file names
'------------------------------------------------------ ENDS





'------------------------------------------------------ STARTS
' Testing URLs
'Private Const S_FALSE = &H1
Private Const S_OK = &H0
Private Declare Function IsValidURL Lib "URLMON.DLL" (ByVal pbc As Long, ByVal szURL As String, ByVal dwReserved As Long) As Long
'------------------------------------------------------ ENDS



'------------------------------------------------------ STARTS
Private Const FOURCC_MEM      As Long = &H204D454D
Private Const MMIO_CREATERIFF As Long = &H20
Private Const MMIO_DIRTY      As Long = &H10000000
Private Const MMIO_CREATE     As Long = &H1000
Private Const MMIO_WRITE      As Long = &H1
Private Const MMIO_READWRITE  As Long = &H2
Private Const WAVE_FORMAT_PCM As Long = 1
Private Const SEEK_SET        As Long = 0
Private Const MMIO_FINDCHUNK  As Long = &H10
Private Const MMIO_FINDRIFF   As Long = &H20

Private Type MMCKINFO
    ckid            As Long
    ckSize          As Long
    fccType         As Long
    dwDataOffset    As Long
    dwFlags         As Long
End Type

Private Type WAVEFORMATEX
    wFormatTag      As Integer
    nChannels       As Integer
    nSamplesPerSec  As Long
    nAvgBytesPerSec As Long
    nBlockAlign     As Integer
    wBitsPerSample  As Integer
    cbSize          As Integer
End Type

Private PBK_NUMOFCHANNELS As Long  '= 1 '2     ' 1
Private PBK_SAMPLERATE     As Long '= 5512  ' 11025 ' 44100 ' 22050
Private PBK_BITNESS        As Long '= 16
Private PBK_BUFFERSIZEMS   As Single '= 0.3

Private Declare Function mmioClose Lib "winmm.dll" ( _
                         ByVal hmmio As Long, _
                         Optional ByVal uFlags As Long) As Long
Private Declare Function mmioOpen Lib "winmm.dll" _
                         Alias "mmioOpenW" ( _
                         ByVal szFileName As Long, _
                         ByRef lpmmioinfo As Any, _
                         ByVal dwOpenFlags As Long) As Long
Private Declare Function mmioStringToFOURCC Lib "winmm.dll" _
                         Alias "mmioStringToFOURCCA" ( _
                         ByVal sz As String, _
                         ByVal uFlags As Long) As Long
Private Declare Function mmioAscend Lib "winmm.dll" ( _
                         ByVal hmmio As Long, _
                         ByRef lpck As MMCKINFO, _
                         ByVal uFlags As Long) As Long
Private Declare Function mmioCreateChunk Lib "winmm.dll" ( _
                         ByVal hmmio As Long, _
                         ByRef lpck As MMCKINFO, _
                         ByVal uFlags As Long) As Long
Private Declare Function mmioWrite Lib "winmm.dll" ( _
                         ByVal hmmio As Long, _
                         ByRef pch As Any, _
                         ByVal cch As Long) As Long
Private Declare Function mmioDescend Lib "winmm.dll" ( _
                         ByVal hmmio As Long, _
                         ByRef lpck As MMCKINFO, _
                         ByRef lpckParent As Any, _
                         ByVal uFlags As Long) As Long


Dim WithEvents tSound   As clsTrickSound2
Attribute tSound.VB_VarHelpID = -1

Dim IsRecording As Boolean
Dim IsPlayback  As Boolean
Dim capBuffer() As Integer
Dim capCount    As Long
Dim plyIndex    As Long
'------------------------------------------------------ ENDS

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
     
     
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
     

Private Const PI As Double = 3.141592654
Private opacitylevel As Long
Private lastRandomText As String
Private dropTimerCount As Long
Private flashVal As Integer
Private flashCount As Integer
Private controlPressed  As String
Private vbKeyCPressed As Boolean
Private vbKeyFPressed As Boolean
Private vbKeyF1Pressed As Boolean
Private vbKeyF3Pressed As Boolean
Private vbKeyF5Pressed As Boolean
Private storedSearchString As String
Private storedSearchLineNo As Integer
Private buzzerCnt As Integer
Private recordingTimerCount As Integer
Private playingTimerCount As Integer
Private playingTimerMax As Integer

Private recording As Boolean
Private playing As Boolean
Private foundRecording As Boolean

Private WithEvents m_oProxy As cSmtpProxy
Attribute m_oProxy.VB_VarHelpID = -1

' timer and vars necessary to allow the animation on the config button
Private totalBusyCounter As Integer
Private busyCounter As Integer


'---------------------------------------------------------------------------------------
' Procedure : configBusyTimer_Timer
' Author    : beededea
' Date      : 11/09/2022
' Purpose   : do the hourglass timer
'---------------------------------------------------------------------------------------
'
Private Sub configBusyTimer_Timer()
    Dim busyFilename As String

    On Error GoTo configBusyTimer_Timer_Error

    busyFilename = ""

    totalBusyCounter = totalBusyCounter + 1
    busyCounter = busyCounter + 1
    If busyCounter >= 7 Then busyCounter = 1
    busyFilename = App.Path & "\Resources\images\config-busy" & busyCounter & ".jpg"
    btnPicConfig.Picture = LoadPicture(busyFilename)
    
    If totalBusyCounter >= 20 Then
        Call makeConfigAvailable
        
        configBusyTimer.Enabled = False
        busyCounter = 1
        totalBusyCounter = 1
        btnPicConfig.Refresh

        busyFilename = App.Path & "\Resources\images\btnConfig" & ".jpg"
        btnPicConfig.Picture = LoadPicture(busyFilename)
    End If

    On Error GoTo 0
    Exit Sub

configBusyTimer_Timer_Error:

    With err
         If .Number <> 0 Then
            MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure configBusyTimer_Timer of Form FireCallMain"
            Resume Next
          End If
    End With
End Sub

'Note all new events and procedures are moved to the bottom of this file, top event space is reserved for the main form events.

' set the focus to the text entry field whenever the form itself is clicked. The same done for almost all other controls on the form.
'---------------------------------------------------------------------------------------
' Procedure : Form_Click
' Author    : beededea
' Date      : 02/05/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Click()
     
   On Error GoTo Form_Click_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity

    txtTextEntry.SetFocus

   On Error GoTo 0
   Exit Sub

Form_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure Form_Click of Form FireCallMain"
End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : Form_KeyDown
' Author    : beededea
' Date      : 02/05/2025
' Purpose   : the standard Form_KeyDown routine that captures all keypresses on the form - keyPreview = true
'---------------------------------------------------------------------------------------
'
Private Sub Form_KeyDown(KeyCode As Integer, ByRef Shift As Integer)
   On Error GoTo Form_KeyDown_Error

    Call getKeyPress(KeyCode)

   On Error GoTo 0
   Exit Sub

Form_KeyDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure Form_KeyDown of Form FireCallMain"
End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 25/03/2025
' Purpose   : form_load is just used to initialise some vars the real work is done by formLoadTasks
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    ' initialise some global variables, cannot be in the form_initialise as that runs after form_load
    
   On Error GoTo Form_Load_Error

    Randomize
    
    Set messageQueue = New Collection
    
    ' Create new TrickSound object
    Set tSound = New clsTrickSound2
    
    Set validImageArrayList = New Collection
    Set invalidImageArrayList = New Collection
    Set executableSuffixArrayList = New Collection
    
'    pollingTimerID = 1 ' do not enable these, for reference only
'    iconiseTimerID = 2
'    emailTimerID = 3

    Dim outputFileArray(0)
    Dim inputFileArray(0)
    Dim combinedFileArray(0)
    
    dropTimerCount = 0
    opacitylevel = 0
    inputDataChangedFlag = False
    outputDataChangedFlag = False

    flashVal = 0
    flashCount = 0
    currindex = 0
    currentOpacity = 255
    
    mainMnuPopmenu.Visible = False
    'Me.Height = 10590
    btnLid.Left = 135 ' sets the emoji lid position at runtime as it tends to get moved around within the IDE
    picBtnLidShadow.Left = 200
    btnLid.Top = 5630
    picBtnLidShadow.Top = 5650
    picEmoji.Top = -1200
                    

                    
                    
    controlPressed = vbNullString
    CTRL_1 = False
    vbKeyCPressed = False
    vbKeyFPressed = False
    vbKeyF1Pressed = False
    vbKeyF3Pressed = False
    vbKeyF5Pressed = False
    
    dropboxErrCnt = 0
    
    buzzerCnt = 0
    FCWLastSoundPlayed = vbNullString
    FCWLastAwakeString = vbNullString
    FCWLastShutdown = vbNullString
    FCWAllowShutdowns = vbNullString
    storedSearchString = vbNullString
    storedSearchLineNo = 0
    
    nowBeingModifiedFlag = False
    
    msgBoxShowing = False
    
    ioMethodADO = False
            
    PBK_NUMOFCHANNELS = 1 '2     ' 1
    PBK_SAMPLERATE = 5512      ' 11025 ' 44100 ' 22050
    PBK_BITNESS = 16
    PBK_BUFFERSIZEMS = 0.3
    
    emailTString = "Kantancerous"
    
    msgBoxOut = True
    msgLogOut = True
    
    ' read available audio input devices
    Call enumerateRecordingDevices
    
    Call formLoadTasks ' < just so we can call the same routine from other places, you cannot call Form_Load

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure Form_Load of Form FireCallMain"
    
End Sub


'
'---------------------------------------------------------------------------------------
' Procedure : formLoadTasks
' Author    : beededea
' Date      : 25/03/2025
' Purpose   : the standard form_load routine calls our own routine that can be called directly elsewhere
'---------------------------------------------------------------------------------------
'
Public Sub formLoadTasks()

    Dim slicence As Integer
    
    ' assign some global variable values to valid amounts, do not remove as these are reset regularly when
    ' formLoadTasks is run (after saving the prefs. for example)
        
    On Error GoTo formLoadTasks_Error

    slicence = 0
    dropTimerCount = 0
    opacitylevel = 255
    inputDataChangedFlag = False
    outputDataChangedFlag = False
    flashVal = 0
    flashCount = 0
    currindex = 0
    recordingTimerCount = 0
    recording = False
    playing = False
    
    
    Call addExecutableSuffixArrayList

    'validImageTypes = ".jpg,.jpeg,.bmp,.ico,.png,.tif,.tiff,.gif,.cur,.wmf,.emf"
    Call addValidImageTypes
    
    'knownButInvalidImageTypes = ".pict,.icns,.ani,.svg,.NEF,.CR2,.ORF,.RW2,.ARW,.DNG,.wps,.AI,.PDF,.PSD,.RAW,.INDD"
    Call addInvalidImageTypes
    
    ' populate the emoji dropdown
    Call populateEmoji
    
    ' get the location of this tool's settings file (appdata)
    Call getToolSettingsFile
    
    ' check first usage and display licence screen
    Call checkLicenceState(slicence)

    ' check the Windows version
    Call testWinVer(classicThemeCapable)
    
    ' read the dock settings from the new configuration file
    Call readSettingsFile("Software\FireCallWin", FCWSettingsFile)
    
    ' set the input and output listBoxes to first time run contents
    If slicence = 0 Then Call setListBoxFirstRun
    
    ' validate the inputs of any data from the input settings file
    Call validateInputs
    
    ' read the recording quality and set the sample rate
    Call setSampleRate
    
    ' call the testMissingFields function to check the missing fields related to the input and output filenames
    If fTestMissingFields = False Then '
        'Call btnConfig_Click
        Call btnPicConfig_Click
        Exit Sub
    End If
    
    ' call the testInputsOutputs function to check the entries related to the input and outputs
    If fTestInputsOutputs = False Then
        'Call btnConfig_Click
        Call btnPicConfig_Click
        Exit Sub
    End If

    ' set the backups
    Call setBackups
    
    inputFileModificationTime = FileDateTime(FCWSharedInputFile)
    oldInputFileModificationTime = inputFileModificationTime
    
    outputFileModificationTime = FileDateTime(FCWSharedOutputFile)
    oldOutputFileModificationTime = outputFileModificationTime
    
    ' check the selected monitor properties and determine the number of twips per pixel for this screen
    Call monitorProperties(FireCallMain)

    ' adjust position on startup, placing possibly lost form onto correct monitor
    Call makeVisibleFormElements
    
    'alter the state of controls on the form, comboboxes mainly
    Call adjustMainControls
   
    ' set/unset the tooltips for all the form's controls
    Call setTooltips

    ' populate the input listbox
    Call populateInputBox
    
    ' populate the output listbox
    Call populateOutputBox
    
    ' populate the combined listbox
    If FCWSingleListBox = "1" Then Call populateCombinedBox
    
    If FCWSingleListBox = "1" Then
        mnuSwitchChatBoxes.Caption = "Switch to Split Chat Box Mode"
    Else
        mnuSwitchChatBoxes.Caption = "Switch to Single Chat Box"
    End If
    
    'enable/disable the scrollbars for the input and output listboxes
    Call handleScrollbars
       
    ' set the z-ordering of the window
    Call setZOrder(True) ' only runs the z-reordering at certain points controlled by the boolean.
    
    ' enable a timer that sets the z-order dynamically
    'zOrderTimer.Enabled = True ' unused
    
    ' set the position/size of some visual items that require the frm to be visible
    Call setVisualItems
            
    ' start the iconise timer iconise the main form to the stamp icon, in code or default types when in the IDE
    Call startTheIconiseTimers
    
    ' call the checkDropboxRunning function to check the dropbox process is running
    If FCWCheckServiceProcesses = "1" Then
        If FCWServiceProvider = "0" Then
            If fCheckDropboxRunning = False Then
                remoteNetworkDisabled = True
                Exit Sub
            End If
        End If
    End If
    
'    ' call the checkGoogleDriveRunning function to check the GoogleDrive process is running
'    If FCWServiceProvider = "1" Then
'        If checkGoogleDriveRunning = False Then
'            foundGoogleDriveDisabled = True
'            Exit Sub
'        End If
'    End If
'
'    ' call the checkOneDriveRunning function to check the OneDrive process is running
'    If FCWServiceProvider = "1" Then
'        If checkOneDriveRunning = False Then
'            foundOneDriveDisabled = True
'            Exit Sub
'        End If
'    End If
    
    ' start the polling timers in code or default types when in the IDE
    Call startThePollingTimers


   On Error GoTo 0
   Exit Sub

formLoadTasks_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure formLoadTasks of Form FireCallMain"

End Sub

' We have two comboboxes to store the audio input devices. The main combobox on the main form is used on form
' startup, reason this is done this way is because the enumeration must be done on form_load for the recording
' button to operate in HQ mode. Although we normally store the config. data in the prefs form, if we read that
' construct on startup it will try to load the whole prefs form and the prefs program variables are not ready
' for that to occur.

' Basically, we cannot have the combobox on another form and instead we keep the two in synch.

'---------------------------------------------------------------------------------------
' Procedure : enumerateRecordingDevices
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub enumerateRecordingDevices()

    Dim sName As Variant
    Dim devCount As Integer
    
   On Error GoTo enumerateRecordingDevices_Error

    devCount = 0
    
    'If FCWCaptureMethod = "0" Then
        cmbHiddenCaptureDevices.Clear
                
        ' Fill combo with available playback devices
        For Each sName In tSound.CaptureDevices
            devCount = devCount + 1
            cmbHiddenCaptureDevices.AddItem sName
        Next
            
        If devCount = 0 Then
            recordingIsPossible = False
            debugLog "%Err-I-ErrorNumber 22 - No Audio Devices Found, the recording message functionality will be disabled."
        Else
            cmbHiddenCaptureDevices.ListIndex = 0
            cmbHiddenCaptureDevices.Text = "No recording devices found"
            recordingIsPossible = True
        End If
        

    'End If

   On Error GoTo 0
   Exit Sub

enumerateRecordingDevices_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure enumerateRecordingDevices of Form FireCallMain"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : setBackups
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setBackups()

   On Error GoTo setBackups_Error

    If FCWAutomaticBackups = "1" Then
        backupTimer.Enabled = True
    End If
    
    ' backup the output file
    If FCWBackupOnStart = "1" Then
        Call backupOutputFile(FCWSharedOutputFile, "startup")
    End If

   On Error GoTo 0
   Exit Sub

setBackups_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure setBackups of Form FireCallMain"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : setVisualItems
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setVisualItems()

    Dim CurrentDPI As Long
    Dim NewSize As Long
    
'    If FCWMaximiseFormX = "0" Then
'        'FireCallMain.Left = FireCallMain.Left - 300
'    Else
'        FireCallMain.Left = Val(FCWMaximiseFormX)
'    End If
'
'    If FCWMaximiseFormY = "0" Then
'        'FireCallMain.Top = FireCallMain.Top
'    Else
'        FireCallMain.Top = Val(FCWMaximiseFormY)
'    End If

   On Error GoTo setVisualItems_Error

    Me.Show ' explicitly show the form
    
    ' check the selected monitor properties and determine the number of twips per pixel for this screen
    ' only calling TwipsPerPixelX/Y once on startup
'    screenTwipsPerPixelX = fTwipsPerPixelX
'    screenTwipsPerPixelY = fTwipsPerPixelY

    NewSize = Val(FCWFormWidth) * CurrentDPI / 96

    If FCWFormWidth = "0" Then
        'FireCallMain.Top = FireCallMain.Top
    Else
        FireCallMain.Width = Val(FCWFormWidth)
    End If
    
'    MsgBox "Harry - send me this please " & vbCrLf _
'        & "screenTwipsPerPixelX - " & screenTwipsPerPixelX & vbCrLf _
'        & "Current DPI - " & CurrentDPI & vbCrLf & vbCrLf _
'        & "Screen Width - " & Screen.Width & " twips or " & Screen.Width / screenTwipsPerPixelX & " pixels " & vbCrLf _
'        & "Screen Height - " & Screen.Height & " twips or " & Screen.Height / screenTwipsPerPixelY & " pixels " & vbCrLf & vbCrLf _
'        & "Form Width - " & FCWFormWidth & " twips " & vbCrLf _
'        & "Form Height - " & FireCallMain.Height & " twips "

    ' default positions prior to any resizing/shifting
    Call putImageInPlace
    
    Call formResizeSub
    
    linRed.X2 = 540
    
    ' set focus to the input text box so we can start typing immediately
    txtTextEntry.Text = "Type your text here..." ' never rely on the IDE as this specific value is checked
    txtTextEntry.SetFocus

   On Error GoTo 0
   Exit Sub

setVisualItems_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure setVisualItems of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Resize
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Resize()
   On Error GoTo Form_Resize_Error
   
    If FireCallMain.WindowState = vbMinimized Then Exit Sub

    Call formResizeSub

   On Error GoTo 0
   Exit Sub

Form_Resize_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure Form_Resize of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : formResizeSub
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub formResizeSub()
' credit Magic Ink
' https://www.vbforums.com/showthread.php?824699-RESOLVED-Form-Placement-Considering-Aero-Borders
    
    Dim desiredClientHeight As Long
    Dim desiredClientMinWidth As Long
    Dim desiredClientMaxWidth As Long
    Dim windowBorderWidth As Long
    Dim a As Long
    
    
   On Error GoTo formResizeSub_Error

    desiredClientMinWidth = 10065
    desiredClientHeight = 10185
    desiredClientMaxWidth = 25000
    windowBorderWidth = 0
    a = 0
    

    
    ' Width and Heigth are the size of the component, including the borders
    ' ScaleWidth and ScaleHeight works together with ScaleLeft, ScaleTop and
    ' ScaleMode to define the coordinate system for the component. By default,
    ' ScaleTop and ScaleHeight are zero, and ScaleWidth and ScaleHeigth are Width and Height minus the border,
    ' in vbTwips (the default ScaleMode)
    
    ' width         = full window + theme border
    ' ScaleWidth    = window without any theme border
    windowBorderWidth = Me.Width - Me.ScaleWidth
    '
'    borderSizeLeft = fBorderSize(FireCallMain).Left
'    borderSizeRight = fBorderSize(FireCallMain).Right
'    borderSizeTop = fBorderSize(FireCallMain).Top
'    borderSizeBottom = fBorderSize(FireCallMain).Bottom
    
    If Me.Width > 25000 Then ' maximum
        windowBorderWidth = Me.Width - Me.ScaleWidth
        Me.Width = windowBorderWidth + desiredClientMaxWidth
        Exit Sub
    End If
    If Me.Width < 10185 Then ' minimum
        Me.Width = windowBorderWidth + desiredClientMinWidth
        Exit Sub
    End If
    
    Me.Height = Me.Height - Me.ScaleHeight + desiredClientHeight
    
'     Me.Width = WidthInPixels * (Width / ScaleWidth)
'    Me.Height = HeightInPixels * (Height / ScaleHeight)
    
    txtTextEntry.Width = Me.ScaleWidth - 3700
    btnSendText.Left = txtTextEntry.Width + 355
    ' 10185  9945 240
    ' 14590  14355 235
    ' 10305  10065 240
    
    ' 10905  10785 = 120
    ' 11100 - 10980 = 120
    
    picSideBar.Left = Me.ScaleWidth - 2655 '+ Abs(fBorderSize(Me).Right)  ' 2715
    'picSideBar.Left = 9945 - 2655
    
    lbxOutputTextArea.Width = picSideBar.Left - 120
    lbxInputTextArea.Width = picSideBar.Left - 120
    lbxCombinedTextArea.Width = picSideBar.Left - 120
    'Me.Refresh
    'picSideBar.Refresh
    
    'DoEvents
    

   On Error GoTo 0
   Exit Sub

formResizeSub_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure formResizeSub of Form FireCallMain"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : thisForm_Unload
' Author    : beededea
' Date      : 18/08/2022
' Purpose   : the standard form unload routine called from several places
'---------------------------------------------------------------------------------------
'
Public Sub thisForm_Unload() ' name follows VB6 standard naming convention
    On Error GoTo Form_Unload_Error

    Call saveMainFormPosition

    Call DestroyToolTip ' destroys any current tooltip
    
    Call unloadAllForms(True)

    On Error GoTo 0
    Exit Sub

Form_Unload_Error:

    With err
         If .Number <> 0 Then
            MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure thisForm_Unload of Class Module module1"
            Resume Next
          End If
    End With
End Sub



'---------------------------------------------------------------------------------------
' Procedure : unloadAllForms
' Author    : beededea
' Date      : 28/06/2023
' Purpose   : unload all VB6 forms
'---------------------------------------------------------------------------------------
'
Public Sub unloadAllForms(ByVal endItAll As Boolean)
    
    Dim NameProcess As String: NameProcess = ""
    Dim fcount As Integer: fcount = 0
    Dim useloop As Integer: useloop = 0
       
    On Error GoTo unloadAllForms_Error
    
    ' the following unload commands take a while to complete resulting in a seeming-delay after a close, this .hide does away with that
    
    Me.Hide
    
    ' stop all VB6 timers in the main form
    
    Call stopPollingTimer
    Call stopIconiseTimer

    recordTimer.Enabled = False
    PlayTimer.Enabled = False
    lampTimer.Enabled = False
    printerTimer.Enabled = False
    dropTimer.Enabled = False
    brightTimer.Enabled = False
    emailTimer.Enabled = False
    houseKeepingTimer.Enabled = False
    opacityFadeOutTimer.Enabled = False
    opacityFadeInTimer.Enabled = False
    zOrderTimer.Enabled = False
    shredderTimer.Enabled = False
    pausePrinterTimer.Enabled = False
    clockTimer.Enabled = False
    emailIconTimer.Enabled = False
    opacityToTimer.Enabled = False
    buzzerTimer.Enabled = False
    sendCommandTimer.Enabled = False
    inputScrollBarTimer.Enabled = False
    outputScrollBarTimer.Enabled = False
    combinedScrollBarTimer.Enabled = False
    shutdownTimer.Enabled = False
    backupTimer.Enabled = False
    configBusyTimer.Enabled = False
    
    FireCallPrefs.themeTimer.Enabled = False
    
    ' unload the native VB6 forms
    
    Unload about
    Unload FireCallPrefs
    Unload licence
    Unload MinimiseForm
    'Unload FireCallMain ' this will be unloaded at the end of the main forms' form_unload
    
    ' remove all variable references to each form in turn
    
    Set about = Nothing
    Set FireCallPrefs = Nothing
    Set licence = Nothing
    Set MinimiseForm = Nothing
    Set FireCallMain = Nothing
   
    On Error Resume Next
    
    If endItAll = True Then End

   On Error GoTo 0
   Exit Sub

unloadAllForms_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure unloadAllForms of Module Module1"
End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : Form_Unload
' Author    : beededea
' Date      : 18/08/2021
' Purpose   : the standard form unload routine
'---------------------------------------------------------------------------------------
'
Private Sub Form_Unload(Cancel As Integer)
    
    On Error GoTo Form_Unload_Error
    
    Call thisForm_Unload

    On Error GoTo 0
    Exit Sub

Form_Unload_Error:

    With err
         If .Number <> 0 Then
            debugLog "Error " & err.Number & " (" & err.Description & ") in procedure Form_Unload of Form FireCallMain"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : saveMainFormPosition
' Author    : beededea
' Date      : 04/08/2023
' Purpose   : called from several locations saves the form X,Y positions
'---------------------------------------------------------------------------------------
'
Public Sub saveMainFormPosition()

   On Error GoTo saveMainFormPosition_Error

    FCWMaximiseFormX = Str$(FireCallMain.Left)
    FCWMaximiseFormY = Str$(FireCallMain.Top)
    If Val(FCWFormWidth) <= 10185 Then
        FCWFormWidth = "10185"
    Else
        FCWFormWidth = Str$(FireCallMain.Width)
    End If
    
    PutINISetting "Software\FireCallWin", "maximiseFormX", FCWMaximiseFormX, FCWSettingsFile
    PutINISetting "Software\FireCallWin", "maximiseFormY", FCWMaximiseFormY, FCWSettingsFile
    PutINISetting "Software\FireCallWin", "formWidth", FCWFormWidth, FCWSettingsFile

   On Error GoTo 0
   Exit Sub

saveMainFormPosition_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure saveMainFormPosition of Module Module1"
    
End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : startTheIconiseTimers
' Author    : beededea
' Date      : 29/04/2025
' Purpose   : start the iconise timer that iconises the main form to the stamp icon
'---------------------------------------------------------------------------------------
'
Private Sub startTheIconiseTimers()

    Dim iconiseIntervalMillisecs As Long
    
    Dim sixtyFive As Long ' just used to avoid multiplying two integers
    Dim oneThousand As Long
    
   On Error GoTo startTheIconiseTimers_Error

    If fInIDE Then
        ' VB6 timers cannot exceed 65 seconds (65535 ms)
        If Val(FCWIconiseDelay) > 65 Then
            sixtyFive = 65
            oneThousand = 1000
            ' when multiplying two integer values and assigning to a long in the IDE it caused an overflow as the IDE
            ' is handling the two numbers internally as integers as they are both below 32768 when VB6 encounters them.
            ' declaring vars as longs is a workaround.
            
            ' iconiseIntervalMillisecs = 65 * 1000 '  < this fails even though iconiseIntervalMillisecs is a long
            iconiseIntervalMillisecs = sixtyFive * oneThousand ' works!
            
        Else
            iconiseIntervalMillisecs = Val(FCWIconiseDelay) * 1000
        End If
        iconiseTimer.Interval = iconiseIntervalMillisecs
        iconiseTimer.Enabled = True
    Else
        ' using a timer in code rather than a VB6 timer as VB6 timers cannot exceed 65 seconds (65535 ms)
        ' and if you want a longer timer you have to roll your own.
        ' in addition, unfortunately the manual code timer method does not work in the IDE
        Call initiateIconiseTimerInCode
    End If

   On Error GoTo 0
   Exit Sub

startTheIconiseTimers_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure startTheIconiseTimers of Form FireCallMain"
End Sub

'
'
'---------------------------------------------------------------------------------------
' Procedure : handleScrollbars
' Author    : beededea
' Date      : 29/04/2025
' Purpose   : the listboxes have a vertical scrollbar by default and we add a horizontal scrollbar
'                showing/hiding these require different methods
'---------------------------------------------------------------------------------------
'
Private Sub handleScrollbars()
    Dim lLength As Long
    
    'disable the scrollbars for the input listbox
   On Error GoTo handleScrollbars_Error

    If FCWEnableScrollbars = "0" Then
        Call SendMessageByNum(lbxInputTextArea.hwnd, LB_SETHORIZONTALEXTENT, 0, 0&)
        Call ShowScrollBar(lbxInputTextArea.hwnd, SB_VERT, False)  ' hides the vertical scrollbar
    Else
        Call ShowScrollBar(lbxInputTextArea.hwnd, SB_VERT, True) ' shows the vertical scrollbar
        ' add the horizontal scroll bar to the upper listbox
        lLength = 2 * (lbxInputTextArea.Width / Screen.TwipsPerPixelX)
        Call SendMessageByNum(lbxInputTextArea.hwnd, LB_SETHORIZONTALEXTENT, lLength, 0&)
    End If
    
    'disable the scrollbars for the output listbox
    If FCWEnableScrollbars = 0 Then
        'the next two line must be in this order
        Call SendMessageByNum(lbxOutputTextArea.hwnd, LB_SETHORIZONTALEXTENT, 0, 0&) ' hides the horizontal scrollbar
        Call ShowScrollBar(lbxOutputTextArea.hwnd, SB_VERT, False)  ' hides the vertical scrollbar
    Else
        Call ShowScrollBar(lbxOutputTextArea.hwnd, SB_VERT, True) ' shows the vertical scrollbar
        ' add the horizontal scroll bar to the upper listbox
        lLength = 2 * (lbxOutputTextArea.Width / Screen.TwipsPerPixelX)
        Call SendMessageByNum(lbxOutputTextArea.hwnd, LB_SETHORIZONTALEXTENT, lLength, 0&)
    End If
    
    ' the scrollbar config code must be here after the reading of the combined data
    If FCWEnableScrollbars = 0 Then
        'the next two line must be in this order
        Call SendMessageByNum(FireCallMain.lbxCombinedTextArea.hwnd, LB_SETHORIZONTALEXTENT, 0, 0&) ' hides the horizontal scrollbar
        Call ShowScrollBar(FireCallMain.lbxCombinedTextArea.hwnd, SB_VERT, False)  ' hides the vertical scrollbar
    Else
        Call ShowScrollBar(FireCallMain.lbxCombinedTextArea.hwnd, SB_VERT, True) ' shows the vertical scrollbar
        ' add the horizontal scroll bar to the upper listbox
        lLength = 2 * (FireCallMain.lbxCombinedTextArea.Width / Screen.TwipsPerPixelX)
        Call SendMessageByNum(FireCallMain.lbxCombinedTextArea.hwnd, LB_SETHORIZONTALEXTENT, lLength, 0&)
    End If

   On Error GoTo 0
   Exit Sub

handleScrollbars_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure handleScrollbars of Form FireCallMain"

End Sub




'
'---------------------------------------------------------------------------------------
' Procedure : setZOrder
' Author    : beededea
' Date      : 29/04/2025
' Purpose   : set the Zorder of the main window, emulating functionality of the YWE version
'---------------------------------------------------------------------------------------
'
Private Sub setZOrder(ByVal formLoad As Boolean)
    
   On Error GoTo setZOrder_Error

    If Val(FCWWindowLevel) = 0 Then
        Call setFormPosition(Me, HWND_BOTTOM)
    ElseIf Val(FCWWindowLevel) = 1 Then
        If formLoad = True Then Call setFormPosition(Me, HWND_TOP)
    ElseIf Val(FCWWindowLevel) = 2 Then
        Call setFormPosition(Me, HWND_TOPMOST)
    End If

   On Error GoTo 0
   Exit Sub

setZOrder_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure setZOrder of Form FireCallMain"
End Sub
 

'
'---------------------------------------------------------------------------------------
' Procedure : fTestInputsOutputs
' Author    : beededea
' Date      : 29/04/2025
' Purpose   : check that the three required preference settings have valid values.
'---------------------------------------------------------------------------------------
'
Private Function fTestInputsOutputs() As Boolean
   On Error GoTo fTestInputsOutputs_Error

    fTestInputsOutputs = True
    
    If Not FCWSharedInputFile = vbNullString And Not fFExists(FCWSharedInputFile) Then
        MsgBox ("%Err-I-ErrorNumber 01 - The Shared Input File you have set is not accessible.")
        fTestInputsOutputs = False
        Exit Function
    End If
    If Not FCWSharedOutputFile = vbNullString And Not fFExists(FCWSharedOutputFile) Then
        MsgBox ("%Err-I-ErrorNumber 02 - The Shared Output File you have set is not accessible.")
        fTestInputsOutputs = False
        Exit Function
    End If
    If Not FCWSharedInputFile = vbNullString And Not fDirExists(FCWExchangeFolder) Then
        MsgBox ("%Err-I-ErrorNumber 03 - The Exchange Folder you have set is not accessible.")
        fTestInputsOutputs = False
        Exit Function
    End If

   On Error GoTo 0
   Exit Function

fTestInputsOutputs_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fTestInputsOutputs of Form FireCallMain"
    
End Function


'
'---------------------------------------------------------------------------------------
' Procedure : fTestMissingFields
' Author    : beededea
' Date      : 29/04/2025
' Purpose   : check that the three required preference settings have values, valid or not
'---------------------------------------------------------------------------------------
'
Private Function fTestMissingFields() As Boolean
   On Error GoTo fTestMissingFields_Error

    fTestMissingFields = True
    
    If FCWSharedInputFile = vbNullString Then
        MsgBox ("Please set the Shared Input File in the preferences.")
        fTestMissingFields = False
        Exit Function
    End If
    If FCWSharedOutputFile = vbNullString Then
        MsgBox ("Please set the Shared Output File in the preferences.")
        fTestMissingFields = False
        Exit Function
    End If
    If FCWExchangeFolder = vbNullString Then
        MsgBox ("Please create and set the Exchange Folder in the preferences.")
        fTestMissingFields = False
        Exit Function
    End If

   On Error GoTo 0
   Exit Function

fTestMissingFields_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fTestMissingFields of Form FireCallMain"
End Function

'
'---------------------------------------------------------------------------------------
' Procedure : btnClose_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   : call the same form unload subroutine called by the form unloading itself
'---------------------------------------------------------------------------------------
'
Private Sub btnClose_Click()
   On Error GoTo btnClose_Click_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    
    Call thisForm_Unload

   On Error GoTo 0
   Exit Sub

btnClose_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnClose_Click of Form FireCallMain"
End Sub



'
'---------------------------------------------------------------------------------------
' Procedure : btnPicAttach_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   : attach a single file to send to the remote chat partner
'---------------------------------------------------------------------------------------
'
Private Sub btnPicAttach_Click()
    Dim retFileName As String
    'Dim retfileTitle As String
    Dim attachedFile As String
    Dim fileNameToCopy As String
    Dim answer As VbMsgBoxResult
    
    'initialise the dimensioned variables
   On Error GoTo btnPicAttach_Click_Error

    answer = vbYes
    attachedFile = vbNullString
    
    If currentOpacity < 255 Then Call restoreMainWindowOpacity

    Call addTargetFile(attachedFile, retFileName)
    
    txtTextEntry.SetFocus ' brings the app to the font
    
    If retFileName <> vbNullString Then
    
        'retFileName = RTrim$(retFileName) ' this does NOT strip the padded fixed length, null padded string

        txtHiddenRetFileName.Text = retFileName ' just assigning it to a text field strips the buffered bit, leaving just the filename
        ' in this case we assign it to a hidden text box designed just for this functionality
        retFileName = txtHiddenRetFileName.Text
        
        fileNameToCopy = fGetFileNameFromPath(retFileName) ' remove the path
        
        
        
        If retFileName = FCWExchangeFolder & "\" & fileNameToCopy Then
            MsgBox ("%Err-I-ErrorNumber 04 - Both input and output files are the same file in the same location. Attach failed.")
            Exit Sub
        End If
        
        If fFExists(FCWExchangeFolder & "\" & fileNameToCopy) Then
            answer = MsgBox("This file already exists in this location, do you wish to overwrite?", vbExclamation + vbYesNo)
        End If
    
        If answer = vbYes Then
            FileCopy retFileName, FCWExchangeFolder & "\" & fileNameToCopy
            Call sendSomething("<><>" & fileNameToCopy)
        End If
    
    End If

   On Error GoTo 0
   Exit Sub

btnPicAttach_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicAttach_Click of Form FireCallMain"
    
End Sub


'
'---------------------------------------------------------------------------------------
' Procedure : cmbEmojiSelection_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   : display the small resized icon in the small emoji box
'---------------------------------------------------------------------------------------
'
Private Sub cmbEmojiSelection_Click()
    
    Dim fullPath As String
    'Dim emojiSet As String
    
   On Error GoTo cmbEmojiSelection_Click_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    
    If FCWEmojiSetDesc = vbNullString Then FCWEmojiSetDesc = "standard"
    fullPath = App.Path & "\resources\Emojis\" & FCWEmojiSetDesc & "\base\" & cmbEmojiSelection.List(cmbEmojiSelection.ListIndex)
    
    If fFExists(fullPath) Then
        picEmojiSmall.Picture = LoadPicture(fullPath)
    End If
    
    picEmojiSmall.ScaleMode = 3 ' pixels
    picEmojiSmall.AutoRedraw = True
    picEmojiSmall.PaintPicture picEmojiSmall.Picture, _
    0, 0, picEmojiSmall.ScaleWidth, picEmojiSmall.ScaleHeight, _
    0, 0, picEmojiSmall.Picture.Width / 26.46, _
    picEmojiSmall.Picture.Height / 26.46
    
    picEmojiSmall.Picture = picEmojiSmall.Image
    'lbxInputTextArea.Refresh
    picEmojiSmall.Refresh

   On Error GoTo 0
   Exit Sub

cmbEmojiSelection_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmbEmojiSelection_Click of Form FireCallMain"
    
End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : btnEmojiSet_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   : send your emoji state to the chat partner
'---------------------------------------------------------------------------------------
'
Private Sub btnEmojiSet_Click()

    Dim fullPath As String
    
   On Error GoTo btnEmojiSet_Click_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity
        
    If FCWEmojiSetDesc = vbNullString Then FCWEmojiSetDesc = "standard"
    fullPath = App.Path & "\resources\Emojis\" & FCWEmojiSetDesc & "\telly\" & cmbEmojiSelection.List(cmbEmojiSelection.ListIndex)
   
    If fFExists(fullPath) Then
        picOutputEmoji.Picture = LoadPicture(fullPath)
    End If
    
    txtTextEntry.Text = "<o><o>" & fExtractFileNameNoSuffix(cmbEmojiSelection.List(cmbEmojiSelection.ListIndex))
    'txtTextEntry.Text = "<o><o> " & cmbEmojiSelection.List(cmbEmojiSelection.ListIndex)
    Call sendSomething(txtTextEntry.Text)
    txtTextEntry.Text = vbNullString
    btnLid.Visible = False
    picBtnLidCatch.Visible = False
    picBtnLidShadow.Visible = False
    txtTextEntry.SetFocus

   On Error GoTo 0
   Exit Sub

btnEmojiSet_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnEmojiSet_Click of Form FireCallMain"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : houseKeepingTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub houseKeepingTimer_Timer()
   On Error GoTo houseKeepingTimer_Timer_Error

    Call houseKeepingTimerLogic(False)

   On Error GoTo 0
   Exit Sub

houseKeepingTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure houseKeepingTimer_Timer of Form FireCallMain"
End Sub



'
'---------------------------------------------------------------------------------------
' Procedure : IconiseTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :  The VB6 Iconise timer the equivalent of the initiateIconiseTimerInCode
'---------------------------------------------------------------------------------------
'
Private Sub IconiseTimer_Timer()
    'Dim lastInputVar As LASTINPUTINFO
    
    ' disable this timer when working in the runtime
   On Error GoTo IconiseTimer_Timer_Error

    If Not fInIDE Then
        Exit Sub ' this timer should only work in the IDE
    End If
    
    If Val(FCWIconiseDelay) = 0 Then
        iconiseTimer.Enabled = False
        Exit Sub
    End If
    
    Call getIdleTime

    If idleTime > Val(FCWIconiseDelay) * 1000 Then
        If FCWIconiseDesktop = "True" Then
            opacityFadeOutTimer.Enabled = True
            MinimiseForm.Visible = True
            iconiseTimer.Enabled = False
        Else
        
        End If
    End If

   On Error GoTo 0
   Exit Sub

IconiseTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure IconiseTimer_Timer of Form FireCallMain"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : lbxCombinedTextArea_DblClick
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxCombinedTextArea_DblClick()
   On Error GoTo lbxCombinedTextArea_DblClick_Error

    picTextChangeBright.Visible = False ' set the change lamp to dull
    picTextChangeDull.Visible = True
    inputDataChangedFlag = False

    ' when using the keys to select the top list box, the scrollbar is always displayed even when switched off
    ' in this case we disable it two seconds after the last keypress by using a timer to disable it
    If LTrim$(Str$(FCWEnableScrollbars)) = "0" Then
        If combinedScrollBarTimer.Enabled = False Then combinedScrollBarTimer.Enabled = True
    Else
        combinedScrollBarTimer.Enabled = False
    End If
    Call lbxTextAreaClick(lbxCombinedTextArea, True)

   On Error GoTo 0
   Exit Sub

lbxCombinedTextArea_DblClick_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxCombinedTextArea_DblClick of Form FireCallMain"
End Sub

' interpret the keys pressed and identify to the program where the keypress occurred
'---------------------------------------------------------------------------------------
' Procedure : lbxCombinedTextArea_KeyDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxCombinedTextArea_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo lbxCombinedTextArea_KeyDown_Error

    controlPressed = "lbxCombinedTextArea"
    Call getKeyPress(KeyCode)

   On Error GoTo 0
   Exit Sub

lbxCombinedTextArea_KeyDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxCombinedTextArea_KeyDown of Form FireCallMain"
End Sub
'after a key has been pressed on the combined area undo the CTRL key var
'---------------------------------------------------------------------------------------
' Procedure : lbxCombinedTextArea_KeyUp
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxCombinedTextArea_KeyUp(KeyCode As Integer, Shift As Integer)
   On Error GoTo lbxCombinedTextArea_KeyUp_Error

    CTRL_1 = False

   On Error GoTo 0
   Exit Sub

lbxCombinedTextArea_KeyUp_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxCombinedTextArea_KeyUp of Form FireCallMain"
End Sub
' show the alternative right click menu and set the bulbs to dull
'---------------------------------------------------------------------------------------
' Procedure : lbxCombinedTextArea_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxCombinedTextArea_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo lbxCombinedTextArea_MouseDown_Error

    If Button = 2 Then
        mnuLBOpenSharedInputFile.Visible = True
        mnuLBOpenSharedOutputFile.Visible = True
        
        mnuOutputEditLine.Visible = False
        mnuOutputDeleteLine.Visible = False
        mnuInputCopyLine.Visible = False
        mnuInputQuoteLine.Visible = False
        mnuOutputCopyLine.Visible = False
        mnuOutputPasteLine.Visible = False
        mnuFindInput.Visible = False
        mnuFindOutput.Visible = False
        mnuOutputPasteLine.Visible = False
        mnuOutputPasteAndGo.Visible = False
        
'        mnuCombinedDeleteLine.Visible = True
'        mnuCombinedEditLine.Visible = True
        mnuCombinedCopyLine.Visible = True
        mnuCombinedQuoteLine.Visible = True
        mnuFindCombined.Visible = True
        
        DoEvents
        If Clipboard.GetText <> "" Then
            mnuCombinedPasteLine.Visible = True
            mnuCombinedPasteAndGo.Visible = True
        Else
            mnuCombinedPasteLine.Visible = False
            mnuCombinedPasteAndGo.Visible = False
        End If
        
        Me.PopupMenu listBoxMnuPopmenu, vbPopupMenuRightButton
    End If
    
    picTextChangeBright.Visible = False
    picTextChangeDull.Visible = True
    inputDataChangedFlag = False

   On Error GoTo 0
   Exit Sub

lbxCombinedTextArea_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxCombinedTextArea_MouseDown of Form FireCallMain"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : lbxCombinedTextArea_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxCombinedTextArea_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo lbxCombinedTextArea_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip lbxCombinedTextArea.hwnd, "The combined chat box contains both chat partner's texts and messages. This is both the input and output files' contents combined and then sorted.", _
                  TTIconInfo, "Help on the Combined Chat Box", , , , True

   On Error GoTo 0
   Exit Sub

lbxCombinedTextArea_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxCombinedTextArea_MouseMove of Form FireCallMain"
End Sub

' set the change lamp to dull when any activity is enountered in the input box - the scrollbars in this case
'---------------------------------------------------------------------------------------
' Procedure : lbxCombinedTextArea_Scroll
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxCombinedTextArea_Scroll()
   On Error GoTo lbxCombinedTextArea_Scroll_Error

    picTextChangeBright.Visible = False ' set the change lamp to dull
    picTextChangeDull.Visible = True
    inputDataChangedFlag = False
    lbxCombinedTextArea.ToolTipText = ""

   On Error GoTo 0
   Exit Sub

lbxCombinedTextArea_Scroll_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxCombinedTextArea_Scroll of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : lbxInputTextArea_DblClick
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxInputTextArea_DblClick()

   On Error GoTo lbxInputTextArea_DblClick_Error

    picTextChangeBright.Visible = False ' set the change lamp to dull
    picTextChangeDull.Visible = True
    inputDataChangedFlag = False

    ' when using the keys to select the top list box, the scrollbar is always displayed even when switched off
    ' in this case we disable it two seconds after the last keypress by using a timer to disable it
    If LTrim$(Str$(FCWEnableScrollbars)) = "0" Then
        If inputScrollBarTimer.Enabled = False Then inputScrollBarTimer.Enabled = True
    Else
        inputScrollBarTimer.Enabled = False
    End If
    
    Call lbxTextAreaClick(lbxInputTextArea, True)

   On Error GoTo 0
   Exit Sub

lbxInputTextArea_DblClick_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxInputTextArea_DblClick of Form FireCallMain"
End Sub

' interpret the keys pressed and identify to the program where the keypress occurred
'---------------------------------------------------------------------------------------
' Procedure : lbxInputTextArea_KeyDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxInputTextArea_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
   On Error GoTo lbxInputTextArea_KeyDown_Error

    controlPressed = "lbxInputTextArea"
    Call getKeyPress(KeyCode)

   On Error GoTo 0
   Exit Sub

lbxInputTextArea_KeyDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxInputTextArea_KeyDown of Form FireCallMain"
    
End Sub
'after a key has been pressed on the input area undo the CTRL key var
'---------------------------------------------------------------------------------------
' Procedure : lbxInputTextArea_KeyUp
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxInputTextArea_KeyUp(KeyCode As Integer, Shift As Integer)
   On Error GoTo lbxInputTextArea_KeyUp_Error

    CTRL_1 = False

   On Error GoTo 0
   Exit Sub

lbxInputTextArea_KeyUp_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxInputTextArea_KeyUp of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : lbxInputTextArea_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxInputTextArea_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo lbxInputTextArea_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip lbxInputTextArea.hwnd, "The top chat box contains your chat partner's texts and messages. This is known as the input box displaying the contents of the shared input file.", _
                  TTIconInfo, "Help on the Upper Chat Box", , , , True

   On Error GoTo 0
   Exit Sub

lbxInputTextArea_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxInputTextArea_MouseMove of Form FireCallMain"
End Sub

' set the change lamp to dull when any activity is enountered in the input box - the scrollbars in this case
'---------------------------------------------------------------------------------------
' Procedure : lbxInputTextArea_Scroll
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxInputTextArea_Scroll()
   On Error GoTo lbxInputTextArea_Scroll_Error

    picTextChangeBright.Visible = False ' set the change lamp to dull
    picTextChangeDull.Visible = True
    inputDataChangedFlag = False
    lbxInputTextArea.ToolTipText = ""

   On Error GoTo 0
   Exit Sub

lbxInputTextArea_Scroll_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxInputTextArea_Scroll of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : lbxOutputTextArea_DblClick
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxOutputTextArea_DblClick()
    ' when using the keys to select the top list box, the scrollbar is always displayed even when switched off
    ' in this case we disable it two seconds after the last keypress by using a timer to disable it

   On Error GoTo lbxOutputTextArea_DblClick_Error

    If LTrim$(Str$(FCWEnableScrollbars)) = "0" Then
        If outputScrollBarTimer.Enabled = False Then outputScrollBarTimer.Enabled = True
    Else
        outputScrollBarTimer.Enabled = False
    End If
    Call lbxTextAreaClick(lbxOutputTextArea, True)

   On Error GoTo 0
   Exit Sub

lbxOutputTextArea_DblClick_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxOutputTextArea_DblClick of Form FireCallMain"

End Sub

' interpret the keys pressed and identify to the program where the keypress occurred
'---------------------------------------------------------------------------------------
' Procedure : lbxOutputTextArea_KeyDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxOutputTextArea_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
   On Error GoTo lbxOutputTextArea_KeyDown_Error

    controlPressed = "lbxOutputTextArea"
    
    Call getKeyPress(KeyCode)

   On Error GoTo 0
   Exit Sub

lbxOutputTextArea_KeyDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxOutputTextArea_KeyDown of Form FireCallMain"
End Sub
'after a key has been pressed on the output area undo the CTRL key var
'---------------------------------------------------------------------------------------
' Procedure : lbxOutputTextArea_KeyUp
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxOutputTextArea_KeyUp(KeyCode As Integer, Shift As Integer)
   On Error GoTo lbxOutputTextArea_KeyUp_Error

    CTRL_1 = False

   On Error GoTo 0
   Exit Sub

lbxOutputTextArea_KeyUp_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxOutputTextArea_KeyUp of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : lbxOutputTextArea_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxOutputTextArea_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo lbxOutputTextArea_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip lbxOutputTextArea.hwnd, "The bottom chat contains your own texts and messages. This is the output box displaying the contents of the shared output file. Beneath your chat box is the text box where you type your messages, pressing the SEND button to dispatch the text.", _
                  TTIconInfo, "Help on the Lower Chat Box", , , , True

   On Error GoTo 0
   Exit Sub

lbxOutputTextArea_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxOutputTextArea_MouseMove of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : lbxOutputTextArea_Scroll
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxOutputTextArea_Scroll()
   On Error GoTo lbxOutputTextArea_Scroll_Error

    lbxOutputTextArea.ToolTipText = ""

   On Error GoTo 0
   Exit Sub

lbxOutputTextArea_Scroll_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxOutputTextArea_Scroll of Form FireCallMain"
End Sub

'add ping request to the listBox right click menus
Private Sub mnuLBoxSendPingRequest_Click()
    Call mnuSendPingRequest_Click
End Sub
'add awake call to the listBox right click menus
Private Sub mnuLBoxSendAwakeCall_Click()
    Call mnuSendAwakeCall_click
End Sub


' hides the main form by starting the timer to fade the form out
Private Sub mnuHideProgram_Click()
    
    picTextChangeBright.Visible = False
    picTextChangeDull.Visible = True
    inputDataChangedFlag = False
    
    opacityFadeOutTimer.Enabled = True
    MinimiseForm.Visible = True
    
End Sub
' open the preferences form
'---------------------------------------------------------------------------------------
' Procedure : btnPicConfig_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicConfig_Click()
    
   On Error GoTo btnPicConfig_Click_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity

   On Error GoTo 0
   Exit Sub

btnPicConfig_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicConfig_Click of Form FireCallMain"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : makeConfigAvailable
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub makeConfigAvailable()
   On Error GoTo makeConfigAvailable_Error

    If FireCallPrefs.Visible = False And FireCallPrefs.WindowState = vbNormal Then
        
        If FireCallPrefs.WindowState = vbMinimized Then
            FireCallPrefs.WindowState = vbNormal
        End If
        
        
        If FireCallMain.Left + FireCallMain.Width + 200 + FireCallPrefs.Width > screenWidthTwips Then
            FireCallPrefs.Left = FireCallMain.Left - (FireCallPrefs.Width + 200)
        Else
            FireCallPrefs.Left = FireCallMain.Left + FireCallMain.Width + 200
        End If
        
        FireCallPrefs.Top = FireCallMain.Top
        
        If FireCallPrefs.Left < 0 Then FireCallPrefs.Left = 0
        If FireCallPrefs.Top < 0 Then FireCallPrefs.Top = 0
                
        'turn off the timer during prefs operation
        Call stopPollingTimer
        
        FireCallPrefs.Visible = True  ' show it again
        FireCallPrefs.SetFocus
    End If

   On Error GoTo 0
   Exit Sub

makeConfigAvailable_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure makeConfigAvailable of Form FireCallMain"
End Sub
' read the assigned text messages for the ten preset buttons at the base of the chat window
'---------------------------------------------------------------------------------------
' Procedure : readButtonTexts
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub readButtonTexts(ByVal buttonNo As Integer, ByRef textMessageArray() As String, Optional ByRef msgCnt As Integer)

    Dim buttonmessage As String
    Dim foundMessage As Boolean
    'Dim useloop As Integer
    
   On Error GoTo readButtonTexts_Error

    buttonmessage = vbNullString
    foundMessage = False
    msgCnt = 0
    
    foundMessage = True
    msgCnt = 0
    Do Until foundMessage = False
        foundMessage = False
        msgCnt = msgCnt + 1
        buttonmessage = fGetINISetting("Software\FireCallWin", "button" & buttonNo & "message" & msgCnt, FCWSettingsFile)
        If buttonmessage <> vbNullString Then
            foundMessage = True
            textMessageArray(msgCnt) = buttonmessage
        End If
    Loop

   On Error GoTo 0
   Exit Sub

readButtonTexts_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure readButtonTexts of Form FireCallMain"
End Sub
' the user pressed the TTFN button - demonstrating the use of GOTO for my young boy
'---------------------------------------------------------------------------------------
' Procedure : btnPicTtfn_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicTtfn_Click()

    ' declaration of vars
    Dim rndResult As Integer
    Dim textMessageArray(10) As String
    Dim msgCnt As Integer
    
    ' initialisation of vars
   On Error GoTo btnPicTtfn_Click_Error

    rndResult = 0
    msgCnt = 0

    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    
    Call readButtonTexts(1, textMessageArray(), msgCnt)
    msgCnt = msgCnt - 1
    
reRunbtnPicTtfn_Click:
    rndResult = Int((msgCnt * Rnd) + 1)
    txtTextEntry.Text = textMessageArray(rndResult)
    If lastRandomText = txtTextEntry.Text Then GoTo reRunbtnPicTtfn_Click

    lastRandomText = txtTextEntry.Text

    'MsgBox txtTextEntry.Text
    Call sendSomething(txtTextEntry.Text)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box

   On Error GoTo 0
   Exit Sub

btnPicTtfn_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicTtfn_Click of Form FireCallMain"
End Sub

' the user pressed the WELL button - demonstrating the use of DO WHILE for my young boy
'---------------------------------------------------------------------------------------
' Procedure : btnPicWell_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicWell_Click()
    Dim rndResult As Integer
    Dim goodText As Boolean
    Dim textMessageArray(10) As String
    Dim msgCnt As Integer
    
   On Error GoTo btnPicWell_Click_Error

    rndResult = 0
    goodText = False
    
    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    
    Call readButtonTexts(2, textMessageArray(), msgCnt)
    msgCnt = msgCnt - 1
                
    Do While goodText = False
        rndResult = Int((msgCnt * Rnd) + 1)
        txtTextEntry.Text = textMessageArray(rndResult)
        If lastRandomText <> txtTextEntry.Text Then goodText = True
    Loop
    
    lastRandomText = txtTextEntry.Text

    Call sendSomething(txtTextEntry.Text)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box


   On Error GoTo 0
   Exit Sub

btnPicWell_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicWell_Click of Form FireCallMain"

End Sub

' the user pressed the NEWS button - demonstrating the use of DO LOOP UNTIL for my young boy
'---------------------------------------------------------------------------------------
' Procedure : btnPicNews_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicNews_Click()
    Dim rndResult As Integer
    Dim goodText As Boolean
    Dim textMessageArray(10) As String
    Dim msgCnt As Integer
    
   On Error GoTo btnPicNews_Click_Error

    rndResult = 0
    goodText = False
    
    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    
    Call readButtonTexts(3, textMessageArray(), msgCnt)
    msgCnt = msgCnt - 1

    Do
        rndResult = Int((msgCnt * Rnd) + 1)
        txtTextEntry.Text = textMessageArray(rndResult)
        If lastRandomText <> txtTextEntry.Text Then goodText = True
    Loop Until goodText = True
    
    lastRandomText = txtTextEntry.Text
    
    Call sendSomething(txtTextEntry.Text)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box

   On Error GoTo 0
   Exit Sub

btnPicNews_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicNews_Click of Form FireCallMain"

End Sub

' the user pressed the MORN button - demonstrating the use of DO LOOP UNTIL for my young boy
'---------------------------------------------------------------------------------------
' Procedure : btnPicMorn_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicMorn_Click()
    Dim rndResult As Integer
    Dim goodText As Boolean
    Dim textMessageArray(10) As String
    Dim msgCnt As Integer
    
   On Error GoTo btnPicMorn_Click_Error

    rndResult = 0
    goodText = False
    
    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    
    Call readButtonTexts(4, textMessageArray(), msgCnt)
    msgCnt = msgCnt - 1
    
    Do
        rndResult = Int((msgCnt * Rnd) + 1)
        txtTextEntry.Text = textMessageArray(rndResult)
        If lastRandomText <> txtTextEntry.Text Then goodText = True
    Loop Until goodText = True
    
    lastRandomText = txtTextEntry.Text
    
    Call sendSomething(txtTextEntry.Text)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box

   On Error GoTo 0
   Exit Sub

btnPicMorn_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicMorn_Click of Form FireCallMain"

End Sub

' the user pressed the WOT button - demonstrating the use of DO UNTIL LOOP for my young boy
'---------------------------------------------------------------------------------------
' Procedure : btnPicWot_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicWot_Click()
    Dim rndResult As Integer
    Dim goodText As Boolean
    Dim textMessageArray(10) As String
    Dim msgCnt As Integer
    
   On Error GoTo btnPicWot_Click_Error

    rndResult = 0
    goodText = False
    
    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    
    Call readButtonTexts(5, textMessageArray(), msgCnt)
    msgCnt = msgCnt - 1
    
    Do Until goodText = True
        rndResult = Int((msgCnt * Rnd) + 1)
        txtTextEntry.Text = textMessageArray(rndResult)
        If lastRandomText <> txtTextEntry.Text Then goodText = True
    Loop
    
    lastRandomText = txtTextEntry.Text
    
    Call sendSomething(txtTextEntry.Text)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box

   On Error GoTo 0
   Exit Sub

btnPicWot_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicWot_Click of Form FireCallMain"

End Sub

' the user pressed the WTH button - demonstrating the use of DO UNTIL LOOP for my young boy
'---------------------------------------------------------------------------------------
' Procedure : BtnPicWth_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub BtnPicWth_Click()
    Dim rndResult As Integer
    Dim goodText As Boolean
    Dim textMessageArray(10) As String
    Dim msgCnt As Integer
    
   On Error GoTo BtnPicWth_Click_Error

    rndResult = 0
    goodText = False
    
    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    
    Call readButtonTexts(6, textMessageArray(), msgCnt)
    msgCnt = msgCnt - 1

    Do
        rndResult = Int((msgCnt * Rnd) + 1)
        txtTextEntry.Text = textMessageArray(rndResult)
        If lastRandomText <> txtTextEntry.Text Then goodText = True
    Loop Until goodText = True
    
    lastRandomText = txtTextEntry.Text
    
    Call sendSomething(txtTextEntry.Text)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box

   On Error GoTo 0
   Exit Sub

BtnPicWth_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure BtnPicWth_Click of Form FireCallMain"

End Sub

' the user pressed the PRG button - demonstrating the use of DO LOOP WHILE for my young boy
'---------------------------------------------------------------------------------------
' Procedure : btnPicPrg_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicPrg_Click()
    Dim rndResult As Integer
    Dim goodText As Boolean
    Dim textMessageArray(10) As String
    Dim msgCnt As Integer
    
   On Error GoTo btnPicPrg_Click_Error

    rndResult = 0
    goodText = False
    
    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    
    Call readButtonTexts(7, textMessageArray(), msgCnt)
    msgCnt = msgCnt - 1

    Do
        rndResult = Int((msgCnt * Rnd) + 1)
        txtTextEntry.Text = textMessageArray(rndResult)
        If lastRandomText <> txtTextEntry.Text Then goodText = True
        ' exit do ' works in a loop
    Loop While goodText = False
    
    lastRandomText = txtTextEntry.Text
    
    Call sendSomething(txtTextEntry.Text)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box

   On Error GoTo 0
   Exit Sub

btnPicPrg_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicPrg_Click of Form FireCallMain"

End Sub

' the user pressed the gdn button - demonstrating the use of DO UNTIL LOOP for my young boy
'---------------------------------------------------------------------------------------
' Procedure : btnPicGdn_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicGdn_Click()
    Dim rndResult As Integer
    Dim goodText As Boolean
    Dim textMessageArray(10) As String
    Dim msgCnt As Integer
    
   On Error GoTo btnPicGdn_Click_Error

    rndResult = 0
    goodText = False
    
    If currentOpacity < 255 Then Call restoreMainWindowOpacity

    Call readButtonTexts(8, textMessageArray(), msgCnt)
    msgCnt = msgCnt - 1

    Do
        rndResult = Int((msgCnt * Rnd) + 1)
        txtTextEntry.Text = textMessageArray(rndResult)
        If lastRandomText <> txtTextEntry.Text Then goodText = True
    Loop Until goodText = True
    
    lastRandomText = txtTextEntry.Text
    
    Call sendSomething(txtTextEntry.Text)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box

   On Error GoTo 0
   Exit Sub

btnPicGdn_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicGdn_Click of Form FireCallMain"

End Sub

' the user pressed the BUSY button - demonstrating the use of DO UNTIL LOOP for my young boy
'---------------------------------------------------------------------------------------
' Procedure : btnPicBusy_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicBusy_Click()
    Dim rndResult As Integer
    Dim goodText As Boolean
    Dim textMessageArray(10) As String
    Dim msgCnt As Integer
    
   On Error GoTo btnPicBusy_Click_Error

    rndResult = 0
    goodText = False
    
    If currentOpacity < 255 Then Call restoreMainWindowOpacity

    Call readButtonTexts(9, textMessageArray(), msgCnt)
    msgCnt = msgCnt - 1

    Do
        rndResult = Int((msgCnt * Rnd) + 1)
        txtTextEntry.Text = textMessageArray(rndResult)
        If lastRandomText <> txtTextEntry.Text Then goodText = True
    Loop Until goodText = True
    
    lastRandomText = txtTextEntry.Text
    Call sendSomething(txtTextEntry.Text)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box

   On Error GoTo 0
   Exit Sub

btnPicBusy_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicBusy_Click of Form FireCallMain"

End Sub

' the user pressed the COD button - demonstrating the use of DO UNTIL LOOP for my young boy
'---------------------------------------------------------------------------------------
' Procedure : btnPicCod_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicCod_Click()
    Dim rndResult As Integer
    Dim goodText As Boolean
    Dim textMessageArray(10) As String
    Dim msgCnt As Integer
    
   On Error GoTo btnPicCod_Click_Error

    rndResult = 0
    goodText = False
    
    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    
    Call readButtonTexts(10, textMessageArray(), msgCnt)
    msgCnt = msgCnt - 1

    Do
        rndResult = Int((msgCnt * Rnd) + 1)
        txtTextEntry.Text = textMessageArray(rndResult)
        If lastRandomText <> txtTextEntry.Text Then goodText = True
    Loop Until goodText = True
    
    lastRandomText = txtTextEntry.Text
    Call sendSomething(txtTextEntry.Text)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box


   On Error GoTo 0
   Exit Sub

btnPicCod_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicCod_Click of Form FireCallMain"
    
End Sub

' the user pressed the out button - demonstrating the use of WHILE WEND for my young boy
'---------------------------------------------------------------------------------------
' Procedure : btnPicOut_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicOut_Click()
    Dim rndResult As Integer
    Dim goodText As Boolean
    Dim textMessageArray(10) As String
    Dim msgCnt As Integer
    
   On Error GoTo btnPicOut_Click_Error

    rndResult = 0
    goodText = False
    
    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    
    Call readButtonTexts(11, textMessageArray(), msgCnt)
    msgCnt = msgCnt - 1
    
    While goodText = False ' an example of a WEND loop for my boy to learn
        rndResult = Int((msgCnt * Rnd) + 1)
        txtTextEntry.Text = textMessageArray(rndResult)
        If lastRandomText <> txtTextEntry.Text Then goodText = True
        ' exit do ' does not work in a while wend
    Wend
    
    lastRandomText = txtTextEntry.Text
    
    Call sendSomething(txtTextEntry.Text)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box


   On Error GoTo 0
   Exit Sub

btnPicOut_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicOut_Click of Form FireCallMain"

End Sub



'refresh the two listboxes containing the chat
'---------------------------------------------------------------------------------------
' Procedure : btnRefresh_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnRefresh_Click()
    
   On Error GoTo btnRefresh_Click_Error

    picTimerLampBright.Visible = True
    picTimerLampDull.Visible = False
    picTimerLampBright.Refresh
    
    Call populateInputBox
    Call populateOutputBox
    If FCWSingleListBox = "1" Then Call populateCombinedBox
    
    lampTimer.Enabled = True
    
    picTextChangeBright.Visible = False
    picTextChangeDull.Visible = True
    inputDataChangedFlag = False

    'forces the two listboxes to a specific height regardless of fonts chosen
    lbxInputTextArea.Height = 4300
    lbxOutputTextArea.Height = 4300
    
    txtTextEntry.SetFocus ' set focus back to the text entry box

   On Error GoTo 0
   Exit Sub

btnRefresh_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnRefresh_Click of Form FireCallMain"
End Sub
' when clicking upon a line in the output box, display any image found in that line, also hide any unwanted scrollbars that VB6 automatically puts back
'---------------------------------------------------------------------------------------
' Procedure : lbxOutputTextArea_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxOutputTextArea_Click() '(Optional ByRef frm As Form)

    ' when using the keys to select the top list box, the scrollbar is always displayed even when switched off
    ' in this case we disable it two seconds after the last keypress by using a timer to disable it
    
   On Error GoTo lbxOutputTextArea_Click_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity

    If LTrim$(Str$(FCWEnableScrollbars)) = "0" Then
        If outputScrollBarTimer.Enabled = False Then outputScrollBarTimer.Enabled = True
    Else
        outputScrollBarTimer.Enabled = False
    End If
    Call lbxTextAreaClick(lbxOutputTextArea)

   On Error GoTo 0
   Exit Sub

lbxOutputTextArea_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxOutputTextArea_Click of Form FireCallMain"
    
End Sub

' when clicking upon a line in the output box, display any image found in that line, or act upon any URL found
'---------------------------------------------------------------------------------------
' Procedure : lbxTextAreaClick
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxTextAreaClick(Optional ByRef srcListBox As ListBox, Optional ByRef textAreaDblClickState As Boolean)
    Dim attachmentString As String
    Dim attachmentFilenamePos As Integer
    Dim attachmentFilename As String
    Dim recordingString As String
    Dim recordingFilenamePos As Integer
    Dim recordingFilename As String
    Dim suffix As String
    Dim suffixNoDot As String
    
    Dim strToSearch As String

    Dim extractedURL As String
    Dim preliminaryURL As String
    
    Dim answer As VbMsgBoxResult
    
    Dim foundFile As Boolean
    Dim foundFolder As Boolean
    Dim imgFilePath As String
    
   On Error GoTo lbxTextAreaClick_Error

    foundFile = False
    foundFolder = False
    foundRecording = False
    binaryFlag = False
    picImagePrintOut.ToolTipText = ""
    binaryFlag = False
    
    'initialise the dimensioned variables
    answer = vbNo
    
    picBtnPlaySound.Visible = False
    
    If InStr(srcListBox.List(srcListBox.ListIndex), "New Folder") > 0 Then foundFolder = True
    
    attachmentString = srcListBox.List(srcListBox.ListIndex)

    If InStr(1, attachmentString, "New File:") > 0 Then
        attachmentFilenamePos = InStr(srcListBox.List(srcListBox.ListIndex), "New File:") + 9
        attachmentFilename = Mid$(attachmentString, attachmentFilenamePos, Len(attachmentString))
        foundFile = True
        attachmentFilePath = FCWExchangeFolder & "\" & attachmentFilename

        If fExtractSuffixWithDot(attachmentFilePath) = ".m4a" Or fExtractSuffixWithDot(attachmentFilePath) = ".wav" Then
            recordingFilenamePos = InStr(srcListBox.List(srcListBox.ListIndex), "New File:") + 9
            recordingFilename = Mid$(attachmentString, attachmentFilenamePos, Len(attachmentString))
            foundRecording = True
            foundFile = False
            picBtnPlaySound.Visible = True
            recordingFilePath = FCWExchangeFolder & "\" & recordingFilename
        End If
    End If
        
    If InStr(1, attachmentString, "New Recording:") > 0 Then
        recordingFilenamePos = InStr(srcListBox.List(srcListBox.ListIndex), "New Recording:") + 14
        recordingFilename = Mid$(attachmentString, recordingFilenamePos, 23)
        foundRecording = True
        foundFile = False
        picBtnPlaySound.Visible = True
        recordingFilePath = FCWExchangeFolder & "\" & recordingFilename
        attachmentFilePath = recordingFilePath
    End If
                        
    If Not fFExists(RTrim$(attachmentFilePath)) Then
        Call displaySelectedImage(App.Path & "\resources\images\documentIcons\document-missing" & ".png")
        'If FCWEnableTooltips = "1" Then
        picImagePrintOut.ToolTipText = attachmentFilename & " This file is missing - it is no longer in the dropbox shared folder."
    Else
        If foundRecording = True Then
            Call displaySelectedImage(App.Path & "\resources\images\documentIcons\document-rec.png")
            If FCWEnableTooltips = "1" Then FireCallMain.picImagePrintOut.ToolTipText = recordingFilename & " - double click to play the recording."
            If textAreaDblClickState = True Then
                Call ShellExecute(Me.hwnd, "Open", recordingFilePath, vbNullString, App.Path, 1)
            End If
            recordingViewTime = Now
        End If
        
        If foundFile = True Or foundFolder = True Then
            picImagePrintOut.Visible = True
            imgFilePath = App.Path & "\Resources\images\lidBackgroundDullShadowed.jpg"
            If fFExists(imgFilePath) Then
                picLidBackground.Picture = LoadPicture(imgFilePath)
            End If
            
            If foundFile = True Then
                ' on a click we reassign the stored full file variable path displayedAttachmentFilePath as that is what is used during a dblClick on the image
                displayedAttachmentFilePath = attachmentFilePath
                'suffix = fExtractSuffix(displayedAttachmentFilePath)
                
                suffix = fExtractSuffixWithDot(displayedAttachmentFilePath)
                suffixNoDot = fExtractSuffix(displayedAttachmentFilePath)

                If fInstrSuffix(validImageArrayList, LCase(suffix)) Then
                    Call displaySelectedImage(displayedAttachmentFilePath)
                ElseIf fInstrSuffix(invalidImageArrayList, LCase(suffix)) <> 0 Then
                    Call displaySelectedImage(App.Path & "\resources\images\documentIcons\document-unknown" & ".png")
                Else
                    Call displaySelectedImage(App.Path & "\resources\images\documentIcons\document-" & suffixNoDot & ".png")

                End If
                If FCWEnableTooltips = "1" Then picImagePrintOut.ToolTipText = attachmentFilename & " - double click to open it using default app."
                
                suffix = fExtractSuffixWithDot(displayedAttachmentFilePath)
                If fInstrSuffix(executableSuffixArrayList, LCase(suffix)) Then
                    binaryFlag = True
                    picImagePrintOut.ToolTipText = attachmentFilename & " - This is an executable program - take care."
                End If
                
                If textAreaDblClickState = True Then
                    If binaryFlag = True Then
                        answer = MsgBox(attachmentFilePath & vbCrLf & vbCrLf & " This is an executable program, running it could be dangerous and unpredictable things may happen." & vbCrLf & vbCrLf & "Are you sure you wish to proceed?", vbExclamation + vbYesNo)
                    Else
                        answer = vbYes
                    End If
                    If answer = vbYes Then
                        If attachmentFilename = "FireCallWin.exe" Then
                            answer = MsgBox(attachmentFilePath & vbCrLf & vbCrLf & " This is the FireCallWin program, it cannot run itself again.", vbExclamation)
                        Else
                            Call ShellExecute(Me.hwnd, "Open", displayedAttachmentFilePath, vbNullString, App.Path, 1)
                        End If
                    End If
                End If

            End If
            
            If foundFolder = True Then
                Call displaySelectedImage(App.Path & "\resources\images\documentIcons\document-dir.png")
                If FCWEnableTooltips = "1" Then picImagePrintOut.ToolTipText = attachmentFilename & " - double click to open the folder in Explorer."
            
                If textAreaDblClickState = True Then
                    Call ShellExecute(Me.hwnd, "Open", attachmentFilePath, vbNullString, App.Path, 1)
                End If
            End If
    
            attachmentViewTime = Now
        End If
    End If
    srcListBox.ToolTipText = ""
    
    ' search the line for something that identifies an URL
    'CTRL_1 = False ' just in case this hasn't been handled by the keyup event
    strToSearch = srcListBox.List(srcListBox.ListIndex)
    ' use a list of search terms ie. http, https and www to see if this might be a URL
    If fMultiInstr(strToSearch, "ANY", "http", "https", "HTTP", "HTTPS", "www.", "WWW.") >= 0 Then
        
        ' if http, first test there are :// chars together in the string, else exit
        If InStr(LCase$(strToSearch), "http") = 0 And InStr(strToSearch, "://") = 0 Then Exit Sub
        
        ' extract the possible URL from the string
        preliminaryURL = Mid$(strToSearch, InStr(strToSearch, "http"))
        
        ' search for a full space denoting the end of a URL, or it is assumed to be the full line
        ' this does not yet handle URLs on split lines
        If InStr(preliminaryURL, " ") = 0 Then
            extractedURL = preliminaryURL
        Else
            extractedURL = Mid$(preliminaryURL, InStr(preliminaryURL, " "))
        End If
        
        ' use the WinAPI to validate the URL
        If fIsGoodURL(extractedURL) Then
            srcListBox.ToolTipText = extractedURL
            If textAreaDblClickState = True Then
                Call ShellExecute(Me.hwnd, "Open", extractedURL, vbNullString, App.Path, 1)
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

lbxTextAreaClick_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxTextAreaClick of Form FireCallMain"
    
End Sub
' KayJay
' utilises the isValidURL API function in Windows
'---------------------------------------------------------------------------------------
' Procedure : fIsGoodURL
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function fIsGoodURL(ByVal sURL As String) As Boolean
   On Error GoTo fIsGoodURL_Error

    sURL = StrConv(sURL, vbUnicode)
    'Now call the function
    fIsGoodURL = (IsValidURL(ByVal 0&, sURL, 0) = S_OK)

   On Error GoTo 0
   Exit Function

fIsGoodURL_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fIsGoodURL of Form FireCallMain"
End Function




' when clicking upon a line in the input box, display any image found in that line, also hide any unwanted scrollbars that VB6 automatically puts back
'---------------------------------------------------------------------------------------
' Procedure : lbxInputTextArea_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxInputTextArea_Click()
    
   On Error GoTo lbxInputTextArea_Click_Error

    picTextChangeBright.Visible = False ' set the change lamp to dull
    picTextChangeDull.Visible = True
    inputDataChangedFlag = False
    
    If currentOpacity < 255 Then Call restoreMainWindowOpacity

    ' when using the keys to select the top list box, the scrollbar is always displayed even when switched off
    ' in this case we disable it two seconds after the last keypress by using a timer to disable it
    If LTrim$(Str$(FCWEnableScrollbars)) = "0" Then
        If inputScrollBarTimer.Enabled = False Then inputScrollBarTimer.Enabled = True
    Else
        inputScrollBarTimer.Enabled = False
    End If
    
    Call lbxTextAreaClick(lbxInputTextArea)

   On Error GoTo 0
   Exit Sub

lbxInputTextArea_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxInputTextArea_Click of Form FireCallMain"
    
End Sub

' when clicking upon a line in the input box, display any image found in that line, also hide any unwanted scrollbars that VB6 automatically puts back
'---------------------------------------------------------------------------------------
' Procedure : lbxCombinedTextArea_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxCombinedTextArea_Click()

   On Error GoTo lbxCombinedTextArea_Click_Error

    picTextChangeBright.Visible = False ' set the change lamp to dull
    picTextChangeDull.Visible = True
    inputDataChangedFlag = False
    
    If currentOpacity < 255 Then Call restoreMainWindowOpacity

    ' when using the keys to select the top list box, the scrollbar is always displayed even when switched off
    ' in this case we disable it two seconds after the last keypress by using a timer to disable it
    If LTrim$(Str$(FCWEnableScrollbars)) = "0" Then
        If combinedScrollBarTimer.Enabled = False Then combinedScrollBarTimer.Enabled = True
    Else
        combinedScrollBarTimer.Enabled = False
    End If
    Call lbxTextAreaClick(lbxCombinedTextArea)

   On Error GoTo 0
   Exit Sub

lbxCombinedTextArea_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxCombinedTextArea_Click of Form FireCallMain"
End Sub
'captures a drag and drop to any of the listBoxes
'---------------------------------------------------------------------------------------
' Procedure : lbxInputTextArea_OLEDragDrop
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxInputTextArea_OLEDragDrop(Data As DataObject, Effect As Long, ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo lbxInputTextArea_OLEDragDrop_Error

    Call lbxOutputTextArea_OLEDragDrop(Data, Effect, Button, Shift, x, y)

   On Error GoTo 0
   Exit Sub

lbxInputTextArea_OLEDragDrop_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxInputTextArea_OLEDragDrop of Form FireCallMain"
End Sub

'captures a drag and drop to any of the listBoxes
'---------------------------------------------------------------------------------------
' Procedure : lbxCombinedTextArea_OLEDragDrop
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxCombinedTextArea_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo lbxCombinedTextArea_OLEDragDrop_Error

    Call lbxOutputTextArea_OLEDragDrop(Data, Effect, Button, Shift, x, y)

   On Error GoTo 0
   Exit Sub

lbxCombinedTextArea_OLEDragDrop_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxCombinedTextArea_OLEDragDrop of Form FireCallMain"
End Sub

'captures a drag and drop to any of the output listBoxes
'---------------------------------------------------------------------------------------
' Procedure : lbxOutputTextArea_OLEDragDrop
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxOutputTextArea_OLEDragDrop(Data As DataObject, Effect As Long, ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
    Dim iconTitle As String
    Dim fileNameToCopy As String
    Dim answer As VbMsgBoxResult
    
    Const wFlags As Long = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
    
    'initialise the dimensioned variables
   On Error GoTo lbxOutputTextArea_OLEDragDrop_Error

    answer = vbYes
    
    
    'only allow the drag and drop of files and not from one part of the listbox to the other
    If Data.GetFormat(vbCFFiles) = True Then
    
        SetWindowPos hwnd, -1&, 0, 0, 0, 0, wFlags
    
        ' if there is more than one file dropped reject the drop
        If Data.Files.Count > 1 Then
             MsgBox "%Err-I-ErrorNumber 05 - Sorry, can only accept one icon drop at a time, you have dropped " & Data.Files.Count, vbSystemModal + vbInformation
            Exit Sub
        End If

        iconTitle = Data.Files(1) ' set the title for all types
        
        txtTextEntry.SetFocus ' brings the app to the font ensuring the mgbox is to the fore
        'this brings the whole form to the fore but sometimes the explorer window might sit on top, works in conjunction with SetWindowPos hWnd, -2&, 0, 0, 0, 0, wFlags

        ' here we will check for a folder
        If fDirExists(iconTitle) Then
            fileNameToCopy = fGetFileNameFromPath(iconTitle)
            
            If iconTitle = FCWExchangeFolder & "\" & fileNameToCopy Then
                MsgBox ("%Err-I-ErrorNumber 06 - Both the input and output folders are the same, you are copying from and to the same location. Drag & drop failed.")
                Exit Sub
            End If
            
            If fDirExists(FCWExchangeFolder & "\" & fileNameToCopy) Then
                answer = MsgBox("A folder of this same name aready exists in this location, do you wish to overwrite?", vbExclamation + vbYesNo)
            End If
            
            If answer = vbYes Then
                Call VBCopyFolder(iconTitle, FCWExchangeFolder & "\" & fileNameToCopy)
                Call sendSomething("<f><f>" & fileNameToCopy)
            End If
        Else
            If fFExists(iconTitle) Then
                fileNameToCopy = fGetFileNameFromPath(iconTitle)
    
                If iconTitle = FCWExchangeFolder & "\" & fileNameToCopy Then
                    MsgBox ("%Err-I-ErrorNumber 07 - Both input and output files are the same file in the same location. Drag & drop failed.")
                    Exit Sub
                End If

                If fFExists(FCWExchangeFolder & "\" & fileNameToCopy) Then
                    answer = MsgBox("This file already exists in this location, do you wish to overwrite?", vbExclamation + vbYesNo)
                End If

                If answer = vbYes Then
                    FileCopy iconTitle, FCWExchangeFolder & "\" & fileNameToCopy
                    Call sendSomething("<><>" & fileNameToCopy)
                End If
            Else
                ' I have encountered folder names (probably created on some older file system) that contained ? chars in them when handled within VB6.
                If InStr(iconTitle, "?") Then
                    MsgBox ("%Err-I-ErrorNumber 08 - For some reason that filename is invalid, possibly containing disallowed chars. Drag & drop failed.")
                Else
                    MsgBox ("%Err-I-ErrorNumber 09 - The file you dragged to the program seems to be unavailable now. Drag & drop failed.")
                End If
            End If
        End If
    End If
    
    SetWindowPos hwnd, -2&, 0, 0, 0, 0, wFlags ' this brings the window to the fore but sometimes the explorer window might sit on top, the earlier .setfocus sorts this

   On Error GoTo 0
   Exit Sub

lbxOutputTextArea_OLEDragDrop_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxOutputTextArea_OLEDragDrop of Form FireCallMain"

End Sub


' Chris Fannin (AbbydonKrafts) http://vbcity.com/forums/t/129391.aspx
' allows the copying of a whole folder
'---------------------------------------------------------------------------------------
' Procedure : VBCopyFolder
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub VBCopyFolder(ByRef strSource As String, ByRef strTarget As String)

    Dim op As SHFILEOPSTRUCT
   On Error GoTo VBCopyFolder_Error

    With op
        .wFunc = FO_COPY ' Set function
        .pTo = strTarget ' Set new path
        .pFrom = strSource ' Set current path
        .fFlags = FOF_SIMPLEPROGRESS ' FOF_SILENT
    End With
    ' Perform operation
    SHFileOperation op

   On Error GoTo 0
   Exit Sub

VBCopyFolder_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure VBCopyFolder of Form FireCallMain"

End Sub
' menu options to do this and that
'---------------------------------------------------------------------------------------
' Procedure : mnuRefresh_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuRefresh_Click()
   On Error GoTo mnuRefresh_Click_Error

    If lbxOutputTextArea.Visible = True Then lbxOutputTextArea.Clear
    If lbxInputTextArea.Visible = True Then lbxInputTextArea.Clear
    If lbxCombinedTextArea.Visible = True Then lbxCombinedTextArea.Clear
    Call btnRefresh_Click

   On Error GoTo 0
   Exit Sub

mnuRefresh_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuRefresh_Click of Form FireCallMain"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuLBRefresh_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLBRefresh_Click()
   On Error GoTo mnuLBRefresh_Click_Error

    If lbxOutputTextArea.Visible = True Then lbxOutputTextArea.Clear
    If lbxInputTextArea.Visible = True Then lbxInputTextArea.Clear
    If lbxCombinedTextArea.Visible = True Then lbxCombinedTextArea.Clear
    Call btnRefresh_Click

   On Error GoTo 0
   Exit Sub

mnuLBRefresh_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuLBRefresh_Click of Form FireCallMain"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuLBOpenSharedInputFile_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLBOpenSharedInputFile_Click()
   On Error GoTo mnuLBOpenSharedInputFile_Click_Error

    Call mnuOpenSharedInputFile_Click

   On Error GoTo 0
   Exit Sub

mnuLBOpenSharedInputFile_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuLBOpenSharedInputFile_Click of Form FireCallMain"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuLBOpenSharedOutputFile_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLBOpenSharedOutputFile_Click()
   On Error GoTo mnuLBOpenSharedOutputFile_Click_Error

    Call mnuOpenSharedOutputFile_Click

   On Error GoTo 0
   Exit Sub

mnuLBOpenSharedOutputFile_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuLBOpenSharedOutputFile_Click of Form FireCallMain"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuCloseProgram_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuCloseProgram_Click()
   On Error GoTo mnuCloseProgram_Click_Error

    Call btnClose_Click

   On Error GoTo 0
   Exit Sub

mnuCloseProgram_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuCloseProgram_Click of Form FireCallMain"
End Sub

' make the Emoji lid disappear or show the right click menu

' show the right click menu
'---------------------------------------------------------------------------------------
' Procedure : cmbEmojiSelection_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbEmojiSelection_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo cmbEmojiSelection_MouseDown_Error

    If Button = 2 Then
        Me.PopupMenu mainMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

cmbEmojiSelection_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmbEmojiSelection_MouseDown of Form FireCallMain"

End Sub
' small button close form
'---------------------------------------------------------------------------------------
' Procedure : btnCloseIt_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnCloseIt_Click()
    
   On Error GoTo btnCloseIt_Click_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    Call btnClose_Click

   On Error GoTo 0
   Exit Sub

btnCloseIt_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnCloseIt_Click of Form FireCallMain"
    
End Sub

' use the win API to place the form in zorder
'---------------------------------------------------------------------------------------
' Procedure : setFormPosition
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub setFormPosition(ByRef frm As Form, ByVal fromPosition As Long)
   On Error GoTo setFormPosition_Error

    Call SetWindowPos(frm.hwnd, fromPosition, 0&, 0&, 0&, 0&, OnTopFlags)

   On Error GoTo 0
   Exit Sub

setFormPosition_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure setFormPosition of Form FireCallMain"
End Sub
' add the emoji filenames to the emoji dropdown
'---------------------------------------------------------------------------------------
' Procedure : populateEmoji
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub populateEmoji()
    Dim MyPath  As String
    'Dim themePresent As Boolean
    Dim myName As String


   On Error GoTo populateEmoji_Error

    If FCWEmojiSetDesc = vbNullString Then FCWEmojiSetDesc = "standard"
    MyPath = App.Path & "\resources\Emojis\" & FCWEmojiSetDesc & "\base\"

    ' populate the emoji box with any .jpg files that exist
    myName = Dir(MyPath, vbNormal)   ' Retrieve the first entry.
    Do While myName <> vbNullString   ' Start the loop.
       myName = Dir   ' Get next entry.
       If myName <> "." And myName <> ".." And myName <> vbNullString And fExtractSuffixWithDot(myName) = ".jpg" Then
        cmbEmojiSelection.AddItem myName
        'Debug.Print myName
       End If
    Loop
    cmbEmojiSelection.ListIndex = 0
    'cmbEmojiSelection.SelLength = 0

   On Error GoTo 0
   Exit Sub

populateEmoji_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure populateEmoji of Form FireCallMain"
    
End Sub
    
'---------------------------------------------------------------------------------------
' Procedure : validateInputs
' Author    : beededea
' Date      : 17/06/2020
' Purpose   : validate the relevant entries from the settings.ini file in user appdata
'---------------------------------------------------------------------------------------
'
Public Sub validateInputs()
    
   ' On Error GoTo validateInputs_Error
    
        ' these next three are validated already
'        FCWSharedOutputFile = fGetINISetting("Software\FireCallWin", "sharedOutputFile", FCWSettingsFile)
'        FCWExchangeFolder = fGetINISetting("Software\FireCallWin", "exchangeFolder", FCWSettingsFile)
'        FCWRefreshInterval = fGetINISetting("Software\FireCallWin", "refreshIntervalIndex", FCWSettingsFile)
        
        If FCWRefreshIntervalSecs = vbNullString Then FCWRefreshIntervalSecs = "20"
        If Val(FCWRefreshIntervalSecs) > 3600 Then FCWRefreshIntervalSecs = "3600"
        
        If FCWAlarmSound = vbNullString Then FCWAlarmSound = "G6AUC.wav"
        
        'General Config Items
        If FCWLoadBottom = vbNullString Then FCWLoadBottom = "1"
        If FCWEnableScrollbars = vbNullString Then FCWEnableScrollbars = "1"
        If FCWEnableTooltips = vbNullString Then FCWEnableTooltips = "1"
        If FCWEnableBalloonTooltips = vbNullString Then FCWEnableBalloonTooltips = "1"
        
        
        If FCWIconiseDelay = vbNullString Then FCWIconiseDelay = "20"

        If FCWEmojiSetIndex = vbNullString Then FCWEmojiSetIndex = "0"
        If FCWEmojiSetDesc = vbNullString Then FCWEmojiSetDesc = "standard"
        
        If FCWSendEmails = vbNullString Then FCWSendEmails = vbNullString  'sendEmails", FCWSettingsFile) '
        If FCWSendErrorEmails = vbNullString Then FCWSendErrorEmails = vbNullString
         
        'If FCWEmailAddress = "" Then FCWEmailAddress = "" 'emailAddress", FCWSettingsFile)
        If FCWAdviceInterval = vbNullString Then FCWAdviceInterval = vbNullString 'adviceInterval", FCWSettingsFile)

        If FCWAdviceIntervalSecs = vbNullString Then FCWAdviceIntervalSecs = "20"
        If Val(FCWAdviceIntervalSecs) > 172800 Then FCWAdviceIntervalSecs = "172800"

        If FCWLastEmail = vbNullString Then FCWLastEmail = "1970-01-01 00:00:01"
        If FCWLastHouseKeep = vbNullString Then FCWLastHouseKeep = "1970-01-01 00:00:01"
        

        If FCWMainFont = vbNullString Then FCWMainFont = "arial" 'textFont", FCWSettingsFile)
        If FCWMainFontSize = vbNullString Then FCWMainFontSize = "8" 'mainFontSize", FCWSettingsFile)
        If FCWMainFontItalics = vbNullString Then FCWMainFontItalics = False
        If FCWMainFontColour = vbNullString Then FCWMainFontColour = "0"
        
        
        If FCWPrefsFont = vbNullString Then FCWPrefsFont = "arial" 'prefsFont", FCWSettingsFile)
        If FCWPrefsFontSize = vbNullString Then FCWPrefsFontSize = "8" 'prefsFontSize", FCWSettingsFile)
        If FCWPrefsFontItalics = vbNullString Then FCWPrefsFontItalics = "false"
        If FCWPrefsFontColour = vbNullString Then FCWPrefsFontColour = "0"
        

        If FCWWindowLevel = vbNullString Then FCWWindowLevel = "0" 'WindowLevel", FCWSettingsFile)
        If FCWOpacity = vbNullString Then FCWOpacity = "100"
           
        If FCWMinimiseFormX = vbNullString Then FCWMinimiseFormX = "0"
        If FCWMinimiseFormY = vbNullString Then FCWMinimiseFormY = "0"
        'if FCWLastSoundPlayed = "" then 'fine
        
        If FCWLastSoundPlayed = vbNullString Then FCWLastSoundPlayed = "0"
        If FCWLastPingResponse = vbNullString Then FCWLastPingResponse = "0"
        If FCWLastAwakeString = vbNullString Then FCWLastAwakeString = "0"
        If FCWLastShutdown = vbNullString Then FCWLastShutdown = "0"
        If FCWAllowShutdowns = vbNullString Then FCWAllowShutdowns = "0"
        
        If FCWMaxLineLengthIndex = vbNullString Then FCWMaxLineLengthIndex = "5"
        'If FCWMaxLineLength = vbNullString Then FCWMaxLineLength = "96" ' this will occur in adjustPrefsControls
        
        If FCWClockStyle = vbNullString Then FCWClockStyle = "1"

        Call validateSmtpInputs
        
        If FCWRecipientEmail = vbNullString Then FCWRecipientEmail = "0"
        If FCWEmailSubject = vbNullString Then FCWEmailSubject = "0"
        If FCWEmailMessage = vbNullString Then FCWEmailMessage = "0"
        
        
        If FCWSingleListBox = vbNullString Then FCWSingleListBox = "0"
        
        If FCWImageDisplay = vbNullString Then FCWImageDisplay = "0"
        If FCWOptHandleData = vbNullString Then FCWOptHandleData = "0"

        If FCWAutomaticHousekeeping = vbNullString Then FCWAutomaticHousekeeping = "0"
        If FCWStartup = vbNullString Then FCWStartup = "0"

        If FCWArchiveDays = vbNullString Then FCWArchiveDays = "15"
        If FCWArchiveDaysIndex = vbNullString Then FCWArchiveDaysIndex = "0"
        
        
        If FCWBackupOnStart = vbNullString Then FCWBackupOnStart = "0"
        If FCWAutomaticBackups = vbNullString Then FCWAutomaticBackups = "0"
        If FCWAutomaticBackupInterval = vbNullString Then FCWAutomaticBackupInterval = "0"
        If FCWServiceProvider = vbNullString Then FCWServiceProvider = "0"
        If FCWCheckServiceProcesses = vbNullString Then FCWCheckServiceProcesses = "0"
        
        If FCWMsgBox13Enabled = vbNullString Then FCWMsgBox13Enabled = "1"
        
        If FCWCaptureDevices = vbNullString Then FCWCaptureDevices = "0"
        If FCWCaptureDevicesIndex = vbNullString Then FCWCaptureDevicesIndex = "0"
        If FCWRecordingQuality = vbNullString Then FCWRecordingQuality = "5"
        If Val(FCWRecordingQuality) > 5 Then FCWRecordingQuality = "5"
        If FCWLastSelectedTab = vbNullString Then FCWLastSelectedTab = "general"
        If FCWIconiseOpacity = vbNullString Then FCWIconiseOpacity = "True"
        ' check the boolean values are present, seems counter intuitive but it is correct
        If FCWIconiseOpacity <> "True" Then
            If FCWIconiseOpacity <> "False" Then FCWIconiseOpacity = "True"
        End If
        If FCWIconiseDesktop = vbNullString Then FCWIconiseDesktop = "True"
        ' check the boolean values are present
        If FCWIconiseDesktop <> "True" Then
            If FCWIconiseDesktop <> "False" Then FCWIconiseDesktop = "True"
        End If
        
        ' validate the archive folder name from the settings
        If FCWArchiveFolder = vbNullString Then ' if it is null
            If Not fDirExists(App.Path & "\archive") Then ' check to see if the default archive folder exists
                MkDir (App.Path & "\archive") ' if not, create it
            End If
            FCWArchiveFolder = App.Path & "\archive"
        Else ' if it has a value already set
            If IsValidPath(FCWArchiveFolder) Then ' check it for a valid path
                If Not fDirExists(FCWArchiveFolder) Then ' check to see if it does not exist
                    MkDir (FCWArchiveFolder) ' if not then create it
                Else
                    FCWArchiveFolder = FCWArchiveFolder
                End If
            Else
                If Not fDirExists(App.Path & "\archive") Then ' check to see if the default archive folder exists
                    MkDir (App.Path & "\archive") ' if not, create it
                End If
                FCWArchiveFolder = App.Path & "\archive"
            End If
        End If
        
                
        ' validate the archive folder name from the settings
        If FCWBackupFolder = vbNullString Then ' if it is null
            If Not fDirExists(App.Path & "\backup") Then ' check to see if the default backup folder exists
                MkDir (App.Path & "\backup") ' if not, create it
            End If
            FCWBackupFolder = App.Path & "\backup"
        Else ' if it has a value already set
            If IsValidPath(FCWBackupFolder) Then ' check it for a valid path
                If Not fDirExists(FCWBackupFolder) Then ' check to see if it does not exist
                    MkDir (FCWBackupFolder) ' if not then create it
                Else
                    FCWBackupFolder = FCWBackupFolder
                End If
            Else
                If Not fDirExists(App.Path & "\backup") Then ' check to see if the default backup folder exists
                    MkDir (App.Path & "\backup") ' if not, create it
                End If
                FCWBackupFolder = App.Path & "\backup"
            End If
        End If
        
   On Error GoTo 0
   Exit Sub

validateInputs_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure validateInputs of form fireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : getToolSettingsFile
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : get this tool's settings file and assign to a global var
'---------------------------------------------------------------------------------------
'
Private Sub getToolSettingsFile()
        
    ' variables declared
    
    'initialise the dimensioned variables
    FCWSettingsFile = vbNullString
    
    ' On Error GoTo getToolSettingsFile_Error
    If debugflg = 1 Then Debug.Print "%getToolSettingsFile"
    
    FCWSettingsDir = fSpecialFolder(feUserAppData) & "\FireCallWin" ' just for this user alone
    FCWSettingsFile = FCWSettingsDir & "\settings.ini"
        
    'if the folder does not exist then create the folder
    If Not fDirExists(FCWSettingsDir) Then
        MkDir FCWSettingsDir
    End If
    'if the Recordings folder does not exist then create the folder
    If Not fDirExists(FCWSettingsDir & "\Recordings") Then
        MkDir FCWSettingsDir & "\Recordings"
    End If
    
    'if the settings.ini does not exist then create the file by copying
    If Not fFExists(FCWSettingsFile) Then
        FileCopy App.Path & "\settings.ini", FCWSettingsFile
    End If
    
'    'confirm the settings file exists, if not use the version in the app itself
'    If Not fFExists(FCWSettingsFile) Then
'        toolSettingsFile = App.Path & "\settings.ini"
'    End If
    
   On Error GoTo 0
   Exit Sub

getToolSettingsFile_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure getToolSettingsFile of Form fireCallMain"

End Sub

' show the right click menu
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
        Me.PopupMenu mainMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

Form_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure Form_MouseDown of Form FireCallMain"

End Sub


' show the alternative right click menu and set the bulbs to dull
'---------------------------------------------------------------------------------------
' Procedure : lbxInputTextArea_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxInputTextArea_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo lbxInputTextArea_MouseDown_Error

    If Button = 2 Then
        mnuLBOpenSharedInputFile.Visible = True
        mnuLBOpenSharedOutputFile.Visible = False
        
        mnuInputCopyLine.Visible = True
        mnuInputQuoteLine.Visible = True
        mnuFindInput.Visible = True
        
        mnuOutputCopyLine.Visible = False
        mnuFindOutput.Visible = False
        mnuOutputPasteLine.Visible = False
        mnuOutputPasteAndGo.Visible = False
        mnuOutputEditLine.Visible = False
        mnuOutputDeleteLine.Visible = False
        
'        mnuCombinedDeleteLine.Visible = False
'        mnuCombinedEditLine.Visible = False
        mnuFindCombined.Visible = False
        mnuCombinedPasteLine.Visible = False
        mnuCombinedPasteAndGo.Visible = False
        mnuCombinedCopyLine.Visible = False
        mnuCombinedQuoteLine.Visible = False

        
        Me.PopupMenu listBoxMnuPopmenu, vbPopupMenuRightButton
    End If
    
    picTextChangeBright.Visible = False
    picTextChangeDull.Visible = True
    inputDataChangedFlag = False
    

   On Error GoTo 0
   Exit Sub

lbxInputTextArea_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxInputTextArea_MouseDown of Form FireCallMain"

End Sub


' show the alternative right click menu
'---------------------------------------------------------------------------------------
' Procedure : lbxOutputTextArea_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxOutputTextArea_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
    Dim theText As String
   On Error GoTo lbxOutputTextArea_MouseDown_Error

    If Button = 2 Then
        If lbxOutputTextArea.SelCount = 1 Then
            'a single line has been selected
            'theText = Left$(lbxOutputTextArea.List(lbxOutputTextArea.ListIndex), 15)
            
            theText = Left$(getCurrentLine(lbxOutputTextArea), 25)

            mnuOutputEditLine.Caption = "Edit The Line - """ & theText & """"
            mnuOutputEditLine.Visible = True
            mnuOutputDeleteLine.Visible = True
        Else
            'nothing or everything has selected
            mnuOutputEditLine.Visible = False
            mnuOutputDeleteLine.Visible = False
        End If

        mnuLBOpenSharedOutputFile.Visible = True
        mnuOutputCopyLine.Visible = True
        mnuFindOutput.Visible = True

        mnuLBOpenSharedInputFile.Visible = False
        mnuInputCopyLine.Visible = False
        mnuInputQuoteLine.Visible = False
        mnuFindInput.Visible = False
        mnuCombinedCopyLine.Visible = False
        mnuCombinedPasteLine.Visible = False
        mnuCombinedPasteAndGo.Visible = False
        mnuFindCombined.Visible = False
        mnuCombinedQuoteLine.Visible = False
        mnuCombinedEditLine.Visible = False
        mnuCombinedDeleteLine.Visible = False
        
        DoEvents
        If Clipboard.GetText <> "" Then
            mnuOutputPasteLine.Visible = True
            mnuOutputPasteAndGo.Visible = True
        Else
            mnuOutputPasteAndGo.Visible = False
            mnuOutputPasteLine.Visible = False
        End If

        Me.PopupMenu listBoxMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

lbxOutputTextArea_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lbxOutputTextArea_MouseDown of Form FireCallMain"

End Sub


'menu options follow

' about form display
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

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuAboutFireCallWin_Click of Form FireCallMain"
End Sub

' menu option to open the shared input file in an an editor or default application
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

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuOpenSharedInputFile_Click of Form FireCallMain"
End Sub

' menu option to open the shared output file in an an editor or default application
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

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuOpenSharedOutputFile_Click of Form FireCallMain"
End Sub

' menu option to open the shared folder in a file manager window
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

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuOpenSharedExchangeFolder_Click of Form FireCallMain"
End Sub
' menu option to open the shared folder in a file manager window
'---------------------------------------------------------------------------------------
' Procedure : mnuLBOpenSharedExchangeFolder_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLBOpenSharedExchangeFolder_Click()
   On Error GoTo mnuLBOpenSharedExchangeFolder_Click_Error

    Call ShellExecute(Me.hwnd, "Open", FCWExchangeFolder, vbNullString, App.Path, 1)

   On Error GoTo 0
   Exit Sub

mnuLBOpenSharedExchangeFolder_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuLBOpenSharedExchangeFolder_Click of Form FireCallMain"
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

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure mnuCoffee_Click of Form fireCallMain"
End Sub


' menu option to open the licence form
'---------------------------------------------------------------------------------------
' Procedure : mnuLicenceA_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLicenceA_Click()
   On Error GoTo mnuLicenceA_Click_Error

    Call LoadFileToTB(licence.txtLicenceTextBox, App.Path & "\licence.txt", False)
    licence.Show

   On Error GoTo 0
   Exit Sub

mnuLicenceA_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuLicenceA_Click of Form FireCallMain"

End Sub

' menu option to open the prefs form
'---------------------------------------------------------------------------------------
' Procedure : mnuProgramPreferences_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuProgramPreferences_Click()
    
   On Error GoTo mnuProgramPreferences_Click_Error

    Call makeConfigAvailable

   On Error GoTo 0
   Exit Sub

mnuProgramPreferences_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuProgramPreferences_Click of Form FireCallMain"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuSupport_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   : menu option to open the support page in default browser
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
        Call ShellExecute(Me.hwnd, "Open", "https://www.deviantart.com/yereverluvinuncleber/art/fireCallMain-VB6-Desktop-784624943", vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub

mnuSupport_Click_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure mnuSupport_Click of Form fireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuSweets_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   : menu option to open the Amazon donation page in default browser
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

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure mnuSweets_Click of Form fireCallMain"
End Sub

' a timer that reduces the opacity to zero then hides the main form
'---------------------------------------------------------------------------------------
' Procedure : opacityFadeOutTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub opacityFadeOutTimer_Timer()
   On Error GoTo opacityFadeOutTimer_Timer_Error
   
       FireCallMain.WindowState = vbMinimized

            opacitylevel = opacitylevel - 10
            If opacitylevel <= 0 Then
                opacitylevel = 0
                opacityFadeOutTimer.Enabled = False
                'FireCallMain.WindowState = vbMinimized
                'FireCallMain.Visible = False
            End If
            
            Call setOpacity(opacitylevel)

   On Error GoTo 0
   Exit Sub

opacityFadeOutTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure opacityFadeOutTimer_Timer of Form FireCallMain"
End Sub
' a timer that makes the main form visible, then increases the opacity to full
'---------------------------------------------------------------------------------------
' Procedure : opacityFadeInTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub opacityFadeInTimer_Timer()
            
   On Error GoTo opacityFadeInTimer_Timer_Error

            If opacitylevel <= 1 Then ' as soon as the form opacity starts to be not 0 then the form is made visible
                'FireCallMain.Visible = True
                FireCallMain.WindowState = vbNormal
                FireCallMain.txtTextEntry.SetFocus
            End If
            
            opacitylevel = opacitylevel + 10
            
            If opacitylevel >= 255 Then
                opacitylevel = 255
                opacityFadeInTimer.Enabled = False
            End If
            
            Call setOpacity(opacitylevel)

   On Error GoTo 0
   Exit Sub

opacityFadeInTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure opacityFadeInTimer_Timer of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : opacityToTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub opacityToTimer_Timer()
    Dim finalOpacitylevel As Integer
    
   On Error GoTo opacityToTimer_Timer_Error

    opacitylevel = opacitylevel - 10
    
    finalOpacitylevel = 255 * (Val(FCWOpacity) / 100)
    
    If opacitylevel <= finalOpacitylevel Then
        opacitylevel = finalOpacitylevel
        opacityToTimer.Enabled = False
    End If

    Call setOpacity(opacitylevel)

   On Error GoTo 0
   Exit Sub

opacityToTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure opacityToTimer_Timer of Form FireCallMain"
End Sub

' hides the vertical scrollbar
'---------------------------------------------------------------------------------------
' Procedure : outputScrollBarTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub outputScrollBarTimer_Timer()
   On Error GoTo outputScrollBarTimer_Timer_Error

    Call ShowScrollBar(lbxOutputTextArea.hwnd, SB_VERT, False)

   On Error GoTo 0
   Exit Sub

outputScrollBarTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure outputScrollBarTimer_Timer of Form FireCallMain"
End Sub

' hides the combined scrollbar
'---------------------------------------------------------------------------------------
' Procedure : combinedScrollBarTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub combinedScrollBarTimer_Timer()
   On Error GoTo combinedScrollBarTimer_Timer_Error

    Call ShowScrollBar(lbxCombinedTextArea.hwnd, SB_VERT, False)

   On Error GoTo 0
   Exit Sub

combinedScrollBarTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure combinedScrollBarTimer_Timer of Form FireCallMain"
End Sub
' play a sound and pause the timer
'---------------------------------------------------------------------------------------
' Procedure : pausePrinterTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub pausePrinterTimer_Timer()
    Dim soundtoplay As String
    
   On Error GoTo pausePrinterTimer_Timer_Error

    dropTimerCount = dropTimerCount + 1
    
    If dropTimerCount = 10 Then
        If FCWPlayVolume = "1" Then
            soundtoplay = App.Path & "\Resources\Sounds\" & "page-fumble.wav"
        Else
            soundtoplay = App.Path & "\Resources\Sounds\" & "page-fumbleQuiet.wav"
        End If
        
        If fFExists(soundtoplay) And btnLid.Visible = False And FCWEnableSounds = "1" Then PlaySound soundtoplay, ByVal 0&, SND_FILENAME Or SND_ASYNC
        dropTimerCount = 0
        pausePrinterTimer.Enabled = False
        dropTimer.Enabled = True
    End If

   On Error GoTo 0
   Exit Sub

pausePrinterTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure pausePrinterTimer_Timer of Form FireCallMain"

End Sub

' make the Emoji lid disappear or show the right click menu
'---------------------------------------------------------------------------------------
' Procedure : picBtnLidCatch_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picBtnLidCatch_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
    
   On Error GoTo picBtnLidCatch_MouseDown_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    
    If Button = 2 Then
        Me.PopupMenu mainMnuPopmenu, vbPopupMenuRightButton
    End If
    
    If Button = 1 Then
        btnLid.Visible = False
        picBtnLidCatch.Visible = False
        picBtnLidShadow.Visible = False
        picLidOpen.Visible = True
        txtTextEntry.SetFocus
    End If

   On Error GoTo 0
   Exit Sub

picBtnLidCatch_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picBtnLidCatch_MouseDown of Form FireCallMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : picBtnLidCatch_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picBtnLidCatch_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picBtnLidCatch_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip picBtnLidCatch.hwnd, "Click on pull catch to remove the cover and display the Emoji Panel below.", _
                  TTIconInfo, "Help on Opening the Lid", , , , True

   On Error GoTo 0
   Exit Sub

picBtnLidCatch_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picBtnLidCatch_MouseMove of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picBtnPlaySound_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picBtnPlaySound_Click()
    
    ' Play
    Dim cmd As String
    Dim ret As Long
    Dim soundFileName As String
    Dim fileSize As Long
    Dim playUsingDefaultApp As Boolean
    
   On Error GoTo picBtnPlaySound_Click_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    
    playingTimerMax = 0
    soundFileName = recordingFilePath
    playUsingDefaultApp = False
    
    If recordingIsPossible = False Then Exit Sub
    
    If recording = True Then Exit Sub ' exit immediately if a recording is taking place
    If toolTipFlag = True Then btnStop.ToolTipText = "Stop Playing"
    
    Call btnStop_Click ' an extra click on the stop button, just in case
    
    playing = True ' set the global playing flag
    
    soundFileName = fGetFileNameFromPath(recordingFilePath)
    If IsNumeric(Mid$(soundFileName, 1, 2)) Then
        playingTimerMax = Val(Mid$(soundFileName, 1, 2)) + 1
    Else ' we cannot determine the length of other sound types as VB6 has no function to do so for so many file types
        playingTimerMax = 1
        playUsingDefaultApp = True
    End If
    
    If fExtractSuffixWithDot(recordingFilePath) = ".m4a" Then playUsingDefaultApp = True
    
    PlayTimer.Enabled = True

    picPlayLampDull.Visible = False
    picPlayLampBright.Visible = True
    
    If fFExists(recordingFilePath) Then
        If playUsingDefaultApp = True Then
            Call ShellExecute(Me.hwnd, "Open", recordingFilePath, vbNullString, App.Path, 1)
        Else
            PlaySound recordingFilePath, ByVal 0&, SND_FILENAME Or SND_ASYNC ' just a wav file for which we have the known length
        End If
    End If

   On Error GoTo 0
   Exit Sub

picBtnPlaySound_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picBtnPlaySound_Click of Form FireCallMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : picBtnPlaySound_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picBtnPlaySound_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

   On Error GoTo picBtnPlaySound_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip picBtnPlaySound.hwnd, "This button is generally hidden but when you have selected a recording to play, the green button will appear. When playing, the green lamp will light up brightly but will change from bright green to dull when it has finished.", _
                  TTIconInfo, "Help on the Buzzer Lamp", , , , True

   On Error GoTo 0
   Exit Sub

picBtnPlaySound_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picBtnPlaySound_MouseMove of Form FireCallMain"
End Sub

'Private Sub picGreenButtonHole_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    If FCWEnableBalloonTooltips = "1" Then CreateToolTip picGreenButtonHole.hWnd, "This hole is the location for the play button, which is generally hidden until you have selected a recording to play.", _
'                  TTIconInfo, "Help on the Green Button Hole", , , , True
'End Sub

' make the buzzer indicator dull after it has been raised
'---------------------------------------------------------------------------------------
' Procedure : picBuzzerBrightLamp_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picBuzzerBrightLamp_Click()
   On Error GoTo picBuzzerBrightLamp_Click_Error

        picBuzzerDullLamp.Visible = True
        picBuzzerBrightLamp.Visible = False
        txtTextEntry.SetFocus

   On Error GoTo 0
   Exit Sub

picBuzzerBrightLamp_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picBuzzerBrightLamp_Click of Form FireCallMain"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : picBuzzerDullLamp_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picBuzzerDullLamp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picBuzzerDullLamp_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip picBuzzerDullLamp.hwnd, "Just above the Clock or the Fire Call button is the buzzer lamp. If your chat partner has buzzed you during your absence, meaning that you did not hear the buzz, the buzz light will stay lit to let you know you've been buzzed. A click on the buzz lamp will extinguish the lamp to its dull state", _
                  TTIconInfo, "Help on the Buzzer Lamp", , , , True

   On Error GoTo 0
   Exit Sub

picBuzzerDullLamp_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picBuzzerDullLamp_MouseMove of Form FireCallMain"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : picBuzzerBrightLamp_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picBuzzerBrightLamp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picBuzzerBrightLamp_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip picBuzzerBrightLamp.hwnd, "Just above the Clock or the Fire Call button is the buzzer lamp. If your chat partner has buzzed you during your absence, meaning that you did not hear the buzz, the buzz light will stay lit to let you know you've been buzzed. A click on the buzz lamp will extinguish the lamp to its dull state", _
                  TTIconInfo, "Help on the Buzzer Lamp", , , , True

   On Error GoTo 0
   Exit Sub

picBuzzerBrightLamp_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picBuzzerBrightLamp_MouseMove of Form FireCallMain"
End Sub

' show the right click menu on the clock
'---------------------------------------------------------------------------------------
' Procedure : picClock_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picClock_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo picClock_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu ClockMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

picClock_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picClock_MouseDown of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picClock_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picClock_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picClock_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip picClock.hwnd, "The clock can be replaced by the buzzer button, just click the screw top left.", _
                  TTIconInfo, "Help on the Clock", , , , True

   On Error GoTo 0
   Exit Sub

picClock_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picClock_MouseMove of Form FireCallMain"

End Sub

' toggle the clock and button, saving the result for the next restart
'---------------------------------------------------------------------------------------
' Procedure : picClockSwitch_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picClockSwitch_Click()

   On Error GoTo picClockSwitch_Click_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    If picRedButton.Visible = False Then
        picRedButton.Visible = True
        picClock.Visible = False
        FCWClockStyle = "RedButton"
    Else
        picRedButton.Visible = False
        picClock.Visible = True
        FCWClockStyle = "VB6Clock"
    End If
    
    If fFExists(FCWSettingsFile) Then
        PutINISetting "Software\FireCallWin", "clockStyle", FCWClockStyle, FCWSettingsFile
    End If
    txtTextEntry.SetFocus

   On Error GoTo 0
   Exit Sub

picClockSwitch_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picClockSwitch_Click of Form FireCallMain"
End Sub

' dbl clicking the button also switches to clock but not vice versa
'---------------------------------------------------------------------------------------
' Procedure : picClock_DblClick
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picClock_DblClick()
    
   On Error GoTo picClock_DblClick_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    
    picRedButton.Visible = True
    picClock.Visible = False
    txtTextEntry.SetFocus

   On Error GoTo 0
   Exit Sub

picClock_DblClick_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picClock_DblClick of Form FireCallMain"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : picClockSwitch_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picClockSwitch_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picClockSwitch_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip picClockSwitch.hwnd, "Pressing the top left screw on the Fire Call button will cause the button to flip to the real-time clock display. A further press on the screw will revert back to the fire call button, ready to buzz!", _
                  TTIconInfo, "Help on Toggling Clock and Buzzer", , , , True

   On Error GoTo 0
   Exit Sub

picClockSwitch_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picClockSwitch_MouseMove of Form FireCallMain"
End Sub

' closes the image display with a sound
'---------------------------------------------------------------------------------------
' Procedure : picCloseMe_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picCloseMe_Click()
    Dim soundtoplay As String
    Dim imgFilePath As String
    
   On Error GoTo picCloseMe_Click_Error

    imgFilePath = App.Path & "\Resources\images\lidBackgroundDull.jpg"
    If fFExists(imgFilePath) Then
        picLidBackground.Picture = LoadPicture(imgFilePath)
    End If
    
    FireCallMain.picEmojiKnobRight.Visible = True

    If FCWPlayVolume = "1" Then
        soundtoplay = App.Path & "\Resources\Sounds\" & "page-fumble.wav"
    Else
        soundtoplay = App.Path & "\Resources\Sounds\" & "page-fumbleQuiet.wav"
    End If
    
    If fFExists(soundtoplay) And btnLid.Visible = False And FCWEnableSounds = "1" Then PlaySound soundtoplay, ByVal 0&, SND_FILENAME Or SND_ASYNC

    picImagePrintOut.Visible = False
    'picPrintOutShadow.Visible = False
    
    FCWImageDisplay = "0"
    PutINISetting "Software\FireCallWin", "imageDisplay", FCWImageDisplay, FCWSettingsFile
    
    
    txtTextEntry.SetFocus

   On Error GoTo 0
   Exit Sub

picCloseMe_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picCloseMe_Click of Form FireCallMain"
    
End Sub
' does some animation and sounds when the emoji is clicked upon
'---------------------------------------------------------------------------------------
' Procedure : picEmoji_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picEmoji_Click()
   On Error GoTo picEmoji_Click_Error

    Call clickOnPicEmoji

   On Error GoTo 0
   Exit Sub

picEmoji_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picEmoji_Click of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picEmoji_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picEmoji_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picEmoji_MouseMove_Error

        If FCWEnableBalloonTooltips = "1" Then CreateToolTip picEmoji.hwnd, "Clicking on the strip of paper (just appearing at the top of the Emoji Window) will allow you to see your chat partner's current Emoji state. A small animation will run when the paper is first clicked upon. Another click on the displayed print out will shred it.", _
                  TTIconInfo, "Help on the Partner's Emoji State", , , , True

   On Error GoTo 0
   Exit Sub

picEmoji_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picEmoji_MouseMove of Form FireCallMain"

End Sub

' select the previous emoji
'---------------------------------------------------------------------------------------
' Procedure : picEmojiKnobLeft_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picEmojiKnobLeft_Click()
    Dim fullPath As String
    Dim soundtoplay As String

   On Error GoTo picEmojiKnobLeft_Click_Error

    If currindex = 0 Then currindex = cmbEmojiSelection.ListIndex
    currindex = currindex - 1
    If currindex < 1 Then currindex = 1
    
    cmbEmojiSelection.ListIndex = currindex
    fullPath = App.Path & "\resources\Emojis\" & FCWEmojiSetDesc & "\telly\" & cmbEmojiSelection.List(currindex)
   
   
    If FCWPlayVolume = "1" Then
        soundtoplay = App.Path & "\Resources\Sounds\" & "keypress.wav"
    Else
        soundtoplay = App.Path & "\Resources\Sounds\" & "keypress.wav"
    End If

    If fFExists(soundtoplay) And btnLid.Visible = False Then
        If FCWEnableSounds = "1" Then PlaySound soundtoplay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
   
    If fFExists(fullPath) Then
        picOutputEmoji.Picture = LoadPicture(fullPath)
    End If
    
    
    txtTextEntry.SetFocus

   On Error GoTo 0
   Exit Sub

picEmojiKnobLeft_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picEmojiKnobLeft_Click of Form FireCallMain"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picEmojiKnobLeft_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picEmojiKnobLeft_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picEmojiKnobLeft_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip picEmojiKnobLeft.hwnd, "The left hand button allows you to select other emojis for viewing on the telly screen. Note that as you select each Emoji, the Emoji drop-down at the top of the program will change as well.", _
                  TTIconInfo, "Help on the Emoji Selection Controls", , , , True

   On Error GoTo 0
   Exit Sub

picEmojiKnobLeft_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picEmojiKnobLeft_MouseMove of Form FireCallMain"
End Sub

' select the next emoji
'---------------------------------------------------------------------------------------
' Procedure : picEmojiKnobRight_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picEmojiKnobRight_Click()
    Dim fullPath As String
    Dim soundtoplay As String

   On Error GoTo picEmojiKnobRight_Click_Error

    If currindex = 0 Then currindex = cmbEmojiSelection.ListIndex
    currindex = currindex + 1
    If currindex > cmbEmojiSelection.ListCount Then currindex = cmbEmojiSelection.ListCount
    
    cmbEmojiSelection.ListIndex = currindex
    fullPath = App.Path & "\resources\Emojis\" & FCWEmojiSetDesc & "\telly\" & cmbEmojiSelection.List(currindex)
    
    If FCWPlayVolume = "1" Then
        soundtoplay = App.Path & "\Resources\Sounds\" & "keypress.wav"
    Else
        soundtoplay = App.Path & "\Resources\Sounds\" & "keypress.wav"
    End If

    If fFExists(soundtoplay) And btnLid.Visible = False Then
        If FCWEnableSounds = "1" Then PlaySound soundtoplay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    If fFExists(fullPath) Then
        picOutputEmoji.Picture = LoadPicture(fullPath)
    End If
    
    txtTextEntry.SetFocus


   On Error GoTo 0
   Exit Sub

picEmojiKnobRight_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picEmojiKnobRight_Click of Form FireCallMain"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : picEmojiKnobRight_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picEmojiKnobRight_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picEmojiKnobRight_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip picEmojiKnobRight.hwnd, "The right hand button allows you to select other emojis for viewing on the telly screen. Note that as you select each Emoji, the Emoji drop-down at the top of the program will change as well.", _
                  TTIconInfo, "Help on the Emoji Selection Controls", , , , True

   On Error GoTo 0
   Exit Sub

picEmojiKnobRight_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picEmojiKnobRight_MouseMove of Form FireCallMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : picEmojiSmall_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picEmojiSmall_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picEmojiSmall_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip picEmojiSmall.hwnd, "The Emoji selection Drop Down. Select an Emoji and press SEND. This will send the emoji to your chat partner. If you open the Emoji panel, bottom right you can see your current Emoji state.", _
                  TTIconInfo, "Help on Emojis", , , , True

   On Error GoTo 0
   Exit Sub

picEmojiSmall_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picEmojiSmall_MouseMove of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picFsoLampBright_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picFsoLampBright_Click()
   On Error GoTo picFsoLampBright_Click_Error

        picFsoLid.Visible = True

   On Error GoTo 0
   Exit Sub

picFsoLampBright_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picFsoLampBright_Click of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picFsoLampDull_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picFsoLampDull_Click()
   On Error GoTo picFsoLampDull_Click_Error

    picFsoLid.Visible = True

   On Error GoTo 0
   Exit Sub

picFsoLampDull_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picFsoLampDull_Click of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picFsoLid_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picFsoLid_Click()
   On Error GoTo picFsoLid_Click_Error

    picFsoLid.Visible = False
    
'    MsgBox " borderSizeLeft=" & borderSizeLeft & " borderSizeRight=" & borderSizeRight & vbCr & _
'            " borderSizeTop=" & borderSizeTop & " borderSizeBottom=" & borderSizeBottom

   On Error GoTo 0
   Exit Sub

picFsoLid_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picFsoLid_Click of Form FireCallMain"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picImageButton_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picImageButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picImageButton_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip picImageButton.hwnd, "This button top right will close any image or icon that is currently being displayed .", _
                  TTIconInfo, "Help on the Image Control", , , , True

   On Error GoTo 0
   Exit Sub

picImageButton_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picImageButton_MouseMove of Form FireCallMain"

End Sub

' open the currently displayed image using default application
'---------------------------------------------------------------------------------------
' Procedure : picImagePrintOut_DblClick
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picImagePrintOut_DblClick()
    ' variables declared
    Dim suffix As String
    Dim answer As VbMsgBoxResult
    Dim attachmentFilename As String
    Dim execStatus As Long
    
    'initialise the dimensioned variables
   On Error GoTo picImagePrintOut_DblClick_Error

    answer = vbNo
    execStatus = 0
    
    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    
    If fFExists(RTrim$(displayedAttachmentFilePath)) Then
        If foundRecording = True Then
            Call picBtnPlaySound_Click
            'PlaySound attachmentFilePath, ByVal 0&, SND_FILENAME Or SND_ASYNC
        Else
            suffix = fExtractSuffixWithDot(attachmentFilePath)
            If fInstrSuffix(executableSuffixArrayList, LCase(suffix)) Then
                'picImagePrintOut.ToolTipText = attachmentFilename & " This file is missing - it is no longer in the dropbox shared folder."

                answer = MsgBox(attachmentFilePath & vbCrLf & vbCrLf & " This is an executable program, running it could be dangerous and unpredictable things may happen." & vbCrLf & vbCrLf & "Are you sure you wish to proceed?", vbExclamation + vbYesNo)
                If answer = vbYes Then
                    attachmentFilename = fGetFileNameFromPath(attachmentFilePath)
                    If attachmentFilename = "FireCallWin.exe" Then
                        answer = MsgBox(attachmentFilePath & vbCrLf & vbCrLf & " This is the FireCallWin program, it cannot run itself again.", vbExclamation)
                    Else
                        execStatus = ShellExecute(Me.hwnd, "Open", displayedAttachmentFilePath, vbNullString, App.Path, 1)
                        If execStatus <= 32 Then MsgBox "No association found for " & suffix & " file type, FireCall cannot run this file type."
                   End If
                End If

            Else
                execStatus = ShellExecute(Me.hwnd, "Open", displayedAttachmentFilePath, vbNullString, App.Path, 1)
                If execStatus <= 32 Then MsgBox "No association found for " & suffix & " file type, FireCall cannot open it. " & vbCrLf & "You need to create an association for this file type in Windows. "
            End If
        End If
    Else
        If fDirExists(attachmentFilePath) Then ' we've checked file existence, now folder.
            execStatus = ShellExecute(Me.hwnd, "Open", displayedAttachmentFilePath, vbNullString, App.Path, 1)
            If execStatus <= 32 Then MsgBox "Attempt to open folder failed."
        Else
            MsgBox "%Err-I-ErrorNumber 10 - File not found, if a recent attachment, Dropbox is possibly still transferring." & vbCrLf & _
            "If an older attachment, the image may have been deleted from the exchange folder"
        End If
    End If

    txtTextEntry.SetFocus

   On Error GoTo 0
   Exit Sub

picImagePrintOut_DblClick_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picImagePrintOut_DblClick of Form FireCallMain"

End Sub





'---------------------------------------------------------------------------------------
' Procedure : picImagePrintOut_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picImagePrintOut_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picImagePrintOut_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu picMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

picImagePrintOut_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picImagePrintOut_MouseDown of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picImagePrintOut_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picImagePrintOut_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picImagePrintOut_MouseMove_Error

        If binaryFlag = True Then
        
            CreateToolTip picImagePrintOut.hwnd, "This is an executable program, double clicking on it will run it straight away, this could be dangerous", _
                  TTIconInfo, "Executable file detected", , , , True
        Else
            If FCWEnableBalloonTooltips = "1" Then
                CreateToolTip picImagePrintOut.hwnd, "When your chat partner sends you an image or other attachment, it will appear in a picture box in the emoji control panel. If the file is a known image then it will display the image itself. If it is a 'known' image format but Windows cannot easily display it then an appropriate icon will appear in its stead. The program will attempt to display the following image types - .jpg,.jpeg,.bmp,.ico,.png,.tif,.tiff,.gif,.cur,.wmf,.emf. If it is not an image but a file then a double click on the icon picture box will open the attachment using the default application. In the above case Windows will open the attached EXE file and run it, so take care! If an audio file is selected Windows media player will open and play the WAV file. What happens is down to your Windows configuration and the default application types.", _
                  TTIconInfo, "Help on the Image and Icon Display", , , , True
            Else
                Call DestroyToolTip
            End If
                  
        End If

   On Error GoTo 0
   Exit Sub

picImagePrintOut_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picImagePrintOut_MouseMove of Form FireCallMain"
                                  
End Sub

'Private Sub picIoMethodDull_Click()
'    ioMethodADO = True
'    picIoMethodBright.Visible = True
'    picIoMethodDull.Visible = False
'
'    If lbxOutputTextArea.Visible = True Then lbxOutputTextArea.Clear
'    If lbxInputTextArea.Visible = True Then lbxInputTextArea.Clear
'    If lbxCombinedTextArea.Visible = True Then lbxCombinedTextArea.Clear
'    Call btnRefresh_Click
'
'End Sub
'Private Sub picIoMethodBright_Click()
'    ioMethodADO = False
'    picIoMethodBright.Visible = False
'    picIoMethodDull.Visible = True
'
'    If lbxOutputTextArea.Visible = True Then lbxOutputTextArea.Clear
'    If lbxInputTextArea.Visible = True Then lbxInputTextArea.Clear
'    If lbxCombinedTextArea.Visible = True Then lbxCombinedTextArea.Clear
'    Call btnRefresh_Click
'End Sub

'if the background of the emoji area is lit (bright) then make it show the dull version
'---------------------------------------------------------------------------------------
' Procedure : picLidBackground_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picLidBackground_Click()
    
    Dim fullPath As String
   On Error GoTo picLidBackground_Click_Error

    btnLid.Visible = True
    picBtnLidCatch.Visible = True
    picBtnLidShadow.Visible = True
    picLidOpen.Visible = False
    fullPath = App.Path & "\resources\images\" & "lidBackgroundDull.jpg"
            
    If fFExists(fullPath) Then
        picLidBackground.Picture = LoadPicture(fullPath)
    End If
    txtTextEntry.SetFocus

   On Error GoTo 0
   Exit Sub

picLidBackground_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picLidBackground_Click of Form FireCallMain"
End Sub

' a small button that displays the picture image
'---------------------------------------------------------------------------------------
' Procedure : picImageButton_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picImageButton_Click()
    
    Dim soundtoplay As String
    Dim imgFilePath As String
    
    'If FCWImageDisplay = "1" Then
   On Error GoTo picImageButton_Click_Error

        imgFilePath = App.Path & "\Resources\images\lidBackgroundDullShadowed.jpg"
        If fFExists(imgFilePath) Then
            picLidBackground.Picture = LoadPicture(imgFilePath)
        End If
    'End If
    
    picImagePrintOut.Visible = True
    'picPrintOutShadow.Visible = True
    
    picEmojiKnobRight.Visible = False
    
    If FCWPlayVolume = "1" Then
        soundtoplay = App.Path & "\Resources\Sounds\" & "page-fumble.wav"
    Else
        soundtoplay = App.Path & "\Resources\Sounds\" & "page-fumbleQuiet.wav"
    End If
    
    If fFExists(soundtoplay) And btnLid.Visible = False And FCWEnableSounds = "1" Then PlaySound soundtoplay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    
    FCWImageDisplay = "1"
    PutINISetting "Software\FireCallWin", "imageDisplay", FCWImageDisplay, FCWSettingsFile
    
    txtTextEntry.SetFocus

   On Error GoTo 0
   Exit Sub

picImageButton_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picImageButton_Click of Form FireCallMain"
End Sub

' right click popup menu for the lid background
'---------------------------------------------------------------------------------------
' Procedure : picLidBackground_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picLidBackground_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo picLidBackground_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mainMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

picLidBackground_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picLidBackground_MouseDown of Form FireCallMain"
End Sub






'---------------------------------------------------------------------------------------
' Procedure : picLidBackground_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picLidBackground_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picLidBackground_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip picLidBackground.hwnd, "Click the control panel background to close the lid on the Emoji Control Panel.", _
                  TTIconInfo, "Help on the Emoji Control Panel Background", , , , True

   On Error GoTo 0
   Exit Sub

picLidBackground_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picLidBackground_MouseMove of Form FireCallMain"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : picLidOpen_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picLidOpen_Click()

   On Error GoTo picLidOpen_Click_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    
    picLidOpen.Visible = False
    btnLid.Visible = True
    picBtnLidCatch.Visible = True
    picBtnLidShadow.Visible = True

   On Error GoTo 0
   Exit Sub

picLidOpen_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picLidOpen_Click of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picLidOpen_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picLidOpen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picLidOpen_MouseMove_Error

        If FCWEnableBalloonTooltips = "1" Then CreateToolTip picLidOpen.hwnd, "Click upon the lid to close the Emoji Control Panel and view the Audio Recording Tools.", _
                  TTIconInfo, "Help on Closing the Lid", , , , True

   On Error GoTo 0
   Exit Sub

picLidOpen_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picLidOpen_MouseMove of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picOutputEmoji_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picOutputEmoji_Click()
   On Error GoTo picOutputEmoji_Click_Error

    Me.Refresh

   On Error GoTo 0
   Exit Sub

picOutputEmoji_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picOutputEmoji_Click of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picOutputEmoji_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picOutputEmoji_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picOutputEmoji_MouseMove_Error

        If FCWEnableBalloonTooltips = "1" Then CreateToolTip picOutputEmoji.hwnd, "The Emoji Control Panel shows your current emoji state on the television screen.", _
                  TTIconInfo, "Help on the Partner's Emoji State", , , , True

   On Error GoTo 0
   Exit Sub

picOutputEmoji_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picOutputEmoji_MouseMove of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picPlayLampDull_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picPlayLampDull_Click()
   On Error GoTo picPlayLampDull_Click_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity

   On Error GoTo 0
   Exit Sub

picPlayLampDull_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picPlayLampDull_Click of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picPlayLampDull_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picPlayLampDull_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picPlayLampDull_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip picPlayLampDull.hwnd, "When you are playing a recording the lamp will light up brightly but will change from bright green to dull when it has finished.", _
                  TTIconInfo, "Help on the Recording Lamp", , , , True

   On Error GoTo 0
   Exit Sub

picPlayLampDull_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picPlayLampDull_MouseMove of Form FireCallMain"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : picPlayLampBright_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picPlayLampBright_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picPlayLampBright_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip picPlayLampBright.hwnd, "When you are playing a recording the lamp will light up brightly but will change from bright green to dull when it has finished.", _
                  TTIconInfo, "Help on the Recording Lamp", , , , True

   On Error GoTo 0
   Exit Sub

picPlayLampBright_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picPlayLampBright_MouseMove of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picRecordLampBright_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picRecordLampBright_Click()
   On Error GoTo picRecordLampBright_Click_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity

   On Error GoTo 0
   Exit Sub

picRecordLampBright_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picRecordLampBright_Click of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picRecordLampBright_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picRecordLampBright_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picRecordLampBright_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip picRecordLampBright.hwnd, "The small red lamp on the speaker panel will light up brightly when recording a message for your chat partner. The maximum length is 65 seconds.", _
                  TTIconInfo, "Help on the Recording Lamp", , , , True

   On Error GoTo 0
   Exit Sub

picRecordLampBright_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picRecordLampBright_MouseMove of Form FireCallMain"
  
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picRecordLampDull_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picRecordLampDull_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picRecordLampDull_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip picRecordLampDull.hwnd, "The small red lamp on the speaker panel will light up brightly when recording a message for your chat partner. The maximum length is 65 seconds.", _
                  TTIconInfo, "Help on the Recording Lamp", , , , True

   On Error GoTo 0
   Exit Sub

picRecordLampDull_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picRecordLampDull_MouseMove of Form FireCallMain"
End Sub

' right click popup menu for the lid background
'---------------------------------------------------------------------------------------
' Procedure : picRedButton_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picRedButton_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
    Dim fullPath As String
    
   On Error GoTo picRedButton_MouseDown_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    
    If Button = 2 Then
        Me.PopupMenu mainMnuPopmenu, vbPopupMenuRightButton
    End If
    
    If Button = 1 Then
        buzzerTimer.Enabled = True

        fullPath = App.Path & "\resources\images\" & "redButtonPressed" & ".jpg"

        If fFExists(fullPath) Then
            picRedButton.Picture = LoadPicture(fullPath)
        End If
    End If

   On Error GoTo 0
   Exit Sub

picRedButton_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picRedButton_MouseDown of Form FireCallMain"
End Sub
' timer to dull the bright the lamp after 5 seconds of being lit
'---------------------------------------------------------------------------------------
' Procedure : lampTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lampTimer_Timer()
   On Error GoTo lampTimer_Timer_Error

    picTimerLampBright.Visible = False
    picTimerLampDull.Visible = True
    lampTimer.Enabled = False

   On Error GoTo 0
   Exit Sub

lampTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure lampTimer_Timer of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picRedButton_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picRedButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picRedButton_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip picRedButton.hwnd, "This big red button, if kept pressed for 5 seconds will buzz your chat partner to get his attention. You will also hear the buzzer sound at your end to confirm the operation.", _
                  TTIconInfo, "Help on the Clock", , , , True

   On Error GoTo 0
   Exit Sub

picRedButton_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picRedButton_MouseMove of Form FireCallMain"
End Sub


'Restore the normal image of the big red button
'---------------------------------------------------------------------------------------
' Procedure : picRedButton_MouseUp
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picRedButton_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)

    Dim fullPath As String
    
   On Error GoTo picRedButton_MouseUp_Error

    buzzerTimer.Enabled = False
    If Button = 1 Then
        fullPath = App.Path & "\resources\images\" & "redButton" & ".jpg"
                
        If fFExists(fullPath) Then
            picRedButton.Picture = LoadPicture(fullPath)
        End If
    End If
    txtTextEntry.SetFocus ' set focus back to the text entry box

   On Error GoTo 0
   Exit Sub

picRedButton_MouseUp_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picRedButton_MouseUp of Form FireCallMain"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : picSideBar_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picSideBar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picSideBar_MouseDown_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mainMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

picSideBar_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picSideBar_MouseDown of Form FireCallMain"
End Sub



' clicking on the speaker grilles controls the volume
'---------------------------------------------------------------------------------------
' Procedure : picSpeakerGrille_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picSpeakerGrille_Click()
   On Error GoTo picSpeakerGrille_Click_Error

    picSpeakerGrille.Visible = False
    picSpeakerGrilleOpen.Visible = True
    FCWPlayVolume = "1"
    
    If fFExists(FCWSettingsFile) Then
        PutINISetting "Software\FireCallWin", "playVolume", FCWPlayVolume, FCWSettingsFile
    End If

   On Error GoTo 0
   Exit Sub

picSpeakerGrille_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picSpeakerGrille_Click of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picSpeakerGrille_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picSpeakerGrille_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picSpeakerGrille_MouseMove_Error

        If FCWEnableBalloonTooltips = "1" Then CreateToolTip picSpeakerGrille.hwnd, "Clicking on the speaker grille will toggle the sound partial mute for the whole application, changing them from loud to quiet and back again as required.", _
                  TTIconInfo, "Help on the Sound Mute", , , , True

   On Error GoTo 0
   Exit Sub

picSpeakerGrille_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picSpeakerGrille_MouseMove of Form FireCallMain"
End Sub

' clicking on the speaker grilles controls the volume
'---------------------------------------------------------------------------------------
' Procedure : picSpeakerGrilleOpen_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picSpeakerGrilleOpen_Click()
   On Error GoTo picSpeakerGrilleOpen_Click_Error

    picSpeakerGrille.Visible = True
    picSpeakerGrilleOpen.Visible = False
    FCWPlayVolume = "0"
    
    If fFExists(FCWSettingsFile) Then
        PutINISetting "Software\FireCallWin", "playVolume", FCWPlayVolume, FCWSettingsFile
    End If

   On Error GoTo 0
   Exit Sub

picSpeakerGrilleOpen_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picSpeakerGrilleOpen_Click of Form FireCallMain"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : picSpeakerGrilleOpen_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picSpeakerGrilleOpen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picSpeakerGrilleOpen_MouseMove_Error

        If FCWEnableBalloonTooltips = "1" Then CreateToolTip picSpeakerGrilleOpen.hwnd, "Clicking on the speaker grille will toggle the sound partial mute for the whole application, changing them from loud to quiet and back again as required.", _
                  TTIconInfo, "Help on the Sound Mute", , , , True

   On Error GoTo 0
   Exit Sub

picSpeakerGrilleOpen_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picSpeakerGrilleOpen_MouseMove of Form FireCallMain"
End Sub

' dull the bright the lamp when clicked
'---------------------------------------------------------------------------------------
' Procedure : picTextChangeBright_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picTextChangeBright_Click()
    
   On Error GoTo picTextChangeBright_Click_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    
    picTextChangeBright.Visible = False
    picTextChangeDull.Visible = True
    inputDataChangedFlag = False
    txtTextEntry.SetFocus

   On Error GoTo 0
   Exit Sub

picTextChangeBright_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picTextChangeBright_Click of Form FireCallMain"
        
End Sub

' the scrollbar re-appears when marked to hide, it does this on any key up/down action on the listbox, a VB6 feature.
' this timer re-hides the scrollbar after a second or two
'---------------------------------------------------------------------------------------
' Procedure : inputScrollBarTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub inputScrollBarTimer_Timer()
   On Error GoTo inputScrollBarTimer_Timer_Error

    Call ShowScrollBar(lbxInputTextArea.hwnd, SB_VERT, False)  ' hides the vertical scrollbar

   On Error GoTo 0
   Exit Sub

inputScrollBarTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure inputScrollBarTimer_Timer of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picTextChangeBright_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picTextChangeBright_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picTextChangeBright_MouseMove_Error

        If FCWEnableBalloonTooltips = "1" Then CreateToolTip picTextChangeBright.hwnd, "When the right lamp is lit continuously it means that you have an update from your chat partner in chat. It can be extinguished by clicking upon your partner's chat box.", _
                  TTIconInfo, "Help on the update Lamp", , , , True

   On Error GoTo 0
   Exit Sub

picTextChangeBright_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picTextChangeBright_MouseMove of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picTextChangeDull_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picTextChangeDull_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picTextChangeDull_MouseMove_Error

        If FCWEnableBalloonTooltips = "1" Then CreateToolTip picTextChangeDull.hwnd, "When the right lamp is lit continuously it means that you have an update from your chat partner in chat. It can be extinguished by clicking upon your partner's chat box.", _
                  TTIconInfo, "Help on the update Lamp", , , , True

   On Error GoTo 0
   Exit Sub

picTextChangeDull_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picTextChangeDull_MouseMove of Form FireCallMain"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : picThermometer_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picThermometer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picThermometer_MouseMove_Error

        If FCWEnableBalloonTooltips = "1" Then CreateToolTip picThermometer.hwnd, "The alcohol level indicates the seconds left until the recording time is reached, a maximum of 65 seconds. When playing it indicates the track length. If you hover the cursor over the thermometer when it is playing or recording, a single line tooltip will also display giving a continuous status.", _
                  TTIconInfo, "Help on Displaying Recording Position", , , , True

   On Error GoTo 0
   Exit Sub

picThermometer_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picThermometer_MouseMove of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picTimerLampBright_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picTimerLampBright_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picTimerLampBright_MouseMove_Error

        If FCWEnableBalloonTooltips = "1" Then CreateToolTip picTimerLampBright.hwnd, "The polling lamp will regularly glow for 5 seconds to indicate that the tool is successfully polling the shared data area. It does this according to an interval set in the preferences.", _
                  TTIconInfo, "Help on Polling", , , , True

   On Error GoTo 0
   Exit Sub

picTimerLampBright_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picTimerLampBright_MouseMove of Form FireCallMain"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : picTimerLampDull_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picTimerLampDull_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picTimerLampDull_MouseMove_Error

        If FCWEnableBalloonTooltips = "1" Then CreateToolTip picTimerLampDull.hwnd, "The polling lamp will regularly glow for 5 seconds to indicate that the tool is successfully polling the shared data area. It does this according to an interval set in the preferences. The polling lamp also has another function, you may double-click on it to refresh both chat windows, this will also initiate a poll of the input file", _
                  TTIconInfo, "Help on Polling", , , , True

   On Error GoTo 0
   Exit Sub

picTimerLampDull_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picTimerLampDull_MouseMove of Form FireCallMain"
End Sub


' a click on the timer lamp will cause a repoll for the data
'---------------------------------------------------------------------------------------
' Procedure : picTimerLampDull_DblClick
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picTimerLampDull_DblClick()
   On Error GoTo picTimerLampDull_DblClick_Error

    Call btnRefresh_Click
    txtTextEntry.SetFocus

   On Error GoTo 0
   Exit Sub

picTimerLampDull_DblClick_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picTimerLampDull_DblClick of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picUtf8LampBright_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picUtf8LampBright_Click()
   On Error GoTo picUtf8LampBright_Click_Error

        picFsoLid.Visible = True

   On Error GoTo 0
   Exit Sub

picUtf8LampBright_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picUtf8LampBright_Click of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picUtf8LampDull_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picUtf8LampDull_Click()
   On Error GoTo picUtf8LampDull_Click_Error

        picFsoLid.Visible = True

   On Error GoTo 0
   Exit Sub

picUtf8LampDull_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picUtf8LampDull_Click of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picWEmailIcon_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picWEmailIcon_Click()
    
   On Error GoTo picWEmailIcon_Click_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    picWEmailIcon.Visible = False

   On Error GoTo 0
   Exit Sub

picWEmailIcon_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure picWEmailIcon_Click of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PlayTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub PlayTimer_Timer()
   On Error GoTo PlayTimer_Timer_Error

    playingTimerCount = playingTimerCount + 1
    If playingTimerMax > 0 Then ' avoid divide by zero
        linRed.X2 = 540 + ((1605 / playingTimerMax) * playingTimerCount)
    End If
    
    picThermometer.ToolTipText = "Seconds to END of recording " & playingTimerMax - playingTimerCount
    'linRed.ToolTipText = "Seconds to END of recording " & playingTimerMax - playingTimerCount
    
    If playingTimerCount = playingTimerMax Then
        Call btnStop_Click
    End If

   On Error GoTo 0
   Exit Sub

PlayTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure PlayTimer_Timer of Form FireCallMain"
End Sub

'  VB6 polling timer, the equivalent of the pollingTimer_CodeTimer
'---------------------------------------------------------------------------------------
' Procedure : pollingTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub pollingTimer_Timer()
   On Error GoTo pollingTimer_Timer_Error

    Call pollingTimer_TimerLogic

   On Error GoTo 0
   Exit Sub

pollingTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure pollingTimer_Timer of Form FireCallMain"
End Sub
' pseudo animate the action of a printer for the remote user's emoji
'---------------------------------------------------------------------------------------
' Procedure : printerTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub printerTimer_Timer()
    
   On Error GoTo printerTimer_Timer_Error

    picEmoji.Top = picEmoji.Top + 10
    If picEmoji.Top > -800 Then picEmoji.Top = picEmoji.Top + 37
    picEmoji.Refresh
    picOutputEmoji.Refresh
    
    If picEmoji.Top >= -30 Then
        printerTimer.Enabled = False
        pausePrinterTimer.Enabled = True
    End If

    'If picEmoji.Top >= 2000 Then printerTimer.Enabled = False

   On Error GoTo 0
   Exit Sub

printerTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure printerTimer_Timer of Form FireCallMain"
    
End Sub

' Sound files are recorded at 11kHz - we could divide the file length by 11000 you get a very good estimate of the duration in seconds.
' but we might change that frequency later, for the moment we place the recording timer value at the beginning of the unique filename.
'---------------------------------------------------------------------------------------
' Procedure : recordTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub recordTimer_Timer()
   On Error GoTo recordTimer_Timer_Error

    recordingTimerCount = recordingTimerCount + 1
    linRed.X2 = 540 + ((1605 / 65) * recordingTimerCount)
    
    picThermometer.ToolTipText = "Seconds to END of recording " & 65 - recordingTimerCount
    'linRed.ToolTipText = "Seconds to END of recording " & 65 - recordingTimerCount

    If recordingTimerCount = 65 Then
        Call btnStop_Click
        
    End If

   On Error GoTo 0
   Exit Sub

recordTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure recordTimer_Timer of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sendCommandTimer_PlayTimer
' Author    : beededea
' Date      : 26/07/2021
' Purpose   : the sendCommandTimer allows texts to be committed to the output file during the reading process
'             but only after a short delay, this prevents insertion of more texts during the refresh and reading of the data files
'             The routine checks to see if polling is complete before committing any change
'---------------------------------------------------------------------------------------
'
Private Sub sendCommandTimer_Timer()
    Dim stringToSend As String

    ' On Error GoTo sendCommandTimer_PlayTimer_Error
    nowBeingModifiedFlag = True ' this is a switch also set during a user run of sendSomething that allows/disallows the polling timer logic to run
    
    While pollingFlag = True  ' flag that indicates that polling is still underway
        ' we wait until the polling is complete, VB6 timers are asynchronous and so this waits until the polling has complete
    Wend

    'stringToSend = sendCommandTimer.Tag  previously used the timer tag but it only allows the one message
    'Call sendSomething(stringToSend)
    'sendCommandTimer.Tag = "" '
    
    If messageQueue.Count <> 0 Then
        ' the messages are stored in a collection, get the first item in the list
        stringToSend = messageQueue(1)
        
        If stringToSend <> "" Then Call sendSomething(stringToSend)
        messageQueue.Remove 1 ' Remove the first item in the collection at index position 1. The others shuffle up one place.
    Else
        ' check the value of the first item in the collection, when empty, none are left, so no more messages to process
        sendCommandTimer.Enabled = False
    End If
    
    On Error GoTo 0
    Exit Sub

sendCommandTimer_PlayTimer_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure sendCommandTimer_PlayTimer of Form FireCallMain"

End Sub
' pseudo animate the shredder activity
'---------------------------------------------------------------------------------------
' Procedure : shredderTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub shredderTimer_Timer()
    Dim soundtoplay As String
    
   On Error GoTo shredderTimer_Timer_Error

    picEmoji.Top = picEmoji.Top + 40
    picEmoji.Refresh

    If picEmoji.Top >= 3350 Then
        If toolTipFlag = True Then picEmoji.ToolTipText = "Click on me to show partner's Emoji status"
        shredderTimer.Enabled = False
        picEmoji.Top = -1200
        
        If FCWPlayVolume = "1" Then
            soundtoplay = App.Path & "\Resources\Sounds\" & "short.wav"
        Else
            soundtoplay = App.Path & "\Resources\Sounds\" & "shortQuiet.wav"
        End If

        If fFExists(soundtoplay) And btnLid.Visible = False Then
            If FCWEnableSounds = "1" Then PlaySound soundtoplay, ByVal 0&, SND_FILENAME Or SND_ASYNC
        End If
    End If

   On Error GoTo 0
   Exit Sub

shredderTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure shredderTimer_Timer of Form FireCallMain"

End Sub

' when the program is first run is has the text "Type your text here...", remove it on the first keypress.
'---------------------------------------------------------------------------------------
' Procedure : txtTextEntry_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtTextEntry_Click()
   On Error GoTo txtTextEntry_Click_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    If RTrim(txtTextEntry.Text) = "Type your text here..." Then
        txtTextEntry.Text = vbNullString
    End If

   On Error GoTo 0
   Exit Sub

txtTextEntry_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtTextEntry_Click of Form FireCallMain"
End Sub
' when a user hits ENTER it generates carriage return, send the text to the output file at that point
'---------------------------------------------------------------------------------------
' Procedure : txtTextEntry_KeyPress
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtTextEntry_KeyPress(KeyAscii As Integer)
   On Error GoTo txtTextEntry_KeyPress_Error

    If txtTextEntry.Text = "Type your text here..." Then
        txtTextEntry.Text = vbNullString
    End If
    
' check for the CR, set the keyascii to 0 to prevent the beeps
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call btnSendText_Click
    End If

   On Error GoTo 0
   Exit Sub

txtTextEntry_KeyPress_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtTextEntry_KeyPress of Form FireCallMain"
End Sub
' user clicks the SEND button instead of a keyboard RETURN
'---------------------------------------------------------------------------------------
' Procedure : btnSendText_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnSendText_Click()
   On Error GoTo btnSendText_Click_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    If txtTextEntry.Text = "Type your text here..." Then
        txtTextEntry.Text = " "
    End If
    Call handleStringInput(txtTextEntry.Text)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus

   On Error GoTo 0
   Exit Sub

btnSendText_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnSendText_Click of Form FireCallMain"
End Sub
' send a PING code to the remote user via the menu
'---------------------------------------------------------------------------------------
' Procedure : mnuSendPingRequest_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuSendPingRequest_Click()
   On Error GoTo mnuSendPingRequest_Click_Error

    txtTextEntry.Text = "<p><p> Refresh Interval:" & FireCallPrefs.cmbRefreshInterval.ItemData(Val(FCWRefreshIntervalIndex)) & " OS:" & WindowsVer & " Version " & App.Major & "." & App.Minor & "." & App.Revision
    Call sendSomething(txtTextEntry.Text)
    txtTextEntry.Text = vbNullString

   On Error GoTo 0
   Exit Sub

mnuSendPingRequest_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuSendPingRequest_Click of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuLBoxSendShutdownRequest_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLBoxSendShutdownRequest_Click()
   On Error GoTo mnuLBoxSendShutdownRequest_Click_Error

    mnuSendShutdownRequest_Click

   On Error GoTo 0
   Exit Sub

mnuLBoxSendShutdownRequest_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuLBoxSendShutdownRequest_Click of Form FireCallMain"
End Sub


' send a shutdown code to the remote user via the menu
'---------------------------------------------------------------------------------------
' Procedure : mnuSendShutdownRequest_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuSendShutdownRequest_Click()

    Dim dtToday As Date
    Dim UnixTimeinSec As Currency
    
   On Error GoTo mnuSendShutdownRequest_Click_Error

    dtToday = Now
    UnixTimeinSec = DateDiff("s", "1/1/1970", dtToday) & Right$(Format(Timer, "000"), 3)
    
    txtTextEntry.Text = "<z><z>" & UnixTimeinSec
    
    Call sendSomething(txtTextEntry.Text)
    txtTextEntry.Text = vbNullString

   On Error GoTo 0
   Exit Sub

mnuSendShutdownRequest_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuSendShutdownRequest_Click of Form FireCallMain"
End Sub
' hide the lid and show the underlying control panel
'---------------------------------------------------------------------------------------
' Procedure : mnuShowEmojiState_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuShowEmojiState_Click()

   On Error GoTo mnuShowEmojiState_Click_Error

        If btnLid.Visible = False Then
            btnLid.Visible = True
            picBtnLidCatch.Visible = True
            picBtnLidShadow.Visible = True
        Else
            btnLid.Visible = False
            picBtnLidCatch.Visible = False
            picBtnLidShadow.Visible = False
        End If
        
        ' could be phrased as:
        ' btnLid.Visible = btnLid.Visible = False
        ' but the above is clearer

   On Error GoTo 0
   Exit Sub

mnuShowEmojiState_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuShowEmojiState_Click of Form FireCallMain"
End Sub



' cause the main program to iconise to a stamp on the dekstop
'---------------------------------------------------------------------------------------
' Procedure : btnMinimise_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnMinimise_Click()
    
   On Error GoTo btnMinimise_Click_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    Call mnuHideProgram_Click

   On Error GoTo 0
   Exit Sub

btnMinimise_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnMinimise_Click of Form FireCallMain"
End Sub
' initiate the help screen via the default browser
'---------------------------------------------------------------------------------------
' Procedure : btnPicHelp_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicHelp_Click()

    ' variables declared
    Dim answer As VbMsgBoxResult
    
    'initialise the dimensioned variables
   On Error GoTo btnPicHelp_Click_Error

    answer = vbNo
    
    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    
   ' On Error GoTo mnuHelpPdf_click_Error
   If debugflg = 1 Then Debug.Print "%mnuHelpPdf_click"

    answer = MsgBox("This option opens a browser window and displays this tool's help. Proceed?", vbExclamation + vbYesNo)
    If answer = vbYes Then
        If fFExists(App.Path & "\help\FireCallWin Help.html") Then
            Call ShellExecute(Me.hwnd, "Open", App.Path & "\help\FireCallWin Help.html", vbNullString, App.Path, 1)
        Else
            MsgBox ("%Err-I-ErrorNumber 11 - The help file - FireCallWin Help.html - is missing from the help folder.")
        End If
    End If
    txtTextEntry.SetFocus

   On Error GoTo 0
   Exit Sub

mnuHelpPdf_click_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure mnuHelpPdf_click of Form fireCallMain"

   On Error GoTo 0
   Exit Sub

btnPicHelp_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicHelp_Click of Form FireCallMain"

End Sub

' send a PING code to the remote user via the button
'---------------------------------------------------------------------------------------
' Procedure : btnPing_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPing_Click()
   On Error GoTo btnPing_Click_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    Call mnuSendPingRequest_Click
    txtTextEntry.SetFocus

   On Error GoTo 0
   Exit Sub

btnPing_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPing_Click of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtTextEntry_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtTextEntry_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo txtTextEntry_MouseDown_Error

    If Button = vbRightButton Then
        txtTextEntry.Enabled = False
        txtTextEntry.Enabled = True
        Me.PopupMenu mainMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

txtTextEntry_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtTextEntry_MouseDown of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtTextEntry_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtTextEntry_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo txtTextEntry_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip txtTextEntry.hwnd, "This is the text box where you type your messages, pressing the SEND button to dispatch the text.", _
                  TTIconInfo, "Help on Text entry", , , , True

   On Error GoTo 0
   Exit Sub

txtTextEntry_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure txtTextEntry_MouseMove of Form FireCallMain"
End Sub

' a z-order timer that does not currently do anything, the idea was that it would maintain the chosen z-order
Private Sub zOrderTimer_Timer()
    
    'if idletime Call setZOrder(False)
End Sub

' handles cut and paste text that can have a UNIX type EOL or a Windows EOL.
'---------------------------------------------------------------------------------------
' Procedure : handleStringInput
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub handleStringInput(ByVal userText As String)
    'Dim userText As String
    Dim pos0 As Integer
    Dim pos1 As Integer

   On Error GoTo handleStringInput_Error

    pos0 = InStr(userText, vbLf) ' position of vbLf in the string
    pos1 = InStr(userText, vbCrLf) ' position of vbCrLf in the string
    
    ' if we have a valid vbLf and a vbCrLf is absent then this is a copy/paste multi line using unix EOL (likely from desktop .js)
    If pos0 > 0 And pos1 = 0 Then
        ' loop until the user text has been reduced to nothing
        Do While Len(userText) > 0
            pos0 = InStr(userText, vbLf)
            If pos0 = 0 Then ' if no vbLf found then it sends the user text to the output file
                Call sendSomething(userText)
                userText = vbNullString
            Else
                ' if we have found a valid vbLf, we call a routine to populate the array
                ' with the line up to the vbLf
                Call writeSingleLineToEndOfOutputArray(Left$(userText, pos0 - 1))
                ' now reduce the usertext to the next vbLf
                userText = Mid$(userText, pos0 + Len(vbLf))
            End If
        Loop
        If pos1 > 0 Then Call sendMultipleThings ' vbCrLf found
    ElseIf pos1 > 0 Then ' if we have a vbCrLf in the string then this is a standard multi line copy/paste Windows EOL
        ' loop until the user text has been reduced to nothing
        Do While Len(userText) > 0
            pos1 = InStr(userText, vbCrLf)
            If pos1 = 0 Then
                Call sendSomething(userText)
                userText = vbNullString
            Else
                ' we replace sendSomething with a call to a routine just to populate the array
                Call writeSingleLineToEndOfOutputArray(Left$(userText, pos1 - 1))
                userText = Mid$(userText, pos1 + Len(vbCrLf))
            End If
        Loop
        If pos1 > 0 Then Call sendMultipleThings
    Else
        Call sendSomething(userText) ' single line
    End If

   On Error GoTo 0
   Exit Sub

handleStringInput_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure handleStringInput of Form FireCallMain"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : insertStringInput
' Author    : beededea
' Date      : 03/09/2022
' Purpose   : clone of handleStringInput that handles 'inserting' text that can have either a UNIX type EOL or a Windows EOL.
'---------------------------------------------------------------------------------------
Public Sub insertStringInput(ByVal userSuppliedText As String, ByVal thisLineNumber As Long)
    Dim userText As String
    Dim pos0 As Integer
    Dim pos1 As Integer

    On Error GoTo insertStringInput_Error
    
    userText = userSuppliedText

    pos0 = InStr(userText, vbLf) ' position of vbLf in the string
    pos1 = InStr(userText, vbCrLf) ' position of vbCrLf in the string
    
    ' likely from desktop .js ie. we have a valid vbLf but a vbCrLf is absent so this is a copy/paste multi-line using unix EOL
    If pos0 > 0 And pos1 = 0 Then
        ' loop until the user text has been reduced to nothing
        Do While Len(userText) > 0
            pos0 = InStr(userText, vbLf)
            If pos0 = 0 Then ' if no vbLf found then it sends the user text to the output file
                Call insertSomething(userText, thisLineNumber) ' send single line of text
                userText = vbNullString
            Else
                ' if we have found a valid vbLf, we call a routine to populate the array
                ' with the line up to the vbLf
                Call insertLineIntoOutputArray(Left$(userText, pos0 - 1), thisLineNumber)
                ' now reduce the usertext to the next vbLf
                userText = Mid$(userText, pos0 + Len(vbLf))
            End If
        Loop
        If pos1 > 0 Then Call insertMultipleThings    ' vbCrLf found
    ElseIf pos1 > 0 Then ' if we have a vbCrLf in the string then this is a standard multi line copy/paste Windows EOL
        ' loop until the user text has been reduced to nothing
        Do While Len(userText) > 0
            pos1 = InStr(userText, vbCrLf)
            If pos1 = 0 Then
                Call insertSomething(userText, thisLineNumber) ' send single line of text
                userText = vbNullString
            Else
                ' we replace insertSomething with a call to a routine just to populate the array
                Call insertLineIntoOutputArray(Left$(userText, pos1 - 1), thisLineNumber)
                userText = Mid$(userText, pos1 + Len(vbCrLf))
            End If
        Loop
        If pos1 > 0 Then Call insertMultipleThings
    Else
        Call insertSomething(userText, thisLineNumber) ' single line
    End If

    On Error GoTo 0
    Exit Sub

insertStringInput_Error:

    With err
         If .Number <> 0 Then
            MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure insertStringInput of Form FireCallMain"
            Resume Next
          End If
    End With
End Sub


' the tooltips can be made to appear or disappear according to the prefs setting
'---------------------------------------------------------------------------------------
' Procedure : setTooltips
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setTooltips()
   On Error GoTo setTooltips_Error

    toolTipFlag = CBool(Val(FCWEnableTooltips))

    If toolTipFlag = True Then
        btnPing.ToolTipText = "Click to send a ping"
        btnCloseIt.ToolTipText = "Close the program"
        btnMinimise.ToolTipText = "Minimise the program to a desktop icon"
        picBtnLidCatch.ToolTipText = "This lid covers the emoji display, press here to open."
        
        picTextChangeDull.ToolTipText = "This lamp will glow when there has been a recent update"
        btnPicOut.ToolTipText = "Send - Just going out for a while, back later! or similar or similar"
        btnPicCod.ToolTipText = "Send - busy coding here, and you? or similar"
        btnPicBusy.ToolTipText = "Send - Very busy at the moment. or similar"
        btnPicGdn.ToolTipText = "Send - Out in the garden, doing stuff. or similar"
        btnPicPrg.ToolTipText = "Send -  Doing a bit of programming today... or similar"
        BtnPicWth.ToolTipText = "Send - What's the weather like today? or similar"
        btnPicHelp.ToolTipText = "Click to open the help for this utility"
        btnPicConfig.ToolTipText = "Click to open the configuration/preferences for this program"
        btnPicWot.ToolTipText = "Send - What's happening? or similar"
        btnPicMorn.ToolTipText = "Send Good Morning! or similar"
        btnPicNews.ToolTipText = "Send - What news? or similar"
        btnPicAttach.ToolTipText = "Click to attach a single file for transmission or similar"
        btnPicWell.ToolTipText = "Send - Are you well? or similar"
        btnPicTtfn.ToolTipText = "Send - TTFN! or similar"
        picTimerLampDull.ToolTipText = "This lamp will glow when the program is polling for new data"
        'btnRefresh.ToolTipText = "This causes the program to poll for new data and will refresh the two text areas if new data exists."
        btnSendText.ToolTipText = "Click here to send your text"
        btnEmojiSet.ToolTipText = "When you have chosen an Emoji then click here to send."
        btnClose.ToolTipText = "Click to close FireCall"
        cmbEmojiSelection.ToolTipText = "Select from a list of JPG Emojis"
        picTimerLampBright.ToolTipText = "When this lamp glows it is polling!"
        picTextChangeBright.ToolTipText = "This lamp will glow when there has been a recent update"
        picEmoji.ToolTipText = "Click on me to show partner's Emoji status"
        
        picClockSwitch.ToolTipText = "Click here to toggle between the clock and button."
        picBuzzerBrightLamp.ToolTipText = "Click here to cancel the buzzer lamp."
        picBuzzerDullLamp.ToolTipText = "This is the buzzer lamp that will light when your chat partner sounds your buzzer."
        picCloseMe.ToolTipText = "Click here to close the picture."
        
        lblDate.ToolTipText = "The current day of the month."
        lblSeconds.ToolTipText = "The seconds now."
        
        picImagePrintOut.ToolTipText = "Double click on me to open the attachment using the default application."
        
        picEmojiKnobLeft.ToolTipText = "Click to show your previous Emoji"
        picEmojiKnobRight.ToolTipText = "Click to show your next Emoji"
        picSpeakerGrille.ToolTipText = "Click to toggle to high volume"
        picSpeakerGrilleOpen.ToolTipText = "Click to toggle to low volume"
        
        picImageButton.ToolTipText = "Click to show current image"
        picBtnPlaySound.ToolTipText = "Play this recording"
        btnStartRecord.ToolTipText = "Record Button"
        btnStop.ToolTipText = "Stop Button"
        
        If recording = True Then btnStop.ToolTipText = "Stop Recording"
        If playing = True Then btnStop.ToolTipText = "Stop Playing"
        
        picRecordLampDull.ToolTipText = "This lamp glows red when recording"
        picRecordLampBright.ToolTipText = "Speech is being recorded now..."
        picFsoLid.ToolTipText = "Click this cover to reveal the FSO/UTF8 lamps"
        btnLid.ToolTipText = "These are the speech recording controls"
        picFsoLampDull.ToolTipText = "This lamp will glow when we are writing files as ANSI using FSO"
        picFsoLampBright.ToolTipText = "We are currently writing files as ANSI using FSO"
        picUtf8LampDull.ToolTipText = "This lamp will glow when writing files as UTF8"
        picUtf8LampBright.ToolTipText = "We are currently writing files as UTF8"
        
        picPlayLampDull.ToolTipText = "Lamp will light whilst recordings are played."
        picPlayLampBright.ToolTipText = "A recording is being played now"

        picThermometer.ToolTipText = "When recording, shows the time until completion."
        'linRed.ToolTipText = "When recording, shows the time until completion."
        
        'give the two listboxes tooltips
'        lbxInputTextArea.ToolTipText = "Shared input file " & FCWSharedInputFile
'        lbxOutputTextArea.ToolTipText = "Shared output file " & FCWSharedOutputFile

        
        'chkGenStartup.ToolTipText = "Check this box to enable the automatic start of the program when Windows is started."
    Else
        btnPing.ToolTipText = vbNullString
        btnCloseIt.ToolTipText = vbNullString
        btnMinimise.ToolTipText = vbNullString
        picBtnLidCatch.ToolTipText = vbNullString
        
        picTextChangeDull.ToolTipText = vbNullString
        btnPicOut.ToolTipText = vbNullString
        btnPicCod.ToolTipText = vbNullString
        btnPicBusy.ToolTipText = vbNullString
        btnPicGdn.ToolTipText = vbNullString
        btnPicPrg.ToolTipText = vbNullString
        BtnPicWth.ToolTipText = vbNullString
        btnPicHelp.ToolTipText = vbNullString
        btnPicConfig.ToolTipText = vbNullString
        btnPicWot.ToolTipText = vbNullString
        btnPicMorn.ToolTipText = vbNullString
        btnPicNews.ToolTipText = vbNullString
        btnPicAttach.ToolTipText = vbNullString
        btnPicWell.ToolTipText = vbNullString
        btnPicTtfn.ToolTipText = vbNullString
        picTimerLampDull.ToolTipText = vbNullString
        'btnRefresh.ToolTipText = vbNullString
        btnSendText.ToolTipText = vbNullString
        btnEmojiSet.ToolTipText = vbNullString
        btnClose.ToolTipText = vbNullString
        cmbEmojiSelection.ToolTipText = vbNullString
        picTimerLampBright.ToolTipText = vbNullString
        picTextChangeBright.ToolTipText = vbNullString
        picEmoji.ToolTipText = vbNullString
        picImagePrintOut.ToolTipText = vbNullString
        picClockSwitch.ToolTipText = vbNullString
        picBuzzerBrightLamp.ToolTipText = vbNullString
        picBuzzerDullLamp.ToolTipText = vbNullString
        picCloseMe.ToolTipText = vbNullString
        
        lblDate.ToolTipText = vbNullString
        lblSeconds.ToolTipText = vbNullString
        
        picEmojiKnobLeft.ToolTipText = vbNullString
        picEmojiKnobRight.ToolTipText = vbNullString
        picSpeakerGrille.ToolTipText = vbNullString
        picSpeakerGrilleOpen.ToolTipText = vbNullString
        
        picImageButton.ToolTipText = vbNullString

        picBtnPlaySound.ToolTipText = vbNullString
        btnStartRecord.ToolTipText = vbNullString
        btnStop.ToolTipText = vbNullString
        picRecordLampDull.ToolTipText = vbNullString
        picRecordLampBright.ToolTipText = vbNullString
        picFsoLid.ToolTipText = vbNullString
        btnLid.ToolTipText = vbNullString
        picFsoLampDull.ToolTipText = vbNullString
        picFsoLampBright.ToolTipText = vbNullString
        picUtf8LampDull.ToolTipText = vbNullString
        picUtf8LampBright.ToolTipText = vbNullString
        
        picPlayLampDull.ToolTipText = vbNullString
        picPlayLampBright.ToolTipText = vbNullString
        
        'give the two listboxes tooltips
        lbxInputTextArea.ToolTipText = vbNullString
        lbxOutputTextArea.ToolTipText = vbNullString
        
        picThermometer.ToolTipText = vbNullString
        'linRed.ToolTipText = vbNullString
        
        'chkGenStartup.ToolTipText = ""

    End If

   On Error GoTo 0
   Exit Sub

setTooltips_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure setTooltips of Form FireCallMain"
End Sub
' copy text from the input listbox via the menu
'---------------------------------------------------------------------------------------
' Procedure : mnuInputCopyLine_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuInputCopyLine_click()
   On Error GoTo mnuInputCopyLine_click_Error

    Call copyText(lbxInputTextArea)

   On Error GoTo 0
   Exit Sub

mnuInputCopyLine_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuInputCopyLine_click of Form FireCallMain"
End Sub
' copy text from the output listbox via the menu
'---------------------------------------------------------------------------------------
' Procedure : mnuOutputCopyLine_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuOutputCopyLine_click()
   On Error GoTo mnuOutputCopyLine_click_Error

    Call copyText(lbxOutputTextArea)

   On Error GoTo 0
   Exit Sub

mnuOutputCopyLine_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuOutputCopyLine_click of Form FireCallMain"
End Sub

' copy text from the combined listbox via the menu
'---------------------------------------------------------------------------------------
' Procedure : mnuCombinedCopyLine_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuCombinedCopyLine_click()
   On Error GoTo mnuCombinedCopyLine_click_Error

    Call copyText(lbxCombinedTextArea)

   On Error GoTo 0
   Exit Sub

mnuCombinedCopyLine_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuCombinedCopyLine_click of Form FireCallMain"
End Sub

' copy text from either of the two listboxes
'---------------------------------------------------------------------------------------
' Procedure : copyText
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub copyText(ByRef srcBox As ListBox, Optional quote As Boolean)

    Dim findStr As Integer
    Dim actualText As String
    Dim finalString As String
    Dim useloop As Integer
   
   On Error GoTo copyText_Error

    If srcBox.SelCount = 0 Then Exit Sub
   
    If srcBox.SelCount = 1 Then
        ' extract the component without the timestamp, first 23 chars removed
        ' find the first four spaces prior to the string
        
        findStr = InStr(23, srcBox.Text, "    ")
        ' the string is the next section to the end of the line after the four spaces
        actualText = LTrim(Mid$(srcBox.Text, findStr, Len(srcBox.Text)))
        If quote = True Then actualText = "[quote] " & actualText & " [/quote] "

        writeToClipboard actualText
    Else
        For useloop = 1 To srcBox.ListCount - 1
         If srcBox.Selected(useloop) Then
            findStr = InStr(23, srcBox.List(useloop), "    ")
            actualText = LTrim(Mid$(srcBox.List(useloop), findStr, Len(srcBox.List(useloop))))
            finalString = finalString & actualText & vbLf
        End If
        Next useloop
        If quote = True Then finalString = "[quote] " & finalString & " [/quote] "
        
        writeToClipboard finalString
    End If

   On Error GoTo 0
   Exit Sub

copyText_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure copyText of Form FireCallMain"
   
End Sub
'---------------------------------------------------------------------------------------
' Procedure : writeToClipboard
' Author    : beededea
' Date      : 20/06/2022
' Purpose   : help preventing a "clipboard not available error"
'---------------------------------------------------------------------------------------
'
Private Sub writeToClipboard(stringToWrite As String)

    Dim clipboardRetries As Integer
    clipboardRetries = 0

    On Error GoTo Clip_Error
    
    Clipboard.Clear
    Call Sleep(100)
    Clipboard.SetText stringToWrite
    
    Exit Sub
    
Clip_Error:
 
        If clipboardRetries > 10 Then
            MsgBox "Buggeration ! Unable to Set clipboard contents" & vbCrLf & "Try again later"
        Else
            clipboardRetries = clipboardRetries + 1
            Call Sleep(100)
            Resume Next
        End If
End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuCombinedPasteLine_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuCombinedPasteLine_click()
   On Error GoTo mnuCombinedPasteLine_click_Error

    Call mnuOutputPasteLine_click

   On Error GoTo 0
   Exit Sub

mnuCombinedPasteLine_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuCombinedPasteLine_click of Form FireCallMain"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuOutputPasteLine_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuOutputPasteLine_click()
   On Error GoTo mnuOutputPasteLine_click_Error

    DoEvents
    txtTextEntry.Text = Clipboard.GetText
    txtTextEntry.SetFocus ' set focus back to the text entry box

   On Error GoTo 0
   Exit Sub

mnuOutputPasteLine_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuOutputPasteLine_click of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuCombinedPasteAndGo_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuCombinedPasteAndGo_click()
   On Error GoTo mnuCombinedPasteAndGo_click_Error

    Call pasteAndGoHandler

   On Error GoTo 0
   Exit Sub

mnuCombinedPasteAndGo_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuCombinedPasteAndGo_click of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuOutputPasteAndGo_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuOutputPasteAndGo_click()
   On Error GoTo mnuOutputPasteAndGo_click_Error

    Call pasteAndGoHandler

   On Error GoTo 0
   Exit Sub

mnuOutputPasteAndGo_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuOutputPasteAndGo_click of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : pasteAndGoHandler
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub pasteAndGoHandler()
    'DoEvents
    
    Dim clipboardRetries As Integer
   On Error GoTo pasteAndGoHandler_Error

    clipboardRetries = 0

    On Error GoTo Clip_Error
    Call Sleep(100)
    txtTextEntry.Text = ""
    txtTextEntry.Text = Clipboard.GetText
    
    Call handleStringInput(txtTextEntry.Text)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box

    
    Exit Sub
    
Clip_Error:
 
        If clipboardRetries > 10 Then
            MsgBox "Buggeration ! Unable to read clipboard contents " & vbCrLf & "Try again later"
        Else
            clipboardRetries = clipboardRetries + 1
            Call Sleep(100)
            Resume
        End If

   On Error GoTo 0
   Exit Sub

pasteAndGoHandler_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure pasteAndGoHandler of Form FireCallMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuSwitchChatBoxes_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuSwitchChatBoxes_click()
   On Error GoTo mnuSwitchChatBoxes_click_Error

    If FCWSingleListBox = "0" Then
        FCWSingleListBox = "1"
        mnuSwitchChatBoxes.Caption = "Switch to Split Chat Box Mode"
    Else
        FCWSingleListBox = "0"
        mnuSwitchChatBoxes.Caption = "Switch to Single Chat Box"
    End If
    PutINISetting "Software\FireCallWin", "singleListBox", FCWSingleListBox, FCWSettingsFile
    
    If FCWSingleListBox = "1" Then
        lbxInputTextArea.Visible = False
        lbxOutputTextArea.Visible = False
        lbxCombinedTextArea.Height = 8415 ' force it a specific height, otherwise it defaults too short
        lbxCombinedTextArea.Visible = True
    Else
        lbxInputTextArea.Visible = True
        lbxOutputTextArea.Visible = True
        lbxCombinedTextArea.Visible = False
    End If
    
    Call btnRefresh_Click

   On Error GoTo 0
   Exit Sub

mnuSwitchChatBoxes_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuSwitchChatBoxes_click of Form FireCallMain"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : getkeypress
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : getting a keypress from the keyboard
'---------------------------------------------------------------------------------------
'
Private Sub getKeyPress(ByVal KeyCode As Integer)

    '36 home
    '40 is down
    '38 is up
    '37 is left
    '39 is right
    ' 33 page up
    ' 34 page down
    ' 35 end
    
    ' On Error GoTo getkeypress_Error
    
    If debugflg = 1 Then Debug.Print "%" & "getkeypress"
        
    Select Case KeyCode
        Case vbKeyControl
            CTRL_1 = True
        Case vbKeyC
            If CTRL_1 = True Then vbKeyCPressed = True
        Case vbKeyF
            If CTRL_1 = True Then vbKeyFPressed = True
        Case vbKeyF1
            vbKeyF1Pressed = True
        Case vbKeyF3
            vbKeyF3Pressed = True
        Case vbKeyF5
            vbKeyF5Pressed = True
    End Select
    
    If CTRL_1 And vbKeyCPressed Then
        ' if input listbox copy the current line
        If controlPressed = "lbxInputTextArea" Then
            Call copyText(lbxInputTextArea)
            controlPressed = vbNullString
        End If
        
        If controlPressed = "lbxOutputTextArea" Then
            Call copyText(lbxOutputTextArea)
            controlPressed = vbNullString
        End If
        
        If controlPressed = "lbxCombinedTextArea" Then
            Call copyText(lbxCombinedTextArea)
            controlPressed = vbNullString
        End If
        
        CTRL_1 = False
        vbKeyCPressed = False

    End If
    
    
    
    If CTRL_1 And vbKeyFPressed Then
        ' if input listbox copy the current line
        If controlPressed = "lbxInputTextArea" Then
            Call mnuFind(lbxInputTextArea, inputLineCount, "first")
            controlPressed = vbNullString
        End If
        
        If controlPressed = "lbxOutputTextArea" Then
            Call mnuFind(lbxOutputTextArea, outputLineCount, "first")
            controlPressed = vbNullString
        End If
        
        If controlPressed = "lbxCombinedTextArea" Then
            Call mnuFind(lbxCombinedTextArea, inputLineCount + outputLineCount, "first")
            controlPressed = vbNullString
        End If
        
        CTRL_1 = False
        vbKeyFPressed = False
        
    End If
    
        
    If vbKeyF3Pressed Then
        ' find
        If controlPressed = "lbxInputTextArea" Then
            If storedSearchString = vbNullString Then
                Call mnuFind(lbxInputTextArea, inputLineCount, "first")
            Else
                Call mnuFind(lbxInputTextArea, inputLineCount, "second")
            End If
            controlPressed = vbNullString
        End If
        
        If controlPressed = "lbxOutputTextArea" Then
            If storedSearchString = vbNullString Then
                Call mnuFind(lbxOutputTextArea, outputLineCount, "first")
            Else
                Call mnuFind(lbxOutputTextArea, outputLineCount, "second")
            End If
            controlPressed = vbNullString
        End If
        
        If controlPressed = "lbxCombinedTextArea" Then
            If storedSearchString = vbNullString Then
                Call mnuFind(lbxCombinedTextArea, inputLineCount + outputLineCount, "first")
            Else
                Call mnuFind(lbxCombinedTextArea, inputLineCount + outputLineCount, "second")
            End If
            controlPressed = vbNullString
        End If
        
        vbKeyF3Pressed = False
        
    End If
    
    If vbKeyF5Pressed Then
        Call mnuRefresh_Click
        vbKeyF5Pressed = False
    End If
    
    If vbKeyF1Pressed Then
        
    End If

        
    If vbKeyF1Pressed Then
        If FCWEnableBalloonTooltips = "1" Then
            Call DestroyToolTip
            FCWEnableBalloonTooltips = "0"
        Else
            FCWEnableBalloonTooltips = "1"
        End If
    End If
    
    ' the ignore list
    Select Case KeyCode
        Case vbKeyReturn
            ' return key
        Case vbKeyControl
            'Ctrl key
        Case vbKeyC
            'vbKeyCPressed = True
        Case vbKeyF
            'vbKeyFPressed = True
        Case vbKeyF3
            'vbKeyF3Pressed
        Case vbKeyF5
            'vbKeyF5Pressed
        Case vbKeyF1
            'vbKeyF1Pressed
        Case 16
            'shift
        Case 35
            ' end
        Case 36
            ' home
        Case 37
            ' left
        Case 38
            'up
        Case 39
            ' right
        Case 40
            'down
        Case 46
            'delete
        Case 8
            'backspace
        Case 112
            'F1
        Case 114
            'F3
        Case 116
            'F5
        Case Else
            ' on any normal textual/numeric keypress revert focus to the text area below the chatboxes
            If ActiveControl.Name = "lbxCombinedTextArea" Or ActiveControl.Name = "lbxInputTextArea" Or ActiveControl.Name = "lbxOutputTextArea" Then
                txtTextEntry.SetFocus
                If KeyCode <> 13 Then ' handles a RETURN keypress when listboxes have focus generating a CRLF before the next text entry
                    txtTextEntry.Text = Chr$(KeyCode)
                    txtTextEntry.SelStart = Len(txtTextEntry.Text) + 1
                End If
            End If
            
    End Select
   
    
    On Error GoTo 0
   Exit Sub

getkeypress_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure getkeypress of Form FireCallMain"

End Sub

' This is the timer that places the clock hands in the correct position on a second by second basis.
'---------------------------------------------------------------------------------------
' Procedure : clockTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub clockTimer_Timer()

    Dim intUseloop As Integer
    Dim lngStartPosX As Long
    Dim lngStartPosY As Long
    Dim dblSecondDegrees As Double
    Dim dblMinuteDegrees As Double
    Dim dblHourDegrees As Double
    Dim intSecondLbl As Integer
    
    'init the vars
   On Error GoTo clockTimer_Timer_Error

    intUseloop = 0
    lngStartPosX = 0
    lngStartPosY = 0
    dblSecondDegrees = 0
    dblMinuteDegrees = 0
    dblHourDegrees = 0
    intSecondLbl = 0
    
    'convert the time for each hand to degrees
    dblSecondDegrees = Second(Now) * 6 - 90
    dblMinuteDegrees = (Minute(Now) + Second(Now) / 60) * 6 - 90
    dblHourDegrees = (Hour(Now) + Minute(Now) / 60) * 30 - 90
    
    lngStartPosX = 1185
    lngStartPosY = 1400
    
    intSecondLbl = Second(Now)
    If intSecondLbl <= 9 Then
        lblSeconds.Caption = "0" & intSecondLbl
    Else
        lblSeconds.Caption = intSecondLbl
    End If
    
    lblDate.Caption = Day(Now)
    
    For intUseloop = 0 To 1 ' place the main image (0) and its layered companion (1)
    
      'Hour
      HourHand(intUseloop).x1 = lngStartPosX
      HourHand(intUseloop).y1 = lngStartPosY
      HourHand(intUseloop).X2 = 580 * Cos(dblHourDegrees * PI / 180) + HourHand(intUseloop).x1
      HourHand(intUseloop).Y2 = 580 * Sin(dblHourDegrees * PI / 180) + HourHand(intUseloop).y1
      
      'Minute
      MinuteHand(intUseloop).x1 = lngStartPosX
      MinuteHand(intUseloop).y1 = lngStartPosY
      MinuteHand(intUseloop).X2 = 720 * Cos(dblMinuteDegrees * PI / 180) + MinuteHand(intUseloop).x1
      MinuteHand(intUseloop).Y2 = 720 * Sin(dblMinuteDegrees * PI / 180) + MinuteHand(intUseloop).y1
      
      'Second
      SecondHand(intUseloop).x1 = lngStartPosX
      SecondHand(intUseloop).y1 = lngStartPosY
      SecondHand(intUseloop).X2 = 900 * Cos(dblSecondDegrees * PI / 180) + SecondHand(intUseloop).x1
      SecondHand(intUseloop).Y2 = 900 * Sin(dblSecondDegrees * PI / 180) + SecondHand(intUseloop).y1
      
      'Second Stub
      SecondHandStub(intUseloop).x1 = lngStartPosX
      SecondHandStub(intUseloop).y1 = lngStartPosY
      SecondHandStub(intUseloop).X2 = (200 * Cos((dblSecondDegrees - 180) * PI / 180) + SecondHandStub(intUseloop).x1)
      SecondHandStub(intUseloop).Y2 = (200 * Sin((dblSecondDegrees - 180) * PI / 180) + SecondHandStub(intUseloop).y1)
      
    Next intUseloop
    
    If MinuteHand(0).Visible = False Then
        MinuteHand(0).Visible = True
        MinuteHand(1).Visible = True
        SecondHand(0).Visible = True
        SecondHand(1).Visible = True
        HourHand(0).Visible = True
        HourHand(1).Visible = True
    End If

   On Error GoTo 0
   Exit Sub

clockTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure clockTimer_Timer of Form FireCallMain"

End Sub

' timer used to pseudo animate the pretence of a flashing emoji control area
'---------------------------------------------------------------------------------------
' Procedure : brightTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub brightTimer_Timer()
        Dim fullPath As String
        
   On Error GoTo brightTimer_Timer_Error

        picEmojiKnobLeft.Visible = False
        picEmojiKnobRight.Visible = False
        picImageButton.Visible = False
        
        flashCount = flashCount + 1
        If flashVal = 1 Then
            flashVal = 2
        Else
            flashVal = 1
        End If
        
        If flashVal = 1 Then fullPath = App.Path & "\resources\images\" & "lidBackgroundBright.jpg"
        If flashVal = 2 Then fullPath = App.Path & "\resources\images\" & "lidBackgroundDull.jpg"
        
        If flashCount > 5 Then
            fullPath = App.Path & "\resources\images\" & "lidBackgroundBright.jpg"
            brightTimer.Enabled = False
            flashCount = 0
        End If
        
        If fFExists(fullPath) Then
            picLidBackground.Picture = LoadPicture(fullPath)
        End If

   On Error GoTo 0
   Exit Sub

brightTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure brightTimer_Timer of Form FireCallMain"
End Sub

' The buzzer plays locally and sends a buzz code to the remote chat partner
'---------------------------------------------------------------------------------------
' Procedure : buzzerTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub buzzerTimer_Timer()
    Dim soundtoplay As String
    
   On Error GoTo buzzerTimer_Timer_Error

    buzzerCnt = buzzerCnt + 1
    
    ' after 5 seconds send a buzzer
    If buzzerCnt >= 3 Then
        buzzerCnt = 0
        
        If FCWPlayVolume = "1" Then
            soundtoplay = App.Path & "\Resources\Sounds\" & "buzzer.wav"
        Else
            soundtoplay = App.Path & "\Resources\Sounds\" & "buzzerQuiet.wav"
        End If
        
        If fFExists(soundtoplay) Then
            PlaySound soundtoplay, ByVal 0&, SND_FILENAME Or SND_ASYNC
        End If
        
        txtTextEntry.Text = "<b><b>"
        Call sendSomething(txtTextEntry.Text)
        txtTextEntry.Text = vbNullString
        txtTextEntry.SetFocus
    End If

   On Error GoTo 0
   Exit Sub

buzzerTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure buzzerTimer_Timer of Form FireCallMain"

End Sub
' the image of the Emoji animated as it drops
'---------------------------------------------------------------------------------------
' Procedure : dropTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub dropTimer_Timer()
   On Error GoTo dropTimer_Timer_Error

    Call dropTimer_TimerSub

   On Error GoTo 0
   Exit Sub

dropTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure dropTimer_Timer of Form FireCallMain"
End Sub
' the image of the Emoji animated as it drops
'---------------------------------------------------------------------------------------
' Procedure : dropTimer_TimerSub
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub dropTimer_TimerSub()
        Dim fullPath As String

   On Error GoTo dropTimer_TimerSub_Error

        picEmoji.Top = picEmoji.Top + 350
        picEmoji.Refresh
        picOutputEmoji.Refresh
        
        If picEmoji.Top >= 1999 Then
            dropTimer.Enabled = False
            
            picEmojiKnobLeft.Visible = True
            picEmojiKnobRight.Visible = True
            picImageButton.Visible = True
           
            fullPath = App.Path & "\resources\images\" & "lidBackgroundDull.jpg"
                            
            If fFExists(fullPath) Then
                picLidBackground.Picture = LoadPicture(fullPath)
            End If
            
            dropTimerCount = 0
            If toolTipFlag = True Then picEmoji.ToolTipText = "Click on me to shred the emoji"
        End If

   On Error GoTo 0
   Exit Sub

dropTimer_TimerSub_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure dropTimer_TimerSub of Form FireCallMain"

End Sub

' menu option to send an awake call to the remote user
'---------------------------------------------------------------------------------------
' Procedure : mnuSendAwakeCall_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuSendAwakeCall_click()
    Dim dtToday As Date
    Dim UnixTimeinSec As Currency
    
   On Error GoTo mnuSendAwakeCall_click_Error

    dtToday = Now
    UnixTimeinSec = DateDiff("s", "1/1/1970", dtToday) & Right$(Format(Timer, "000"), 3)
    txtTextEntry.Text = "<t><t>" & UnixTimeinSec
    ' 1635341466000
    ' 1635341466675
    Call sendSomething(txtTextEntry.Text)
    txtTextEntry.Text = vbNullString

   On Error GoTo 0
   Exit Sub

mnuSendAwakeCall_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuSendAwakeCall_click of Form FireCallMain"
End Sub


' menu option to find the first occurrence of a string on the input listbox
'---------------------------------------------------------------------------------------
' Procedure : mnuFindInput_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuFindInput_click()
   On Error GoTo mnuFindInput_click_Error

    Call mnuFind(lbxInputTextArea, "first")

   On Error GoTo 0
   Exit Sub

mnuFindInput_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuFindInput_click of Form FireCallMain"
End Sub
' menu option to find the first occurrence of a string on the output listbox
'---------------------------------------------------------------------------------------
' Procedure : mnuFindOutput_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuFindOutput_click()
   On Error GoTo mnuFindOutput_click_Error

    Call mnuFind(lbxOutputTextArea, "first")

   On Error GoTo 0
   Exit Sub

mnuFindOutput_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuFindOutput_click of Form FireCallMain"
End Sub
' menu option to find the first occurrence of a string on the combined listbox
'---------------------------------------------------------------------------------------
' Procedure : mnuFindCombined_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuFindCombined_click()
   On Error GoTo mnuFindCombined_click_Error

    Call mnuFind(lbxCombinedTextArea, "first")

   On Error GoTo 0
   Exit Sub

mnuFindCombined_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuFindCombined_click of Form FireCallMain"
End Sub


' menu option to show the clock face
'---------------------------------------------------------------------------------------
' Procedure : mnuShowClock_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuShowClock_click()
   On Error GoTo mnuShowClock_click_Error

    Call picClockSwitch_Click

   On Error GoTo 0
   Exit Sub

mnuShowClock_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuShowClock_click of Form FireCallMain"
End Sub
' find option on the listboxes also called by Ctrl/F
'---------------------------------------------------------------------------------------
' Procedure : mnuFind
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuFind(ByRef thisListBox As ListBox, noOfLines As Long, Optional ByVal searchType As String)

    Dim strToFind As String
    Dim useloop As Integer
    Dim answer As VbMsgBoxResult
    Dim foundString As Boolean
    
   On Error GoTo mnuFind_Error

    strToFind = vbNullString
    useloop = 0
    answer = vbNo
    foundString = False
    
    If searchType <> "second" Then
        'frmSearch.Visible = True
        strToFind = InputBox("Enter text to find : ", "Text Search")
        If strToFind = vbNullString Then
            Exit Sub
        End If
        strToFind = LCase$(strToFind)
        storedSearchString = strToFind
        storedSearchLineNo = 0
    Else
        strToFind = storedSearchString
        
    End If
    
    For useloop = storedSearchLineNo + 1 To noOfLines
        If InStr(LCase$(thisListBox.List(useloop)), strToFind) > 0 Then
            foundString = True
            storedSearchLineNo = useloop
            thisListBox.ListIndex = useloop
            thisListBox.Selected(useloop) = True
'            If searchType <> "second" Then
'                answer = MsgBox("Found the text -" & strToFind & "- on line " & useloop & ", search again?", vbYesNo, "Confirm")
'                If answer = vbNo Then
'                    Exit For
'                End If
'            Else
'                Exit For
'            End If
            Exit For
        End If
    Next useloop
    If foundString = False Then
        MsgBox "There are no more occurrences of " & strToFind & " in the current listbox, press F3 to search from the top or CTRl+F to perform a new search."
    ElseIf answer = vbYes Then
        MsgBox "No more occurrences found"
    End If

   On Error GoTo 0
   Exit Sub

mnuFind_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuFind of Form FireCallMain"
    
End Sub


    


'the following routines are only required to handle the menu generation

' right click popup menu for the close button
'---------------------------------------------------------------------------------------
' Procedure : btnClose_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnClose_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo btnClose_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mainMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

btnClose_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnClose_MouseDown of Form FireCallMain"
End Sub
' right click popup menu for the SEND button
'---------------------------------------------------------------------------------------
' Procedure : btnEmojiSet_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnEmojiSet_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo btnEmojiSet_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mainMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

btnEmojiSet_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnEmojiSet_MouseDown of Form FireCallMain"
End Sub

' right click popup menu for the attach button
'---------------------------------------------------------------------------------------
' Procedure : btnPicAttach_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicAttach_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo btnPicAttach_MouseDown_Error

    btnPicAttach.Left = btnPicAttach.Left + 10
    btnPicAttach.Top = btnPicAttach.Top + 10
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mainMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

btnPicAttach_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicAttach_MouseDown of Form FireCallMain"
End Sub


' right click popup menu for the small buttons at the base
'---------------------------------------------------------------------------------------
' Procedure : btnPicBusy_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicBusy_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo btnPicBusy_MouseDown_Error

    Call readStringsIntoTextMenu(9)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu textMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

btnPicBusy_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicBusy_MouseDown of Form FireCallMain"
End Sub

' right click popup menu for the small buttons at the base
'---------------------------------------------------------------------------------------
' Procedure : btnPicCod_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicCod_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo btnPicCod_MouseDown_Error

    Call readStringsIntoTextMenu(10)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu textMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

btnPicCod_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicCod_MouseDown of Form FireCallMain"
End Sub

' right click popup menu for the small buttons at the base
'---------------------------------------------------------------------------------------
' Procedure : btnPicConfig_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicConfig_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo btnPicConfig_MouseDown_Error

    btnPicConfig.Left = btnPicConfig.Left + 10
    btnPicConfig.Top = btnPicConfig.Top + 10
            
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mainMnuPopmenu, vbPopupMenuRightButton
    Else
        configBusyTimer.Enabled = True
    End If

   On Error GoTo 0
   Exit Sub

btnPicConfig_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicConfig_MouseDown of Form FireCallMain"
End Sub

' right click popup menu for the small buttons at the base
'---------------------------------------------------------------------------------------
' Procedure : btnPicGdn_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicGdn_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo btnPicGdn_MouseDown_Error

    Call readStringsIntoTextMenu(8)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu textMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

btnPicGdn_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicGdn_MouseDown of Form FireCallMain"
End Sub

' right click popup menu for the small buttons at the base
'---------------------------------------------------------------------------------------
' Procedure : btnPicHelp_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicHelp_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo btnPicHelp_MouseDown_Error

    btnPicHelp.Left = btnPicHelp.Left + 10
    btnPicHelp.Top = btnPicHelp.Top + 10
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mainMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

btnPicHelp_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicHelp_MouseDown of Form FireCallMain"
End Sub

' right click popup menu for the small buttons at the base
'---------------------------------------------------------------------------------------
' Procedure : btnPicMorn_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicMorn_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo btnPicMorn_MouseDown_Error

    Call readStringsIntoTextMenu(4)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu textMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

btnPicMorn_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicMorn_MouseDown of Form FireCallMain"
End Sub

' right click popup menu for the small buttons at the base
'---------------------------------------------------------------------------------------
' Procedure : btnPicNews_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicNews_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo btnPicNews_MouseDown_Error

    Call readStringsIntoTextMenu(3)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu textMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

btnPicNews_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicNews_MouseDown of Form FireCallMain"
End Sub

' right click popup menu for the small buttons at the base
'---------------------------------------------------------------------------------------
' Procedure : btnPicOut_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicOut_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo btnPicOut_MouseDown_Error

    Call readStringsIntoTextMenu(11)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu textMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

btnPicOut_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicOut_MouseDown of Form FireCallMain"
End Sub

' right click popup menu for the small buttons at the base
'---------------------------------------------------------------------------------------
' Procedure : btnPicPrg_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicPrg_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo btnPicPrg_MouseDown_Error

    Call readStringsIntoTextMenu(7)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu textMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

btnPicPrg_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicPrg_MouseDown of Form FireCallMain"
End Sub

' right click popup menu for the small buttons at the base
'---------------------------------------------------------------------------------------
' Procedure : btnPicTtfn_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicTtfn_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
    
   On Error GoTo btnPicTtfn_MouseDown_Error

    Call readStringsIntoTextMenu(1)

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu textMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

btnPicTtfn_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicTtfn_MouseDown of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : readStringsIntoTextMenu
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub readStringsIntoTextMenu(ByVal buttonNo As Integer)
    Dim textMessageArray(10) As String
    Dim b As Control
    Dim i As Integer
    

   On Error GoTo readStringsIntoTextMenu_Error

    Call readButtonTexts(buttonNo, textMessageArray())
    
    For i = 1 To 10
        If textMessageArray(i) <> "" Then
            If i = 1 Then mnuText1.Caption = textMessageArray(i): mnuText1.Visible = True
            If i = 2 Then mnuText2.Caption = textMessageArray(i): mnuText2.Visible = True
            If i = 3 Then mnuText3.Caption = textMessageArray(i): mnuText3.Visible = True
            If i = 4 Then mnuText4.Caption = textMessageArray(i): mnuText4.Visible = True
            If i = 5 Then mnuText5.Caption = textMessageArray(i): mnuText5.Visible = True
            If i = 6 Then mnuText6.Caption = textMessageArray(i): mnuText6.Visible = True
            If i = 7 Then mnuText7.Caption = textMessageArray(i): mnuText7.Visible = True
            If i = 8 Then mnuText8.Caption = textMessageArray(i): mnuText8.Visible = True
            If i = 9 Then mnuText9.Caption = textMessageArray(i): mnuText9.Visible = True
            If i = 10 Then mnuText10.Caption = textMessageArray(i): mnuText10.Visible = True
        Else

            If i = 1 Then mnuText1.Caption = "": mnuText1.Visible = False
            If i = 2 Then mnuText2.Caption = "": mnuText2.Visible = False
            If i = 3 Then mnuText3.Caption = "": mnuText3.Visible = False
            If i = 4 Then mnuText4.Caption = "": mnuText4.Visible = False
            If i = 5 Then mnuText5.Caption = "": mnuText5.Visible = False
            If i = 6 Then mnuText6.Caption = "": mnuText6.Visible = False
            If i = 7 Then mnuText7.Caption = "": mnuText7.Visible = False
            If i = 8 Then mnuText8.Caption = "": mnuText8.Visible = False
            If i = 9 Then mnuText9.Caption = "": mnuText9.Visible = False
            If i = 10 Then mnuText10.Caption = "": mnuText10.Visible = False
        End If
    Next i
    


   On Error GoTo 0
   Exit Sub

readStringsIntoTextMenu_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure readStringsIntoTextMenu of Form FireCallMain"
    
End Sub
    
    
' right click popup menu for the small buttons at the base
'---------------------------------------------------------------------------------------
' Procedure : btnPicWell_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicWell_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo btnPicWell_MouseDown_Error

    Call readStringsIntoTextMenu(2)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu textMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

btnPicWell_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicWell_MouseDown of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnPicWot_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicWot_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo btnPicWot_MouseDown_Error

    Call readStringsIntoTextMenu(5)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu textMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

btnPicWot_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicWot_MouseDown of Form FireCallMain"
End Sub

' right click popup menu for the small buttons at the base
'---------------------------------------------------------------------------------------
' Procedure : BtnPicWth_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub BtnPicWth_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo BtnPicWth_MouseDown_Error

    Call readStringsIntoTextMenu(6)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu textMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

BtnPicWth_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure BtnPicWth_MouseDown of Form FireCallMain"
End Sub

' right click popup menu for the small buttons at the base
'---------------------------------------------------------------------------------------
' Procedure : btnSendText_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnSendText_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo btnSendText_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mainMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

btnSendText_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnSendText_MouseDown of Form FireCallMain"
End Sub

' right click popup menu for the small buttons at the base
'---------------------------------------------------------------------------------------
' Procedure : btnPing_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPing_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo btnPing_MouseDown_Error

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mainMnuPopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

btnPing_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPing_MouseDown of Form FireCallMain"
End Sub
' right click popup menu to toggle polling
'---------------------------------------------------------------------------------------
' Procedure : mnuTogglePolling_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTogglePolling_click()
    
   On Error GoTo mnuTogglePolling_click_Error

    If Val(FCWRefreshIntervalSecs) = 0 Then
        mnuTogglePolling.Caption = "Disable Polling"
        FCWRefreshIntervalSecs = FireCallPrefs.cmbRefreshInterval.ItemData(Val(FCWRefreshIntervalIndex)) ' the data
    Else
        mnuTogglePolling.Caption = "Re-Enable Polling"
        FCWRefreshIntervalSecs = "0"
    End If

   On Error GoTo 0
   Exit Sub

mnuTogglePolling_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuTogglePolling_click of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : makeVisibleFormElements
' Author    : beededea
' Date      : 01/03/2023
' Purpose   : ' adjust Form Position on startup placing form onto Correct Monitor when placed off screen due to
'               monitor/resolution changes.
'---------------------------------------------------------------------------------------
'
Private Sub makeVisibleFormElements()

    Dim formLeftPixels As Long: formLeftPixels = 0
    Dim formTopPixels As Long: formTopPixels = 0

    ' read the form X/Y params from the toolSettings.ini
'    dockSettingsYPos = GetINISetting("Software\SteamyDockSettings", "dockSettingsYPos", toolSettingsFile)
'    dockSettingsXPos = GetINISetting("Software\SteamyDockSettings", "dockSettingsXPos", toolSettingsFile)
'
'    If dockSettingsYPos <> "" Then
'        dockSettings.Top = Val(dockSettingsYPos)
'    Else
'        dockSettings.Top = Screen.Height / 2 - dockSettings.Height / 2
'    End If
'
'    If dockSettingsXPos <> "" Then
'        dockSettings.Left = Val(dockSettingsXPos)
'    Else
'        dockSettings.Left = Screen.Width / 2 - dockSettings.Width / 2
'    End If

    ' read the form's saved X/Y params from the toolSettings.ini in twips and convert to pixels
    On Error GoTo makeVisibleFormElements_Error
    
    screenHeightTwips = GetDeviceCaps(Me.hdc, VERTRES) * screenTwipsPerPixelY
    screenWidthTwips = GetDeviceCaps(Me.hdc, HORZRES) * screenTwipsPerPixelX ' replaces buggy screen.width

'        FCWMaximiseFormX = fGetINISetting("Software\FireCallWin", "maximiseFormX", FCWSettingsFile)
'        FCWMaximiseFormY = fGetINISetting("Software\FireCallWin", "maximiseFormY", FCWSettingsFile)

    formLeftPixels = Val(fGetINISetting("Software\FireCallWin", "maximiseFormX", FCWSettingsFile)) / screenTwipsPerPixelX
    formTopPixels = Val(fGetINISetting("Software\FireCallWin", "maximiseFormY", FCWSettingsFile)) / screenTwipsPerPixelY

    Call adjustFormPositionToCorrectMonitor(FireCallMain.hwnd, formLeftPixels, formTopPixels)
    

    On Error GoTo 0
    Exit Sub

makeVisibleFormElements_Error:

    With err
         If .Number <> 0 Then
            MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure makeVisibleFormElements of Form FireCallMain"
            Resume Next
          End If
    End With
        
End Sub
'---------------------------------------------------------------------------------------
' Procedure : adjustMainControls
' Author    : beededea
' Date      : 12/05/2020
' Purpose   : adjust the controls so their startup position matches the last write of the config file
'---------------------------------------------------------------------------------------
'
Private Sub adjustMainControls()

    Dim fntFont As String
    Dim fntSize As Integer
    Dim fntWeight As Integer
    Dim fntStyle As Boolean
    Dim fntColour As Long
    Dim fntItalics As Boolean
    Dim fntUnderline As Boolean
    Dim imgFilePath As String

    If FCWMainFont <> vbNullString Then
        Call changeFormFont(FireCallMain, FCWMainFont, Val(FCWMainFontSize), fntWeight, fntStyle, FCWMainFontItalics, FCWMainFontColour)
    End If
    
'    If frm.Name = "FireCallMain" Then
        lblDate.ForeColor = vbWhite
        lblSeconds.ForeColor = vbWhite
'    End If

    'forces the two listboxes to a specific height regardless of fonts chosen.
    lbxInputTextArea.Height = 4300
    lbxOutputTextArea.Height = 4300
    
    If FCWClockStyle = "RedButton" Then
        picRedButton.Visible = True
        picClock.Visible = False
    Else
        picRedButton.Visible = False
        picClock.Visible = True
    End If
    
    If FCWSingleListBox = "1" Then
        lbxInputTextArea.Visible = False
        lbxOutputTextArea.Visible = False
        
        lbxCombinedTextArea.Height = 8415 ' force it a specific height, otherwise it defaults too short
        lbxCombinedTextArea.Visible = True
    Else
        lbxInputTextArea.Visible = True
        lbxOutputTextArea.Visible = True
        
        lbxCombinedTextArea.Visible = False
        
    End If
    
    If FCWPlayVolume = "1" Then
        picSpeakerGrille.Visible = False
        picSpeakerGrilleOpen.Visible = True
    Else
        picSpeakerGrille.Visible = True
        picSpeakerGrilleOpen.Visible = False
    End If


    If FCWImageDisplay = "0" Then
        picImagePrintOut.Visible = False
        'picPrintOutShadow.Visible = False

    Else
        picImagePrintOut.Visible = True
        imgFilePath = App.Path & "\Resources\images\lidBackgroundDullShadowed.jpg"
        If fFExists(imgFilePath) Then
            picLidBackground.Picture = LoadPicture(imgFilePath)
        End If
    End If
    
    If Val(FCWSendEmails) = 1 Or Val(FCWSendErrorEmails) = 1 Then
        ' start the email sending process
        Call startTheEmailTimers
    Else
        emailTimer.Enabled = False
    End If


    If Val(FCWAutomaticHousekeeping) = 1 Then
        Call startTheHouseKeepingTimers
    Else
        houseKeepingTimer.Enabled = False
    End If
    
    

   On Error GoTo 0
   Exit Sub

adjustMainControls_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure adjustMainControls of Form dockSettings on line " & Erl

End Sub

' a timer that runs once and shuts down the program on demand
'---------------------------------------------------------------------------------------
' Procedure : shutdownTimer_timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub shutdownTimer_timer()
   On Error GoTo shutdownTimer_timer_Error

    shutdownTimer.Enabled = False
    Unload FireCallMain
    'Call Form_Unload_Sub ' < we call a sub routine to shut it down

   On Error GoTo 0
   Exit Sub

shutdownTimer_timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure shutdownTimer_timer of Form FireCallMain"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : backupTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub backupTimer_Timer()
   On Error GoTo backupTimer_Timer_Error

    backupTimerCount = backupTimerCount + 1
    If backupTimerCount >= Val(FCWAutomaticBackupInterval) * 60 Then
        Call backupOutputFile(FCWSharedOutputFile, "timer")
        backupTimerCount = 1
        'MsgBox "A timed backup has been taken, once every " & Val(FCWAutomaticBackupInterval) & " hours"
    End If

   On Error GoTo 0
   Exit Sub

backupTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure backupTimer_Timer of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnStartRecord_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnStartRecord_Click()
    
    ' Record from microphone
    
    Dim cmd As String
    Dim ret As Long
    Dim soundFileName As String
    Dim bitDepth As Long
    Dim SampleRate As Long
    
   On Error GoTo btnStartRecord_Click_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity
    
    If recordingIsPossible = False Then Exit Sub
    If playing = True Then Exit Sub
    
    If toolTipFlag = True Then btnStop.ToolTipText = "Stop Recording"
    
    picRecordLampDull.Visible = False
    picRecordLampBright.Visible = True
    
    recordTimer.Enabled = True
    recording = True
    
    ' play an empty wav file to stop any sound file currently playing
    soundFileName = App.Path & "\Resources\Sounds\" & "nothing.wav"
    If fFExists(soundFileName) Then PlaySound soundFileName, ByVal 0&, SND_FILENAME Or SND_ASYNC
    
'    If FCWCaptureMethod = "0" Then
'
'        cmd = "open new Type waveaudio Alias recsound"
'        ret = mciSendString(cmd, vbNullString, 0, 0)
'
'        bitDepth = 16
'        SampleRate = 44100
'        ' bytespersec 192000
'
'        cmd = "set recSound alignment 4 bitspersample " & Str$(bitDepth) & " samplespersec " & Str$(SampleRate) & " channels 1 bytespersec " & Str$(bitDepth * SampleRate * 1 / 8) + " time format milliseconds format tag pcm"
'        ret = mciSendString(cmd, vbNullString, 0, 0)
'
'        cmd = "record recsound"
'        ret = mciSendString("record recsound", vbNullString, 0, 0)
'    Else
    
        ' // Initialize new
        If Not tSound.InitCapture(PBK_NUMOFCHANNELS, _
                                    PBK_SAMPLERATE, PBK_BITNESS, _
                                    PBK_BUFFERSIZEMS * PBK_SAMPLERATE, _
                                    cmbHiddenCaptureDevices.ListIndex) Then
            debugLog "Error during Input Audio capture initialization"
        End If
        
        tSound.StartProcess
        
        IsRecording = True
        capCount = 0
            
'    End If

   On Error GoTo 0
   Exit Sub

btnStartRecord_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnStartRecord_Click of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnStop_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnStop_Click()
    Dim cmd As String
    Dim ret As Long
    Dim soundFileName As String
    Dim fileNameToCopy As String
    Dim fileTimeStamp As String
    
   On Error GoTo btnStop_Click_Error

    If currentOpacity < 255 Then Call restoreMainWindowOpacity
        
    linRed.X2 = 540
    toolTipFlag = CBool(Val(FCWEnableTooltips))
    
    If recordingIsPossible = False Then Exit Sub
    
    If toolTipFlag = True Then
        picThermometer.ToolTipText = "When recording, shows the time until completion."
    Else
        picThermometer.ToolTipText = ""
    End If
                                                                                  
    If playing = True Then
        playing = False
        picPlayLampDull.Visible = True
        picPlayLampBright.Visible = False
        PlayTimer.Enabled = False
        playingTimerCount = 0

        soundFileName = App.Path & "\Resources\Sounds\" & "nothing.wav"
        If fFExists(soundFileName) Then PlaySound soundFileName, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    If recording = True Then
        
        soundFileName = FCWSettingsDir & "\recordings\test.wav"
        
        'LLYYYYMMDDhhmmssddd.wav (LLsecs) where LL is the length in secs and the rest is the timestamp without the separators (UTC)
        fileNameToCopy = Format$(recordingTimerCount, "00") & Format$(Now, "YYYYMMDDHHNNSS") & Right$(Format(Timer, "000"), 3) & ".wav"
        
        recordTimer.Enabled = False
                
        picRecordLampDull.Visible = True
        picRecordLampBright.Visible = False
        
        ' save new recording
        
'        If FCWCaptureMethod = "0" Then
'
'            cmd = "stop recsound " & soundFileName
'            ret = mciSendString(cmd, vbNullString, 0, 0)
'
'            cmd = "save recsound " & soundFileName
'            ret = mciSendString(cmd, vbNullString, 0, 0)
'
'            cmd = "close recsound"
'            ret = mciSendString(cmd, vbNullString, 0, 0)
'
'        Else
        
            If tSound.IsUnavailable Then
            
                ' // Unitialize previous capture session
                tSound.Uninitialize
                
                IsRecording = False
                
                Call trickSoundSave(soundFileName)
                
            End If
        
'        End If
        
        If Not fDirExists(FCWSettingsDir & "\Recordings") Then
            MsgBox FCWSettingsDir & "\Recordings" & " Folder does not exist"
        End If
                
        If fFExists(soundFileName) Then
            FileCopy soundFileName, FCWExchangeFolder & "\" & fileNameToCopy
        Else
            MsgBox FCWSettingsDir & "\Recordings" & " Sound file does not exist"
        End If
        
        ' add the text to the output file
    
        messageQueue.Add "<r><r>" & fileNameToCopy & " (" & recordingTimerCount & "secs) "
        FireCallMain.sendCommandTimer.Enabled = True ' this does a Call sendSomething(stringToSend)
        
        
        recordingTimerCount = 0
        recording = False
    End If

   On Error GoTo 0
   Exit Sub

btnStop_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnStop_Click of Form FireCallMain"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : btnClose_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo btnClose_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnClose.hwnd, "This button will close the utility and all associated windows. It has the same functionality as clicking the 'x' button, top right.", _
                  TTIconInfo, "Help on Closing", , , , True

   On Error GoTo 0
   Exit Sub

btnClose_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnClose_MouseMove of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnCloseIt_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnCloseIt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo btnCloseIt_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnCloseIt.hwnd, "This button will close the utility and all associated windows. It has the same functionality as clicking the 'Close' button, bottom right.", _
                  TTIconInfo, "Help on Closing", , , , True

   On Error GoTo 0
   Exit Sub

btnCloseIt_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnCloseIt_MouseMove of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnEmojiSet_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnEmojiSet_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo btnEmojiSet_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnEmojiSet.hwnd, "The Emoji selection Drop Down. Select an Emoji and press SEND. This will send the emoji to your chat partner. If you open the Emoji panel, bottom right you can see your current Emoji state.", _
                  TTIconInfo, "Help on Emojis", , , , True

   On Error GoTo 0
   Exit Sub

btnEmojiSet_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnEmojiSet_MouseMove of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnLid_MouseDown
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnLid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
   On Error GoTo btnLid_MouseDown_Error

     If currentOpacity < 255 Then Call restoreMainWindowOpacity
    
     If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mainMnuPopmenu, vbPopupMenuRightButton
    End If
   
    btnLid.Left = btnLid.Left + 10
    btnLid.Top = btnLid.Top + 10

   On Error GoTo 0
   Exit Sub

btnLid_MouseDown_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnLid_MouseDown of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnLid_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnLid_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo btnLid_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnLid.hwnd, "The speaker section has three buttons, one to record speech, one to play speech and the other to halt any current action. The small red lamp on the speaker panel will light up brightly when recording a message for your chat partner.", _
                  TTIconInfo, "Help on the Speaker Section", , , , True

   On Error GoTo 0
   Exit Sub

btnLid_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnLid_MouseMove of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnLid_MouseUp
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnLid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo btnLid_MouseUp_Error

    btnLid.Left = btnLid.Left - 10
    btnLid.Top = btnLid.Top - 10

   On Error GoTo 0
   Exit Sub

btnLid_MouseUp_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnLid_MouseUp of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnMinimise_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnMinimise_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo btnMinimise_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnMinimise.hwnd, "This button replaces the standard Window's minimise button but instead causes the program to fade to nothing. The program window then fades away and is replaced by an icon that sits on the desktop. You can place that icon anywhere you like on the desktop and it will remember its position when the program is next restarted.", _
                  TTIconInfo, "Help on Minimisation", , , , True

   On Error GoTo 0
   Exit Sub

btnMinimise_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnMinimise_MouseMove of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnPicAttach_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicAttach_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo btnPicAttach_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnPicAttach.hwnd, "This allows you to select a single file to send to your chat partner. This will open a file selection box. Select a file, press OK and it will be sent. It will be copied to the FCW exchange folder. The chat partner will receive a notification.", _
                  TTIconInfo, "Help on Attaching", , , , True

   On Error GoTo 0
   Exit Sub

btnPicAttach_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicAttach_MouseMove of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnPicBusy_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicBusy_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo btnPicBusy_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnPicBusy.hwnd, "Use this button to send an statement as to how busy you are in general.", _
                  TTIconInfo, "Help on the Busy Button", , , , True

   On Error GoTo 0
   Exit Sub

btnPicBusy_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicBusy_MouseMove of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnPicCod_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicCod_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo btnPicCod_MouseMove_Error

If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnPicCod.hwnd, "Use this button to send an statement as to how busy you are coding day and night...", _
                  TTIconInfo, "Help on the Coding Button", , , , True

   On Error GoTo 0
   Exit Sub

btnPicCod_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicCod_MouseMove of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnPicConfig_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicConfig_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo btnPicConfig_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnPicConfig.hwnd, "The config button opens the Preferences Utility where you can change the program settings.", _
                  TTIconInfo, "Help on Configuration", , , , True

   On Error GoTo 0
   Exit Sub

btnPicConfig_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicConfig_MouseMove of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnPicGdn_MouseMove
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicGdn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo btnPicGdn_MouseMove_Error

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnPicGdn.hwnd, "Use this button to send an statement as to how busy you are with your gardening tasks!", _
                  TTIconInfo, "Help on the TTFN Button", , , , True

   On Error GoTo 0
   Exit Sub

btnPicGdn_MouseMove_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicGdn_MouseMove of Form FireCallMain"
End Sub

Private Sub btnPicHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnPicHelp.hwnd, " This button will display the HTML help file. It will open the browser you have specified as your default browser.", _
                  TTIconInfo, "Help Button", , , , True
End Sub

Private Sub btnPicMorn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnPicMorn.hwnd, "Use this button to send a Good morning.", _
                  TTIconInfo, "Help on the Morn Button", , , , True
End Sub

Private Sub btnPicNews_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnPicNews.hwnd, "Use this button to send an inquiry as to the general news.", _
                  TTIconInfo, "Help on the News Button", , , , True
                  
End Sub

Private Sub btnPicOut_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnPicOut.hwnd, "Use this button to send an statement as to your impending absence", _
                  TTIconInfo, "Help on the Out Button", , , , True
End Sub

Private Sub btnPicPrg_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnPicPrg.hwnd, "Use this button to send a statement as to how busy you currently are programming!", _
                  TTIconInfo, "Help on the TTFN Button", , , , True
End Sub

Private Sub btnPicTtfn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnPicTtfn.hwnd, "Use this button to send a Goodbye message to your chat partner.", _
                  TTIconInfo, "Help on the TTFN Button", , , , True
End Sub

Private Sub btnPicWell_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnPicWell.hwnd, "Use this button to send an inquiry as to health of your chat partner.", _
                  TTIconInfo, "Help on the Well Button", , , , True
End Sub

Private Sub btnPicWot_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnPicWot.hwnd, "Use this button to send an inquiry as to what is happening.", _
                  TTIconInfo, "Help on the WOT Button", , , , True
End Sub

Private Sub BtnPicWth_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If FCWEnableBalloonTooltips = "1" Then CreateToolTip BtnPicWth.hwnd, "Use this button to send an inquiry as to the weather.", _
                  TTIconInfo, "Help on the WTH Button", , , , True
End Sub

Private Sub btnPing_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnPing.hwnd, "If you click here then you will send a PING to your chat partner. Your partner's client will respond with a PING response giving date and time of the response.", _
                  TTIconInfo, "Help on the Ping Button", , , , True
End Sub

Private Sub btnSendText_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnSendText.hwnd, "Pressing the SEND button dispatches the text in the text box, a press on the return key will do the same.", _
                  TTIconInfo, "Help on the SEND Button", , , , True
End Sub

Private Sub btnStartRecord_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnStartRecord.hwnd, "The record button. The maximum length is 65 seconds. Be aware that if you utilise this facility a lot you will fill up your dropbox allocation rather quickly! Best to be brief with your messages and use this function infrequently...", _
                  TTIconInfo, "Help on Recording", , , , True
End Sub

Private Sub btnStop_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If FCWEnableBalloonTooltips = "1" Then CreateToolTip btnStop.hwnd, "The stop button will halt any current action.", _
                  TTIconInfo, "Help on the Stop Button", , , , True
End Sub

'Dim retFileName As String
'Dim retfileTitle As String
'Dim attachedFile As String
'Dim fileNameToCopy As String
'
'Call addTargetFile(attachedFile, retFileName)
'
'If retFileName <> vbNullString Then
'
'    'retFileName = RTrim$(retFileName) ' this does NOT strip the padded fixed length, null padded string
'
'    txtHiddenRetFileName.Text = retFileName ' just assigning it to a text field strips the buffered bit, leaving just the filename
'    ' in this case we assign it to a hidden text box designed just for this functionality
'    retFileName = txtHiddenRetFileName.Text
'
'    Call sendSomething(retFileName)
'    fileNameToCopy = fGetFileNameFromPath(retFileName)
'    FileCopy retFileName, FCWExchangeFolder & "\" & fileNameToCopy
'End If

'


'---------------------------------------------------------------------------------------
' Procedure : addExecutableSuffixArrayList
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub addExecutableSuffixArrayList()
   On Error GoTo addExecutableSuffixArrayList_Error

    executableSuffixArrayList.Add ".bat"
    executableSuffixArrayList.Add ".bin"
    executableSuffixArrayList.Add ".cmd"
    executableSuffixArrayList.Add ".com"
    executableSuffixArrayList.Add ".cpl"
    executableSuffixArrayList.Add ".exe"
    executableSuffixArrayList.Add ".gadget"
    executableSuffixArrayList.Add ".inf1"
    executableSuffixArrayList.Add ".ins"
    executableSuffixArrayList.Add ".inx"
    executableSuffixArrayList.Add ".isu"
    executableSuffixArrayList.Add ".job"
    executableSuffixArrayList.Add ".jse"
    executableSuffixArrayList.Add ".lnk"
    executableSuffixArrayList.Add ".msc"
    executableSuffixArrayList.Add ".msi"
    executableSuffixArrayList.Add ".msp"
    executableSuffixArrayList.Add ".mst"
    executableSuffixArrayList.Add ".paf"
    executableSuffixArrayList.Add ".pif"
    executableSuffixArrayList.Add ".ps1"
    executableSuffixArrayList.Add ".reg"
    executableSuffixArrayList.Add ".rgs"
    executableSuffixArrayList.Add ".scr"
    executableSuffixArrayList.Add ".sct"
    executableSuffixArrayList.Add ".shb"
    executableSuffixArrayList.Add ".shs"
    executableSuffixArrayList.Add ".u3p"
    executableSuffixArrayList.Add ".vb"
    executableSuffixArrayList.Add ".vbe"
    executableSuffixArrayList.Add ".vbs"
    executableSuffixArrayList.Add ".vbscript"
    executableSuffixArrayList.Add ".ws"
    executableSuffixArrayList.Add ".wsf"
    executableSuffixArrayList.Add ".wsh"

   On Error GoTo 0
   Exit Sub

addExecutableSuffixArrayList_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure addExecutableSuffixArrayList of Form FireCallMain"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : addValidImageTypes
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub addValidImageTypes()
   On Error GoTo addValidImageTypes_Error

    validImageArrayList.Add ".jpg"
    validImageArrayList.Add ".jpeg"
    validImageArrayList.Add ".bmp"
    validImageArrayList.Add ".ico"
    validImageArrayList.Add ".png"
    validImageArrayList.Add ".tif"
    validImageArrayList.Add ".tiff"
    validImageArrayList.Add ".cur"
    validImageArrayList.Add ".wmf"
    validImageArrayList.Add ".emf"
    validImageArrayList.Add ".gif"

   On Error GoTo 0
   Exit Sub

addValidImageTypes_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure addValidImageTypes of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : addInvalidImageTypes
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub addInvalidImageTypes()
    
   On Error GoTo addInvalidImageTypes_Error

    invalidImageArrayList.Add ".pict"
    invalidImageArrayList.Add ".icns"
    invalidImageArrayList.Add ".ani"
    invalidImageArrayList.Add ".svg"
    invalidImageArrayList.Add ".nef"
    invalidImageArrayList.Add ".cr2"
    invalidImageArrayList.Add ".orf"
    invalidImageArrayList.Add ".rw2"
    invalidImageArrayList.Add ".arw"
    invalidImageArrayList.Add ".dng"
    invalidImageArrayList.Add ".wps"
    invalidImageArrayList.Add ".ai"
    'invalidImageArrayList.Add ".pdf"
    invalidImageArrayList.Add ".psd"
    invalidImageArrayList.Add ".raw"
    invalidImageArrayList.Add ".indd"
    invalidImageArrayList.Add ".heic"
    invalidImageArrayList.Add ".heif"

   On Error GoTo 0
   Exit Sub

addInvalidImageTypes_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure addInvalidImageTypes of Form FireCallMain"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : setListBoxFirstRun
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setListBoxFirstRun()
    ' check the existence of the default files
    
    ' set the text boxes
   On Error GoTo setListBoxFirstRun_Error

    FCWSharedInputFile = App.Path & "\input.txt"
    If Not fFExists(FCWSharedInputFile) Then FCWSharedInputFile = vbNullString

    FCWSharedOutputFile = App.Path & "\output.txt"
    If Not fFExists(FCWSharedOutputFile) Then FCWSharedOutputFile = vbNullString
        
    FCWExchangeFolder = App.Path

   On Error GoTo 0
   Exit Sub

setListBoxFirstRun_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure setListBoxFirstRun of Form FireCallMain"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : trickSoundSave
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub trickSoundSave(cdlgFileName As String)
    Dim tFmt    As WAVEFORMATEX
    Dim hWave   As Long
    Dim chkRIFF As MMCKINFO
    Dim chkData As MMCKINFO
    
   On Error GoTo trickSoundSave_Error

    On Error GoTo Cancel
    
    If capCount = 0 Then
        MsgBox "Zero size"
        Exit Sub
    End If
    
    'cdlg.ShowSave
    
    tFmt.wFormatTag = WAVE_FORMAT_PCM
    tFmt.nChannels = PBK_NUMOFCHANNELS
    tFmt.nSamplesPerSec = PBK_SAMPLERATE
    tFmt.wBitsPerSample = PBK_BITNESS
    tFmt.nBlockAlign = tFmt.nChannels * (tFmt.wBitsPerSample \ 8)
    tFmt.nAvgBytesPerSec = tFmt.nBlockAlign * tFmt.nSamplesPerSec
    
    ' // Create wave
    hWave = mmioOpen(StrPtr(cdlgFileName), ByVal 0&, MMIO_WRITE Or MMIO_CREATE)
    If hWave = 0 Then
        MsgBox "Error creating wave file"
        GoTo Cancel
    End If
    
    ' // Create RIFF-WAVE chunk
    chkRIFF.fccType = mmioStringToFOURCC("WAVE", 0)
    If mmioCreateChunk(hWave, chkRIFF, MMIO_CREATERIFF) Then
        MsgBox "Error creating RIFF-WAVE chunk"
        GoTo Cancel
    End If
    
    ' // Create fmt chunk
    chkData.ckid = mmioStringToFOURCC("fmt", 0)
    If mmioCreateChunk(hWave, chkData, 0) Then
        MsgBox "Error creating fmt chunk"
        GoTo Cancel
    End If
    
    ' // Write format
    If mmioWrite(hWave, tFmt, Len(tFmt)) = -1 Then
        MsgBox "Error writing format"
        GoTo Cancel
    End If
    
    ' // Update fmt-chunk size
    mmioAscend hWave, chkData, 0
    
    ' // Create data chunk
    chkData.ckid = mmioStringToFOURCC("data", 0)
    If mmioCreateChunk(hWave, chkData, 0) Then
        MsgBox "Error creating data chunk"
        GoTo Cancel
    End If
    
    ' // Write data
    If mmioWrite(hWave, capBuffer(0, 0), capCount * PBK_NUMOFCHANNELS * 2) = -1 Then
        MsgBox "Error writing data"
        GoTo Cancel
    End If
    
    ' // Update data-chunk size
    mmioAscend hWave, chkData, 0
    mmioAscend hWave, chkRIFF, 0
    
Cancel:
    
    If hWave Then
        mmioClose hWave
    End If

   On Error GoTo 0
   Exit Sub

trickSoundSave_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure trickSoundSave of Form FireCallMain"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : tSound_NewData
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tSound_NewData( _
            ByVal DataPtr As Long, _
            ByVal CountBytes As Long)
            
    Dim Index   As Long
    Dim total   As Long
    
   On Error GoTo tSound_NewData_Error

    If IsRecording Then
    
        Index = capCount
        capCount = capCount + (CountBytes \ PBK_NUMOFCHANNELS \ 2)
        ReDim Preserve capBuffer(PBK_NUMOFCHANNELS - 1, capCount - 1)
        
        ' // Copy captured data to buffer
        tSound.CopyData VarPtr(capBuffer(0, Index)), DataPtr, CountBytes
        
        'Redraw
        
    ElseIf IsPlayback Then
        
        total = (capCount - plyIndex) * PBK_NUMOFCHANNELS * 2
        
        If total > CountBytes Then
            total = CountBytes
        End If
        
        If total > 0 Then
        
            tSound.CopyData DataPtr, VarPtr(capBuffer(0, plyIndex)), total
        
            plyIndex = plyIndex + (CountBytes \ PBK_NUMOFCHANNELS \ 2)
        
        Else
            
            'cmdPlayback.value = vbUnchecked
            
        End If
        
    End If

   On Error GoTo 0
   Exit Sub

tSound_NewData_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure tSound_NewData of Form FireCallMain"
    
End Sub



Private Sub errMessages()

'MsgBox ("%Err-I-ErrorNumber 01 - The Shared Input File you have set is not accessible.")
'MsgBox ("%Err-I-ErrorNumber 02 - The Shared Output File you have set is not accessible.")
'MsgBox ("%Err-I-ErrorNumber 03 - The Exchange Folder you have set is not accessible.")
'MsgBox ("%Err-I-ErrorNumber 04 - Both input and output files are the same file in the same location. Attach failed.")
'MsgBox "%Err-I-ErrorNumber 05 - Sorry, can only accept one icon drop at a time, you have dropped " & data.Files.count, vbSystemModal + vbInformation
'MsgBox ("%Err-I-ErrorNumber 06 - Both the input and output folders are the same, you are copying from and to the same location. Drag & drop failed.")
'MsgBox ("%Err-I-ErrorNumber 07 - Both input and output files are the same file in the same location. Drag & drop failed.")
'MsgBox ("%Err-I-ErrorNumber 08 - For some reason that filename is invalid, possibly containing disallowed chars. Drag & drop failed.")
'MsgBox ("%Err-I-ErrorNumber 09 - The file you dragged to the program seems to be unavailable now. Drag & drop failed.")
'MsgBox "%Err-I-ErrorNumber 10 - File not found, if a recent attachment, Dropbox is possibly still transferring." & vbCrLf
'MsgBox ("%Err-I-ErrorNumber 11 - The help file - FireCallWin Help.html - is missing from the help folder.")
'MsgBox ("%Err-I-ErrorNumber 12 - FCW was unable to access the shared output file. " & vbCrLf & FCWSharedOutputFile & vbCrLf & " with " & dropboxErrCnt & " attempts")
'MsgBox ("%Err-I-ErrorNumber 13 - FCW was unable to access the shared input file. " & vbCrLf & FCWSharedInputFile & vbCrLf & " with " & dropboxErrCnt & " attempts")
'MsgBox ("%Err-I-ErrorNumber 14 - Sharing is not currently active. Outgoing messages will be saved but will not progress further.")
'MsgBox "%Err-I-ErrorNumber 15 - The output file is close to the maximum limit, please split and shorten the output file"
'MsgBox "%Err-I-ErrorNumber 16 - The output file is too long at 32,766 lines long, please split and shorten the output file. FCW will not process new messages."
'MsgBox "%Err-I-ErrorNumber 17 - The combined chat box is close to the maximum limit of lines of text, please split and shorten the input/output files or select the two chatbox option"
'MsgBox "%Err-I-ErrorNumber 18 - The combined chat box is too long at 32,766 lines long, please split and shorten the input/output files or select the two chatbox option. FCW will not process new messages."
'MsgBox "%Err-I-ErrorNumber 19 - The input file is close to the maximum limit, please split and shorten the input file"
'MsgBox "%Err-I-ErrorNumber 20 - The input file is too long at 32,766 lines long, please split and shorten the input file. FCW will not process new messages"
'MsgBox "%Err-I-ErrorNumber 21 - The polling timer is not active, the prefs are set to No Timed Refresh" & vbCrLf & "Increase value if you want it to poll for new data,"
'MsgBox "%Err-I-ErrorNumber 22 - No Audio Devices Found, the recording message functionality will be disabled."
'MsgBox "%Err-I-ErrorNumber 23 - ADO Error number 3004, a File Write Error. Dropbox synch. error (backlog)."
'Err-I-ErrorNumber 24 - No valid timestamp generated.
End Sub



'---------------------------------------------------------------------------------------
' Procedure : setSampleRate
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setSampleRate()
   On Error GoTo setSampleRate_Error

    If recordingIsPossible = True Then
        If FCWRecordingQuality = "5" Then
            PBK_NUMOFCHANNELS = 2
            PBK_SAMPLERATE = 44100
        ElseIf FCWRecordingQuality = "4" Then
            PBK_NUMOFCHANNELS = 2
            PBK_NUMOFCHANNELS = 1
            PBK_SAMPLERATE = 33075
        ElseIf FCWRecordingQuality = "3" Then
            PBK_NUMOFCHANNELS = 1
            PBK_SAMPLERATE = 22050
        ElseIf FCWRecordingQuality = "2" Then
            PBK_NUMOFCHANNELS = 1
            PBK_SAMPLERATE = 11025
        ElseIf FCWRecordingQuality = "1" Then
            PBK_NUMOFCHANNELS = 1
            PBK_SAMPLERATE = 5512
        End If
    End If

   On Error GoTo 0
   Exit Sub

setSampleRate_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure setSampleRate of Form FireCallMain"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuFindFile_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuFindFile_Click()
    Dim execStatus As Long
    Dim folderPath As String
    
    
   On Error GoTo mnuFindFile_Click_Error

    execStatus = 0
    
    If fDirExists(displayedAttachmentFilePath) Then ' if it is a folder already
        execStatus = ShellExecute(Me.hwnd, "Open", displayedAttachmentFilePath, vbNullString, vbNullString, 1)
        If execStatus <= 32 Then MsgBox "Attempt to open folder failed."
    Else
        'obtain the folder from the scommand
        folderPath = fGetDirectory(displayedAttachmentFilePath)  ' extract the default folder from the batch full path
        If fDirExists(folderPath) Then
            execStatus = ShellExecute(hwnd, "open", folderPath, vbNullString, vbNullString, 1)
            If execStatus <= 32 Then MsgBox "Attempt to open folder failed."
        End If
    End If



   On Error GoTo 0
   Exit Sub

mnuFindFile_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuFindFile_Click of Form FireCallMain"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuOpenFile_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuOpenFile_Click()
   On Error GoTo mnuOpenFile_Click_Error

    Call picImagePrintOut_DblClick

   On Error GoTo 0
   Exit Sub

mnuOpenFile_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuOpenFile_Click of Form FireCallMain"
End Sub





'
'Option Explicit 'In a blank Form
'
'Private Const LB_ADDSTRING   As Long = &H180
'Private Const LB_GETTEXT     As Long = &H189
'Private Const LB_GETTEXTLEN  As Long = &H18A
'Private Const LB_GETCOUNT    As Long = &H18B
'Private Const LB_INITSTORAGE As Long = &H1A8
'
'Private Declare Function SendMessageW Lib "user32.dll" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Declare Function SysReAllocStringLen Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long, Optional ByVal Length As Long) As Long
'
'Private WithEvents LB As VB.ListBox
'Private m_hWndLB      As Long
'
'Private Sub Form_Activate()
'    Dim FN As Integer, I As Long
'
'    FN = FreeFile
'
'    Open "Test.txt" For Output As FN
'        For I = 0& To SendMessageW(m_hWndLB, LB_GETCOUNT, 0&, 0&) - 1&
'            Print #FN, GetListBoxItem(m_hWndLB, I)
'        Next
'    Close FN
'End Sub
'
'Private Sub Form_Load()
'    Dim I As Long
'
'    Set LB = Controls.Add("VB.ListBox", "LB")
'    LB.Visible = True
'    m_hWndLB = LB.hWnd
'
'    SendMessageW m_hWndLB, LB_INITSTORAGE, &H10000, &H60000 '<-- 65536 items * 6 Bytes per item = 384 KB
'
'    For I = 0& To &H10000
'        SendMessageW m_hWndLB, LB_ADDSTRING, 0&, StrPtr(FormatNumber(I, 0&))
'    Next
'End Sub
'
'Private Sub Form_Resize()
'    LB.Move ScaleLeft, ScaleTop, ScaleWidth, ScaleHeight
'End Sub
'
'Private Function GetListBoxItem(ByVal hWndLB As Long, ByVal Index As Long) As String
'    Dim sBuffer As String
'
'    SysReAllocStringLen VarPtr(sBuffer), , SendMessageW(hWndLB, LB_GETTEXTLEN, Index, 0&)
'    SysReAllocStringLen VarPtr(GetListBoxItem), StrPtr(sBuffer), SendMessageW(hWndLB, LB_GETTEXT, Index, StrPtr(sBuffer))
'End Function



'---------------------------------------------------------------------------------------
' Procedure : mnuText1_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuText1_click()
    
   On Error GoTo mnuText1_click_Error

    Call sendSomething(mnuText1.Caption)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box

   On Error GoTo 0
   Exit Sub

mnuText1_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuText1_click of Form FireCallMain"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuText2_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuText2_click()
    
   On Error GoTo mnuText2_click_Error

    Call sendSomething(mnuText2.Caption)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box

   On Error GoTo 0
   Exit Sub

mnuText2_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuText2_click of Form FireCallMain"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuText3_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuText3_click()
    
   On Error GoTo mnuText3_click_Error

    Call sendSomething(mnuText3.Caption)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box

   On Error GoTo 0
   Exit Sub

mnuText3_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuText3_click of Form FireCallMain"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuText4_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuText4_click()
    
   On Error GoTo mnuText4_click_Error

    Call sendSomething(mnuText4.Caption)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box

   On Error GoTo 0
   Exit Sub

mnuText4_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuText4_click of Form FireCallMain"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuText5_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuText5_click()
    
   On Error GoTo mnuText5_click_Error

    Call sendSomething(mnuText5.Caption)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box

   On Error GoTo 0
   Exit Sub

mnuText5_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuText5_click of Form FireCallMain"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuText6_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuText6_click()
    
   On Error GoTo mnuText6_click_Error

    Call sendSomething(mnuText6.Caption)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box

   On Error GoTo 0
   Exit Sub

mnuText6_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuText6_click of Form FireCallMain"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuText7_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuText7_click()
    
   On Error GoTo mnuText7_click_Error

    Call sendSomething(mnuText7.Caption)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box

   On Error GoTo 0
   Exit Sub

mnuText7_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuText7_click of Form FireCallMain"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuText8_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuText8_click()
    
   On Error GoTo mnuText8_click_Error

    Call sendSomething(mnuText8.Caption)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box

   On Error GoTo 0
   Exit Sub

mnuText8_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuText8_click of Form FireCallMain"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuText9_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuText9_click()
    
   On Error GoTo mnuText9_click_Error

    Call sendSomething(mnuText9.Caption)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box

   On Error GoTo 0
   Exit Sub

mnuText9_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuText9_click of Form FireCallMain"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuText10_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuText10_click()
    
   On Error GoTo mnuText10_click_Error

    Call sendSomething(mnuText10.Caption)
    txtTextEntry.Text = vbNullString
    txtTextEntry.SetFocus ' set focus back to the text entry box

   On Error GoTo 0
   Exit Sub

mnuText10_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuText10_click of Form FireCallMain"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuInputQuoteLine_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuInputQuoteLine_click()
   On Error GoTo mnuInputQuoteLine_click_Error

    Call copyText(lbxInputTextArea, True)
    Call pasteAndGoHandler

   On Error GoTo 0
   Exit Sub

mnuInputQuoteLine_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuInputQuoteLine_click of Form FireCallMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuCombinedQuoteLine_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuCombinedQuoteLine_click()
   On Error GoTo mnuCombinedQuoteLine_click_Error

    Call copyText(lbxCombinedTextArea, True)
    Call pasteAndGoHandler

   On Error GoTo 0
   Exit Sub

mnuCombinedQuoteLine_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuCombinedQuoteLine_click of Form FireCallMain"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : sendEmailMain
' Author    : beededea
' Date      : 29/01/2022
' Purpose   : This is a duplicate of sendEmailPrefs, the reason it is duplicated rather than dropped into
'             a shared module is due to the withEvents clause on m_oProxy. Events are only generated by forms.
'             I have yet to extract this code and make it operate through the use of a class - this is not yet done.
'
' STARTTLS is an email protocol command that tells an email server that an email client,
' including an email client running in a web browser, wants to turn an existing insecure connection
' into a secure one. We use a proxy to inject that command into the CDO stream by diverting the stream
' from the desired port to the LNG_PROXY_PORT where our proxy is ready to take over.
'
'---------------------------------------------------------------------------------------
'
Private Function sendEmailMain(ByVal strSender As String, _
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
    Dim Flds ' as variant
    Dim attachment ' as variant
    Dim securityStr As String
    Dim decryptedPassword As String

    On Error GoTo sendEmailMain_Error

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
        
        decryptedPassword = AesDecryptString(FCWSmtpPassword, emailTString)

        .Item(schema & "sendpassword") = decryptedPassword
        .Update
    End With
    
    If FireCallPrefs.chkAppendConfig.Value = 1 Then
        securityStr = " SMTP server " & FCWSmtpServer & securityStr
        securityStr = securityStr & " Port:" & Val(FCWSmtpPort) & " Authentication Method:" & FireCallPrefs.cmbSmtpAuthenticate.List(Val(FCWSmtpAuthenticate))
        strSubject = strSubject & securityStr
        strBody = strBody & securityStr
    End If
    
    strSubject = strSubject & " at interval of: " & FCWAdviceIntervalSecs & " secs"

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

    sendEmailMain = True

    On Error GoTo 0
    Exit Function

sendEmailMain_Error:

    With err
         If .Number <> 0 Then
            debugLog "Error " & err.Number & " (" & err.Description & ") in procedure sendEmailMain of Form FireCallMain"
            Resume Next
          End If
    End With

End Function



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
    'pvLog "->" & Replace(sText, vbCrLf, vbCrLf & "  ")

   On Error GoTo 0
   Exit Sub

m_oProxy_RecvFromClient_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure m_oProxy_RecvFromClient of Form FireCallMain"
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
    'pvLog "<-" & Replace(sText, vbCrLf, vbCrLf & "  ")

   On Error GoTo 0
   Exit Sub

m_oProxy_RecvFromServer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure m_oProxy_RecvFromServer of Form FireCallMain"
End Sub

'Private Sub pvLog(sText As String)
'    txtEmailLog.SelStart = &H7FFF
'    txtEmailLog.SelText = sText & vbCrLf
'    txtEmailLog.SelStart = &H7FFF
'End Sub



'---------------------------------------------------------------------------------------
' Procedure : initiateEmail
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub initiateEmail(emailBody As String)

    Dim a As Boolean
    
    'MsgBox "Test email message sent. Error from the server, if any, should appear within 30 seconds. Check your Email and press get new messages!"
    
    'if the starttls option is selected then do this
   On Error GoTo initiateEmail_Error

    If FCWSmtpSecurity = 1 Then ' STARTTLS
        Set m_oProxy = New cSmtpProxy
        If m_oProxy.Init(FCWSmtpServer, FCWSmtpPort, LNG_PROXY_PORT) Then
            'pvLog "SMTP proxy listening on " & LNG_PROXY_PORT
        End If
    End If
     
    'MsgBox ("FCWRecipientEmail " & FCWRecipientEmail & " FCWEmailSubject " & FCWEmailSubject & vbCrLf & " FCWEmailMessage " & FCWEmailMessage)
    
    If FCWRecipientEmail <> "" And FCWEmailSubject <> "" And FCWEmailMessage <> "" Then
        a = sendEmailMain(FCWRecipientEmail, _
                            FCWRecipientEmail, _
                            FCWEmailSubject, _
                            emailBody)
    End If

   On Error GoTo 0
   Exit Sub

initiateEmail_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure initiateEmail of Form FireCallMain"
                        
End Sub


'---------------------------------------------------------------------------------------
' Procedure : emailTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub emailTimer_Timer()
   On Error GoTo emailTimer_Timer_Error

    Call emailTimer_TimerLogic

   On Error GoTo 0
   Exit Sub

emailTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure emailTimer_Timer of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : emailIconTimer_Timer
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub emailIconTimer_Timer()
   On Error GoTo emailIconTimer_Timer_Error

    emailIconTimer.Enabled = False
    
    ' sometimes the overall processing prevents images from appearing in their expected state
    ' so we give the process a nudge
    picWEmailIcon.Visible = False
    DoEvents
    'picWEmailIcon.Refresh
    picWEmailIcon.ToolTipText = "An Email task is underway"

   On Error GoTo 0
   Exit Sub

emailIconTimer_Timer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure emailIconTimer_Timer of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnPicAttach_MouseUp
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicAttach_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo btnPicAttach_MouseUp_Error

    btnPicAttach.Left = btnPicAttach.Left - 10
    btnPicAttach.Top = btnPicAttach.Top - 10

   On Error GoTo 0
   Exit Sub

btnPicAttach_MouseUp_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicAttach_MouseUp of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnPicConfig_MouseUp
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicConfig_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo btnPicConfig_MouseUp_Error

    btnPicConfig.Left = btnPicConfig.Left - 10
    btnPicConfig.Top = btnPicConfig.Top - 10

   On Error GoTo 0
   Exit Sub

btnPicConfig_MouseUp_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicConfig_MouseUp of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnPicHelp_MouseUp
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPicHelp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo btnPicHelp_MouseUp_Error

    btnPicHelp.Left = btnPicHelp.Left - 10
    btnPicHelp.Top = btnPicHelp.Top - 10

   On Error GoTo 0
   Exit Sub

btnPicHelp_MouseUp_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnPicHelp_MouseUp of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuOutputEditLine_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuOutputEditLine_click()
    Dim editedText As String
    Dim theText As String
    
    'theText = lbxOutputTextArea.List(lbxOutputTextArea.ListIndex)
    
    ' get the current line from the chosen list box
   On Error GoTo mnuOutputEditLine_click_Error

    theText = getCurrentLine(lbxOutputTextArea)

    ' uses an ordinary inputbox to allow the user to edit the text, might do a custom form later.
    editedText = InputBox("Edit Current Line", "Editing The Output", theText)
    
    If theText = editedText Then Exit Sub
    
    ' we need a new routine that handles cut and paste text that can have a UNIX type EOL or a Windows EOL.
    ' just as does handleStringInput - but that calls sendSomething
    ' which only write data to the end of the array and then to the file
    ' we can use writeOutputFile but we will need a new handleStringInput to handle the i/o of single or mutiple
    ' strings of text, then a new insertNewLinesIntoOutputArray routine.
    '


    ' pass the newly edited text and the current line number that is being edited
    Call insertStringInput(editedText, lbxOutputTextArea.ListIndex) ' < this latter parameter will need to be changed to allow editing of both the output andthe combined text boxes

    txtTextEntry.SetFocus

    
    ' test the contents of the input as normal to break down the line if multiple CRs
    ' read up to the line
    ' write the same
    ' insert the newly edited line
    ' write the remaining lines

   On Error GoTo 0
   Exit Sub

mnuOutputEditLine_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuOutputEditLine_click of Form FireCallMain"
    
End Sub

' get text from either of the two listboxes
'---------------------------------------------------------------------------------------
' Procedure : getCurrentLine
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function getCurrentLine(ByRef srcBox As ListBox, Optional quote As Boolean)

    Dim findStr As Integer
   
   On Error GoTo getCurrentLine_Error

    If srcBox.SelCount = 0 Then Exit Function
   
    'If srcBox.SelCount = 1 Then
    
    ' extract the component without the timestamp, first 23 chars removed
    ' find the first four spaces prior to the string
    
    findStr = InStr(23, srcBox.Text, "    ")
    ' the string is the next section to the end of the line after the four spaces
    getCurrentLine = LTrim(Mid$(srcBox.Text, findStr, Len(srcBox.Text)))

   On Error GoTo 0
   Exit Function

getCurrentLine_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure getCurrentLine of Form FireCallMain"
        
End Function

 ' set the opacity of the main window, emulating functionality of the YWE version
'---------------------------------------------------------------------------------------
' Procedure : restoreMainWindowOpacity
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub restoreMainWindowOpacity()
    
  Dim Opacity As Long
  
  'MsgBox "restoring"

   On Error GoTo restoreMainWindowOpacity_Error

  Opacity = 255
  Call setOpacity(Opacity)

   On Error GoTo 0
   Exit Sub

restoreMainWindowOpacity_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure restoreMainWindowOpacity of Form FireCallMain"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuSynchWindowsTime_click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuSynchWindowsTime_click()
    ' run the selected program
   On Error GoTo mnuSynchWindowsTime_click_Error

    Call ShellExecute(FireCallMain.hwnd, "Open", "w32tm /resync", vbNullString, vbNullString, 0)
    MsgBox ("A synch system time command has been initiated.")

   On Error GoTo 0
   Exit Sub

mnuSynchWindowsTime_click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuSynchWindowsTime_click of Form FireCallMain"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuBringToCentre_Click
' Author    : beededea
' Date      : 29/04/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuBringToCentre_Click()
   On Error GoTo mnuBringToCentre_Click_Error

    Call centreMainScreen

   On Error GoTo 0
   Exit Sub

mnuBringToCentre_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure mnuBringToCentre_Click of Form FireCallMain"
    
End Sub
