VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "GR Productions File Searcher"
   ClientHeight    =   5100
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7110
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPreview 
      Height          =   750
      Left            =   5805
      ScaleHeight     =   690
      ScaleWidth      =   1020
      TabIndex        =   45
      Top             =   1020
      Visible         =   0   'False
      Width           =   1080
      Begin VB.Image imgPreview 
         Height          =   690
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdSearchNow 
      Caption         =   "&Search now"
      Default         =   -1  'True
      Height          =   375
      Left            =   5700
      TabIndex        =   44
      Top             =   630
      Width           =   1395
   End
   Begin VB.CommandButton cmdSearchResults 
      Caption         =   "Search in &results"
      Height          =   375
      Left            =   5130
      TabIndex        =   43
      Top             =   780
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Height          =   1290
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   675
      Width           =   5160
      Begin VB.CheckBox chkSubmaps 
         Caption         =   "Search in s&ubmaps"
         Height          =   315
         Left            =   75
         TabIndex        =   4
         Tag             =   "d"
         Top             =   915
         Value           =   1  'Checked
         Width           =   5040
      End
      Begin VB.ComboBox cmbName 
         Height          =   315
         Left            =   1275
         TabIndex        =   1
         Tag             =   "d"
         Top             =   240
         Width           =   3870
      End
      Begin MSComctlLib.ImageCombo cmbPath 
         Height          =   330
         Left            =   1260
         TabIndex        =   3
         Tag             =   "d"
         ToolTipText     =   "enter multiple paths, by adding a comma (,) between the paths"
         Top             =   600
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Indentation     =   20
         ImageList       =   "imglstCombo"
      End
      Begin VB.Label Label2 
         Caption         =   "S&earch in"
         Height          =   270
         Left            =   135
         TabIndex        =   2
         Top             =   615
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "&Name"
         Height          =   330
         Left            =   135
         TabIndex        =   0
         Top             =   225
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Index           =   2
      Left            =   1590
      TabIndex        =   10
      Top             =   735
      Visible         =   0   'False
      Width           =   5160
      Begin VB.CheckBox chkCheck 
         Caption         =   "Do not check &files"
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   29
         Tag             =   "d"
         Top             =   615
         Width           =   4965
      End
      Begin VB.CheckBox chkCheck 
         Caption         =   "Do not check &read-only files"
         Height          =   255
         Index           =   3
         Left            =   195
         TabIndex        =   32
         Tag             =   "d"
         Top             =   1455
         Width           =   4965
      End
      Begin VB.CheckBox chkCheck 
         Caption         =   "Do not check s&ystem files"
         Height          =   255
         Index           =   2
         Left            =   165
         TabIndex        =   31
         Tag             =   "d"
         Top             =   1200
         Width           =   4965
      End
      Begin VB.CheckBox chkCheck 
         Caption         =   "Do not check &hidden files"
         Height          =   255
         Index           =   1
         Left            =   165
         TabIndex        =   30
         Tag             =   "d"
         Top             =   915
         Width           =   4965
      End
      Begin VB.TextBox txtSize 
         Height          =   315
         Left            =   2130
         TabIndex        =   28
         Tag             =   "d"
         Text            =   "0"
         Top             =   225
         Width           =   840
      End
      Begin VB.ComboBox cmbSize 
         Height          =   315
         ItemData        =   "Form1.frx":08CA
         Left            =   1005
         List            =   "Form1.frx":08DA
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Tag             =   "d"
         Top             =   225
         Width           =   1095
      End
      Begin MSComCtl2.UpDown UDtxtSize 
         Height          =   315
         Left            =   2970
         TabIndex        =   34
         Tag             =   "d"
         Top             =   225
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         BuddyControl    =   "txtSize"
         BuddyDispid     =   196619
         OrigLeft        =   3000
         OrigTop         =   225
         OrigRight       =   3240
         OrigBottom      =   540
         Max             =   99999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "  kB"
         Height          =   195
         Left            =   3240
         TabIndex        =   35
         Tag             =   "d"
         Top             =   270
         Width           =   285
      End
      Begin VB.Label Label3 
         Caption         =   "Siz&e is "
         Height          =   255
         Left            =   165
         TabIndex        =   26
         Top             =   315
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Index           =   1
      Left            =   1275
      TabIndex        =   9
      Top             =   660
      Visible         =   0   'False
      Width           =   5160
      Begin VB.Frame Frame2 
         Height          =   1125
         Left            =   90
         TabIndex        =   33
         Top             =   615
         Width           =   4935
         Begin VB.TextBox txtPrevDay 
            Height          =   285
            Left            =   2025
            TabIndex        =   25
            Text            =   "1"
            Top             =   675
            Width           =   375
         End
         Begin VB.TextBox txtPrevMonth 
            Height          =   285
            Left            =   2055
            TabIndex        =   23
            Text            =   "1"
            Top             =   450
            Width           =   420
         End
         Begin MSComCtl2.UpDown UDtxtPrevMonth 
            Height          =   240
            Left            =   2445
            TabIndex        =   37
            Top             =   450
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            _Version        =   393216
            Value           =   2
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtPrevMonth"
            BuddyDispid     =   196625
            OrigLeft        =   2535
            OrigTop         =   480
            OrigRight       =   2775
            OrigBottom      =   765
            Max             =   999
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UDtxtPrevDay 
            Height          =   285
            Left            =   2445
            TabIndex        =   38
            Top             =   690
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            Value           =   2
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtPrevDay"
            BuddyDispid     =   196624
            OrigLeft        =   2535
            OrigTop         =   480
            OrigRight       =   2775
            OrigBottom      =   765
            Max             =   999
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.DTPicker DTBetween1 
            Height          =   300
            Left            =   1410
            TabIndex        =   20
            Tag             =   "d"
            Top             =   150
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   393216
            Format          =   24444929
            CurrentDate     =   37577
         End
         Begin MSComCtl2.DTPicker DTBetween2 
            Height          =   300
            Left            =   2835
            TabIndex        =   21
            Tag             =   "d"
            Top             =   150
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   529
            _Version        =   393216
            Format          =   24444929
            CurrentDate     =   37577
         End
         Begin VB.OptionButton optDateSpec 
            Caption         =   "&between"
            Height          =   270
            Index           =   0
            Left            =   270
            TabIndex        =   19
            Tag             =   "d"
            Top             =   150
            Value           =   -1  'True
            Width           =   1170
         End
         Begin VB.OptionButton optDateSpec 
            Caption         =   "during the &previous"
            Height          =   270
            Index           =   1
            Left            =   300
            TabIndex        =   22
            Tag             =   "d"
            Top             =   390
            Width           =   2000
         End
         Begin VB.OptionButton optDateSpec 
            Caption         =   "during the p&revious"
            Height          =   270
            Index           =   2
            Left            =   285
            TabIndex        =   24
            Tag             =   "d"
            Top             =   660
            Width           =   2000
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "day(s)"
            Height          =   195
            Left            =   2835
            TabIndex        =   40
            Top             =   765
            Width           =   420
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "month(s)"
            Height          =   195
            Left            =   2835
            TabIndex        =   39
            Top             =   510
            Width           =   600
         End
         Begin VB.Label Label5 
            Caption         =   "and"
            Height          =   285
            Left            =   2505
            TabIndex        =   36
            Top             =   165
            Width           =   315
         End
      End
      Begin VB.OptionButton optDate 
         Caption         =   "Search all &files changed"
         Height          =   300
         Index           =   1
         Left            =   195
         TabIndex        =   18
         Tag             =   "d"
         Top             =   375
         Width           =   4080
      End
      Begin VB.OptionButton optDate 
         Caption         =   "A&ll files"
         Height          =   300
         Index           =   0
         Left            =   150
         TabIndex        =   17
         Tag             =   "d"
         Top             =   105
         Value           =   -1  'True
         Width           =   4080
      End
   End
   Begin MSComDlg.CommonDialog cmd1 
      Left            =   3315
      Top             =   2310
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.fst"
      Filter          =   "File Searcher Task File (*.fst)|*.fst|All files (*.*)|*.*"
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   5460
      ScaleHeight     =   480
      ScaleWidth      =   495
      TabIndex        =   41
      Top             =   480
      Visible         =   0   'False
      Width           =   555
   End
   Begin MSComctlLib.ImageList imglstListViewSmall 
      Left            =   3015
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08FC
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0D54
            Key             =   "unknown"
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkZip 
      Caption         =   "Search in &zipfiles"
      Height          =   315
      Left            =   90
      TabIndex        =   16
      Tag             =   "d"
      Top             =   1980
      Value           =   1  'Checked
      Width           =   2640
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   5115
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3390
      Visible         =   0   'False
      Width           =   1995
   End
   Begin MSComctlLib.ImageList imglstCombo 
      Left            =   4560
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6548
            Key             =   "cdrom"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BD3C
            Key             =   "desktop"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E4F0
            Key             =   "browse"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10CA4
            Key             =   "computer"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":17090
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D8F4
            Key             =   "documents"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":24158
            Key             =   "floppy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2994C
            Key             =   "drive"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2F140
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":318F4
            Key             =   "arrowdown"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":31D48
            Key             =   "arrowup"
         EndProperty
      EndProperty
   End
   Begin VB.FileListBox fileZip 
      Height          =   1455
      Left            =   4440
      Pattern         =   "*.zip"
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1845
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.FileListBox file 
      Height          =   1650
      Hidden          =   -1  'True
      Left            =   5745
      Pattern         =   "*.exe"
      System          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1830
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.DirListBox dir 
      Height          =   1665
      Left            =   5325
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3555
      Visible         =   0   'False
      Width           =   1755
   End
   Begin MSComctlLib.ListView Lijst 
      Height          =   2985
      Left            =   45
      TabIndex        =   6
      Top             =   2580
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   5265
      View            =   3
      Arrange         =   2
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "imglstListViewBig"
      SmallIcons      =   "imglstListViewSmall"
      ColHdrIcons     =   "imglstCombo"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "In folder"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5685
      TabIndex        =   5
      Top             =   1155
      Width           =   1395
   End
   Begin MSComctlLib.TabStrip Tabs 
      Height          =   2595
      Left            =   60
      TabIndex        =   7
      Top             =   315
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4577
      ShowTips        =   0   'False
      TabMinWidth     =   1235
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "N&ame and location"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Date"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ad&vanced"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   13
      Top             =   4830
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   476
      SimpleText      =   "gfdfdhghgfd"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "1:22"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstListViewBig 
      Left            =   1995
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3219C
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":325F4
            Key             =   "unknown"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.Animation ani 
      Height          =   675
      Left            =   6210
      TabIndex        =   42
      Top             =   225
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   1191
      _Version        =   393216
      Center          =   -1  'True
      FullWidth       =   53
      FullHeight      =   45
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Save results"
         Index           =   0
         Begin VB.Menu mnuFileSave 
            Caption         =   "&To desktop"
            Index           =   0
         End
         Begin VB.Menu mnuFileSave 
            Caption         =   "&Save to ..."
            Index           =   1
            Shortcut        =   ^S
         End
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save searchtask ..."
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Load searchtask ..."
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Windows File Search"
         Index           =   5
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Exit"
         Index           =   7
      End
   End
   Begin VB.Menu menuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit 
         Caption         =   "&Select all"
         Index           =   0
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Invert selection"
         Index           =   1
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu menuView 
      Caption         =   "&View"
      Begin VB.Menu mnuView 
         Caption         =   "&Name && location"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Date"
         Index           =   1
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Advanced"
         Index           =   2
      End
      Begin VB.Menu mnuView 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Listview"
         Index           =   4
         Begin VB.Menu mnuViewList 
            Caption         =   "&Big icons"
            Index           =   0
         End
         Begin VB.Menu mnuViewList 
            Caption         =   "&Small icons"
            Index           =   1
         End
         Begin VB.Menu mnuViewList 
            Caption         =   "&List"
            Index           =   2
         End
         Begin VB.Menu mnuViewList 
            Caption         =   "&Details"
            Index           =   3
         End
      End
      Begin VB.Menu mnuView 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Options"
         Index           =   6
         Begin VB.Menu mnuOptions 
            Caption         =   "Show &icons"
            Checked         =   -1  'True
            Index           =   0
            Tag             =   "ShowIcons"
         End
         Begin VB.Menu mnuOptions 
            Caption         =   "Show &movie"
            Checked         =   -1  'True
            Index           =   1
            Tag             =   "ShowMovie"
         End
         Begin VB.Menu mnuOptions 
            Caption         =   "&Load last used settings on start-up"
            Checked         =   -1  'True
            Index           =   2
            Tag             =   "LoadLastUsed"
         End
         Begin VB.Menu mnuOptions 
            Caption         =   "Show &fileinfo in tooltips"
            Checked         =   -1  'True
            Index           =   3
            Tag             =   "FileInfoInToolTips"
         End
         Begin VB.Menu mnuOptions 
            Caption         =   "&Quick preview of pictures"
            Checked         =   -1  'True
            Index           =   4
            Tag             =   "QuickPreview"
         End
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&About ..."
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu menuSel 
      Caption         =   "&menuSel"
      Visible         =   0   'False
      Begin VB.Menu mnuFileText 
         Caption         =   "menu for %file%"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSel 
         Caption         =   "&Open"
         Index           =   0
      End
      Begin VB.Menu mnuSel 
         Caption         =   "&Rename"
         Index           =   1
      End
      Begin VB.Menu mnuSel 
         Caption         =   "&Properties ..."
         Index           =   2
      End
      Begin VB.Menu mnuSel 
         Caption         =   "menu for %selection%"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mnuSel 
         Caption         =   "C&ut"
         Enabled         =   0   'False
         Index           =   4
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuSel 
         Caption         =   "&Copy"
         Enabled         =   0   'False
         Index           =   5
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuSel 
         Caption         =   "&Delete"
         Index           =   6
      End
      Begin VB.Menu mnuSel 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuSel 
         Caption         =   "Explore folder ..."
         Index           =   8
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public StopClicked As Boolean
Public Unloading As Boolean
Private ItemsFound As Double
Private ListHasBeenSorted As Boolean
Private SearchedInZip As Boolean
Private FSTFileName As String
Private Sub chkCheck_Click(Index As Integer)
Dim X%
If Index = 0 Then
  For X = 1 To 3
    chkCheck(X).Enabled = (chkCheck(Index).Value = 0)
  Next
End If
Select Case Index
  Case 1 'hidden files
    file.Hidden = (chkCheck(Index).Value = 0)
  Case 2
    file.System = (chkCheck(Index).Value = 0)
  Case 3
    file.ReadOnly = (chkCheck(Index).Value = 0)
End Select
fileZip.Hidden = file.Hidden
fileZip.System = file.System
fileZip.ReadOnly = file.ReadOnly
file.Refresh
fileZip.Refresh
End Sub

Private Sub cmbName_Click()
'clicked on clear history?
If cmbName.ListIndex = cmbName.ListCount - 1 Then
  'clear the combobox
  cmbName.Clear
  cmbName.AddItem "Clear history"
  'clear the history in the registry
  DeleteSetting AppName, "Recent"
  SaveSetting AppName, "Recent", "Folder", cmbPath.Text
End If
End Sub

Private Sub cmbPath_Click()
Dim P$
'browse for a file?
If cmbPath.SelectedItem.Key = "browse" Then
  P = GetFolder("Select folder to search in ...")
  If P <> "" Then
    'folder selected; set the new folder in cmbPath
    cmbPath.Text = P
  End If
End If
End Sub

Private Sub cmbPath_GotFocus()
cmbPath.SelStart = 0
cmbPath.SelLength = Len(cmbPath.Text)
End Sub

Private Sub cmbSize_Click()
'display more info?
If cmbSize.ListIndex = 3 Then
  Label4 = "Kb (range: [" + Trim(Str(Val(txtSize) - 1)) + " -> " + Trim(Str(Val(txtSize) + 1)) + "]  Kb)"
Else
  Label4 = "Kb"
End If
End Sub

Private Sub cmdSearchNow_Click()
'search for files?
Dim F As String, Obj As Object
Dim X%, P$, IsInList As Boolean
Dim C$, Folders$()

'disable all controls with "d" in the Tag
For Each Obj In Form1.Controls
  If Left(Obj.Tag, 1) = "d" Then
    If Obj.Enabled = False Then
      'if object is disabled, save it for later
      Obj.Tag = Obj.Tag + "e"
    End If
    Obj.Enabled = False
  End If
Next

'set caption
C = "all files "
If Len(cmbName) Then C = C + "with name '" + cmbName + "' "
If optDate(1).Value Then
  If optDateSpec(0) Then
    C = C + "between " + Str(DTBetween1) + " and " + Str(DTBetween2)
  End If
  If optDateSpec(1) Then
    C = C + "during the previous" + Str(txtPrevMonth) + " month(s)"
  End If
  If optDateSpec(2) Then
    C = C + "during the previous" + Str(txtPrevDay) + " day(s)"
  End If
End If
If cmbSize.ListIndex <> 0 Then
  C = C + " size is " + cmbSize + Str(Val(txtSize)) + " Kb"
End If
If Len(C) = 0 Then C = "all files"
If chkCheck(0).Value Then
  C = C + " (folders only)"
End If
Caption = "Searching for " + C

'preparing to search
cmdSearchNow.Enabled = False
Lijst.Sorted = False
cmdSearchResults.Visible = False
StopClicked = False
picPreview.Visible = False

status.Panels(1) = "Please wait while resetting ..." 'display message, because it can take some time
DoEvents
Lijst.ListItems.Clear
DoEvents
file.Pattern = "*" + cmbName + "*"

cmdStop.Enabled = True

'create a new history
P = cmbName.Text
For X = 0 To cmbName.ListCount - 1
  If cmbName.List(X) = cmbName.Text And X <= cmbName.ListCount - 1 Then
    cmbName.RemoveItem X
  End If
Next
cmbName.Text = P
cmbName.AddItem cmbName.Text, 0

'clear icons in columnheader
For X = 1 To Lijst.ColumnHeaders.Count
  Lijst.ColumnHeaders(X).Icon = 0
Next

'save the history and last used folder in the register
SaveSetting AppName, "Recent", "(no key)", "key used to raise no error if 'Recent'folder is empty"
DeleteSetting AppName, "Recent"
SaveSetting AppName, "Recent", "Folder", cmbPath.Text
For X = 0 To cmbName.ListCount - 2
  If X <= 10 Then SaveSetting AppName, "Recent", "Last Search" + Str(X), cmbName.List(X)
Next

ItemsFound = 0
ListHasBeenSorted = False
status.Panels(2).Visible = chkZip.Value
SearchedInZip = chkZip.Value

If mnuOptions(1).Checked Then ani.Play
mnuOptions(1).Enabled = False

'custom drive selected?
If cmbPath.SelectedItem Is Nothing Then
  'search (multiple) path(s)/drive(s)
  Folders = Split(cmbPath.Text, ",") 'split all paths (splitted with a comma)
  For X = 0 To UBound(Folders)
    status.Panels(1).Text = Folders(X)
    SearchInSub Folders(X)
    If Unloading Then Exit For
  Next
Else
  'preselected item
  P = cmbPath.SelectedItem.Key
  If P = "drives" Then
    For X = InStr(cmbPath.Text, ":") - 1 To Len(cmbPath.Text) Step 3
      SearchInSub Mid(cmbPath.Text, X, 2)
    Next
  ElseIf Right(P, 1) = ":" Then
    SearchInSub P
  ElseIf P = "document folders" Then
    SearchInSub "c:\mijn documenten\"
    SearchInSub "c:\windows\desktop\"
  ElseIf P = "desktop" Then
    SearchInSub "c:\windows\desktop\"
  ElseIf P = "my documents" Then
    SearchInSub "c:\mijn documenten\"
  ElseIf P = "my computer" Then
    For X = 0 To Drive.ListCount - 1
'      InSubSearched = False
      SearchInSub Left(Drive.List(X), 2)
    Next
  End If
End If

'reset the animation
If mnuOptions(1).Checked Then
  ani.Play , 0
  ani.Stop
End If
mnuOptions(1).Enabled = True

'enable all controls, except if the Tag propert is set to "de"
For Each Obj In Me.Controls
  If Left(Obj.Tag, 1) = "d" Then
    If Right(Obj.Tag, 1) <> "e" Then
      Obj.Enabled = True
    Else
      Obj.Tag = Mid(Obj.Tag, 1, Len(Obj.Tag) - 1)
    End If
  End If
Next

cmdSearchNow.Enabled = True
cmdStop.Enabled = False
If Lijst.ListItems.Count Then cmdSearchResults.Visible = True

'display info
status.Panels(2).Text = ""
If Unloading Then mnuFile_Click -1
If Tabs.Tabs(1).Selected Then cmbName.SetFocus 'select search box, when visible
status.Panels(1).Text = Trim(Str(ItemsFound)) + " item" + IIf(ItemsFound <> 1, "s", "") + " found" + IIf(StopClicked, "(search cancelled)", "")
End Sub

Private Sub cmdSearchResults_Click()
Dim Text$, X%
Text = InputBox("Enter text to search for in the filenames of found files" + vbCrLf + "You can use wildcards" + vbCrLf + "P.S.: files in archives will not be shown")
If Text = "" Then Exit Sub
Text = LCase("*" + Text + "*")
With Lijst.ListItems
  For X = 1 To .Count
    If X <= .Count Then
      If LCase(.Item(X).Text) Like Text = False Or Trim(.Item(X).Key) = "" Then
        .Remove X
        X = X - 1
      End If
    End If
  Next
  ListHasBeenSorted = True
  If .Count = 0 Then cmdSearchResults.Visible = False
  status.Panels(1) = Trim(Str(Lijst.ListItems.Count)) + " item" + IIf(Lijst.ListItems.Count = 1, "", "s") + " found"
End With
End Sub

Private Sub cmdStop_Click()
'stop the search
StopClicked = True
cmdStop.Enabled = False
End Sub
Private Sub ShowProps() 'ByVal Filename As String)
'show properties for a file
Dim SEI As SHELLEXECUTEINFO
Dim R As Long ', Files As String, X%
'For X = 1 To Lijst.ListItems.Count
'  If Lijst.ListItems(X).Selected And Lijst.ListItems(X).Key <> "" Then
'    Files = Files + Lijst.ListItems(X).Key + vbNullChar
'  End If
'Next
'Files = Files + vbNullChar
With SEI
  .cbSize = Len(SEI)
  .fMask = SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
  .lpVerb = "properties"
  .lpFile = Lijst.SelectedItem.Key ' Files
  .nShow = 5
End With
ShellExecuteEX SEI
End Sub
Sub DeleteFile()
  'bestandsnamen wissen en naar Prullenbak verplaatsen
'  Dim SHFileOp As SHFILEOPSTRUCT
'  Dim fList As String, x%
'Dim R As Long, Files As String
''For x = 1 To Lijst.ListItems.Count
''  If Lijst.ListItems(x).Selected And Lijst.ListItems(x).Key <> "" Then
''    Files = Files + Lijst.ListItems(x).Key + vbNullChar
''  End If
''Next
''Files = Files + vbNullChar
'Files = Lijst.SelectedItem.Key
''structuur vullen
'With SHFileOp
'  .hwnd = Me.hwnd
'  .wFunc = &H3   'delete
'  .pFrom = fList 'files
'  .fFlags = &H40 'Undohistory allowed
'  .lpszProgressTitle = "Deleting files ..."
'End With
'SHFileOperation SHFileOp
End Sub

Private Sub dir_Change()
'updates filecontrols
file.Path = dir.Path
fileZip.Path = dir.Path
file.Refresh
fileZip.Refresh
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
  Lijst.StartLabelEdit
End If
If Shift = 4 And KeyCode = vbKeyReturn Then
  mnuSel_Click 2
End If
If KeyCode = vbKeyDelete Then
  mnuSel_Click 6
End If
If Shift = 2 And KeyCode = vbKeyE Then
  mnuSel_Click 8
End If
End Sub

Private Sub Form_Load()
'init
Dim X%, Drv$, Reg
Const Ind = 2 'intendation

If GetSetting(AppName, "Options", "AboutBoxShowed") = "" Then
  Form2.Show 1, Me
End If

SetFont "arial"
ani.Open (RightPath(App.Path) + "findfile.avi")

'init combo with pahts
With cmbPath.ComboItems
  .Clear
  .Add , "document folders", "Document folders", "folder", , 0
  .Add , "desktop", "Desktop", "desktop", , Ind
  .Add , "my documents", "My documents", "documents", , Ind
  .Add , "my computer", "My computer", "computer", , 0
  For X = 0 To Drive.ListCount - 1
    If GetDrive(Drive.List(X)) = "drive" Then
      Drv = Drv + UCase(Left(Drive.List(X), 2)) + ","
    End If
  Next
  Drv = Mid(Drv, 1, Len(Drv) - 1)
  .Add , "drives", "Local drives (" + Drv + ")", "drive", , Ind
  For X = 0 To Drive.ListCount - 1
    .Add , LCase(Left(Drive.List(X), 2)), Drive.List(X), GetDrive(Drive.List(X)), , Ind
  Next
  .Add , "browse", "Browse ...", "browse", , 0
End With

'get history, and put it in cmbName
SaveSetting AppName, "Recent", "No key", "No Value"
Reg = GetAllSettings(AppName, "Recent")
For X = UBound(Reg) To 0 Step -1
  If InStr(LCase(Reg(X, 0)), "search") Then
    cmbName.AddItem Reg(X, 1), 0
  End If
Next
cmbName.AddItem "Clear history"

'check what was the last used folder setting
For X = 1 To cmbPath.ComboItems.Count
  If cmbPath.ComboItems(X) = GetSetting(AppName, "Recent", "Folder", "C:") Then
    cmbPath.ComboItems(X).Selected = True
    Exit For
  End If
Next
If cmbPath.SelectedItem Is Nothing Then
  cmbPath.Text = GetSetting(AppName, "Recent", "Folder", "c:\")
End If


mnuViewList_Click GetSetting(AppName, "Options", "View", 3)
For X = 0 To mnuOptions.UBound
  If mnuOptions(X).Tag <> "" Then mnuOptions(X).Checked = GetSetting(AppName, "Options", mnuOptions(X).Tag, True)
Next
mnuSel(1).Caption = mnuSel(1).Caption + vbTab + "F2"
mnuSel(2).Caption = mnuSel(2).Caption + vbTab + "Alt+Enter"
mnuSel(6).Caption = mnuSel(6).Caption + vbTab + "DEL"
mnuSel(8).Caption = mnuSel(8).Caption + vbTab + "Ctrl+E"
mnuFile(7).Caption = mnuFile(7).Caption + vbTab + "Alt+F4"

cmbSize.ListIndex = 0
optDate_Click 0

If mnuOptions(2).Checked Then
  LoadFileSearchTask RightPath(App.Path) + "autosave.fst"
End If

Set Archive = New Collection 'zip
End Sub
Sub SetFont(Font As String)
On Error Resume Next
Dim Obj As Object
For Each Obj In Me.Controls
  Obj.FontName = Font
Next
End Sub

Private Sub Form_Resize()
Dim X%
On Error Resume Next
Const Side = 200
Tabs.Move Side / 4, Side / 2, Width - cmdSearchNow.Width - 3 * Side

Lijst.Top = Tabs.Height + Tabs.Top + Side
Lijst.Move Tabs.Left, Lijst.Top, Width - 1 * Side, Height - 3 * Side - Lijst.Top - status.Height - Side / 2
cmdSearchNow.Move Tabs.Width + Side / 2 + Side, Tabs.Top + Tabs.Tabs(1).Height
cmdStop.Move cmdSearchNow.Left, cmdSearchNow.Top + Side / 2 + cmdSearchNow.Height
cmdSearchResults.Move cmdStop.Left, cmdStop.Top + cmdStop.Height + Side / 2
ani.Move cmdStop.Left + Side, cmdSearchResults.Top + cmdSearchResults.Height + Side * 2
picPreview.Move ani.Left, ani.Top

For X = 0 To 2
  Frame1(X).Move Tabs.Left + Side / 2, Tabs.Top + Tabs.Tabs(1).Height + Side / 2, Tabs.Width - Side, Tabs.Height - Tabs.Tabs(1).Height - Side * 2.2
Next

'name&location tab
Label1.Left = Side
Label2.Left = Side
cmbName.Move Label1.Width + Side / 2, cmbName.Top, Frame1(0).Width - Side - Label1.Width
cmbPath.Move cmbName.Left, cmbName.Top + cmbName.Height + 10, cmbName.Width
chkSubmaps.Move Label1.Left, chkSubmaps.Top, Frame1(0).Width - 2 * Side
chkZip.Move Frame1(0).Left, Frame1(0).Top + Frame1(0).Height + Side / 3, Tabs.Width - Side, Side

'date tab
optDate(0).Move Side, Side / 1.5, Frame1(1).Width - 2 * Side
optDate(1).Move Side, optDate(0).Top + optDate(0).Height, Frame1(1).Width - 2 * Side
Frame2.Move Side / 2, optDate(1).Top + optDate(1).Height - Side / 3, Frame1(1).Width - Side
optDateSpec(0).Top = Side / 1.5
optDateSpec(1).Move optDateSpec(0).Left, optDateSpec(0).Top + optDateSpec(0).Height + 50  ',  2000
optDateSpec(2).Move optDateSpec(1).Left, optDateSpec(1).Top + optDateSpec(1).Height, 2000 'optDateSpec(1).Width
txtPrevMonth.Move optDateSpec(1).Width, optDateSpec(1).Top, txtPrevMonth.Width, optDateSpec(1).Height
txtPrevDay.Move optDateSpec(2).Width, optDateSpec(2).Top, txtPrevMonth.Width, optDateSpec(2).Height
UDtxtPrevMonth.Move txtPrevMonth.Left + txtPrevMonth.Width, txtPrevMonth.Top
UDtxtPrevDay.Move txtPrevDay.Left + txtPrevDay.Width, txtPrevDay.Top

'advanced tab
Label3.Move Side, Side + 40
cmbSize.Move Label3.Width, Label3.Top - 40
txtSize.Move cmbSize.Left + cmbSize.Width, cmbSize.Top
UDtxtSize.Move txtSize.Left + txtSize.Width, txtSize.Top
Label4.Move UDtxtSize.Left + UDtxtSize.Width, txtSize.Top + 40
chkCheck(0).Move Side, cmbSize.Top + cmbSize.Height + Side / 2, Frame2.Width - Side * 2
For X = 1 To chkCheck.UBound
  chkCheck(X).Move chkCheck(X - 1).Left, chkCheck(X - 1).Top + chkCheck(X - 1).Height, chkCheck(X - 1).Width
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unloading = True
StopClicked = True
SaveFileSearchTask RightPath(App.Path) + "autosave.fst"
End Sub

Private Sub Lijst_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim P As String
On Error GoTo Fout
If Right(Lijst.SelectedItem.Key, 1) = "\" Then 'folder renamed
  P = Left(Lijst.SelectedItem.Key, Len(Lijst.SelectedItem.Key) - 1)
  P = Mid(P, 1, InStrRev(P, "\"))
  Name Lijst.SelectedItem.Key As P + NewString
Else
  P = RightPath(Left(Lijst.SelectedItem.Key, InStrRev(Lijst.SelectedItem.Key, "\") - 1))
  Name Lijst.SelectedItem.Key As P + NewString
End If
Lijst.SelectedItem.Key = P + NewString
Exit Sub
Fout:
  MsgBox "Error while renaming '" + Lijst.SelectedItem.Key + "' !" + vbCrLf + vbCrLf + Err.Description, vbCritical
End Sub

Private Sub Lijst_BeforeLabelEdit(Cancel As Integer)
If Lijst.SelectedItem.Key = "" Then
  Cancel = True
  Beep
End If
End Sub

Private Sub Lijst_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim X%
If cmdSearchNow.Enabled = False Then Exit Sub
If ListHasBeenSorted = False And SearchedInZip = True Then
  If MsgBox("If you sort the list, the files in the archives (grayed) will not be shown." + vbCrLf + "Continue?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
End If
'if list is sorted, the files inside a zip are (normally) not under
'the right archive
For X = 1 To Lijst.ListItems.Count
  If X <= Lijst.ListItems.Count Then
    If Lijst.ListItems(X).Key = "" Then
      Lijst.ListItems.Remove X
      X = X - 1
    End If
  End If
Next
ListHasBeenSorted = True
Lijst.SortKey = ColumnHeader.Index - 1
Lijst.Sorted = True
Lijst.SortOrder = Abs(Lijst.SortOrder = 0)
For X = 1 To Lijst.ColumnHeaders.Count
  Lijst.ColumnHeaders(X).Icon = 0
Next
If Lijst.SortOrder = lvwAscending Then
  ColumnHeader.Icon = "arrowdown"
Else
  ColumnHeader.Icon = "arrowup"
End If
End Sub

Private Sub Lijst_DblClick()
If Lijst.SelectedItem Is Nothing Then Exit Sub
If Lijst.SelectedItem.Key = "" Then Exit Sub
ShellExecute hwnd, "open", Lijst.SelectedItem.Key + Chr(0), "", RightPath(Left(Lijst.SelectedItem.Key, InStrRev(Lijst.SelectedItem.Key, "\") - 1)), 1
End Sub

Private Sub Lijst_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim LstItem As ListItem
Set LstItem = Lijst.HitTest(X, Y)
If Button = vbRightButton Then
  'show popupmenu
  If LstItem Is Nothing Then
    For X = 1 To Lijst.ListItems.Count
      Lijst.ListItems(X).Selected = False
    Next
    Set Lijst.SelectedItem = Nothing
    PopupMenu mnuView(4), vbPopupMenuLeftButton Or vbPopupMenuRightButton
    Exit Sub
  Else
    If LstItem.Selected = False Then
      'if rightclicked on a non-selected file
      'deselect all items
      For X = 1 To Lijst.ListItems.Count
        Lijst.ListItems(X).Selected = False
      Next
    End If
    'and select the clicked one
    LstItem.Selected = True
  End If
  If Lijst.SelectedItem.Key = "" Then Exit Sub
  mnuFileText.Caption = "Actions for " + Lijst.SelectedItem.Text
  mnuSel(3).Visible = False
  If GetLijstCount > 1 Then
    mnuSel(3).Visible = True
    mnuSel(3).Caption = "Actions for" + Str(GetLijstCount) + " selected items"
  End If
  PopupMenu menuSel, vbPopupMenuLeftButton Or vbPopupMenuRightButton
  Lijst.Refresh
End If
If LstItem Is Nothing Then
  Set Lijst.SelectedItem = Nothing
End If
End Sub

Private Sub Lijst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Lst As ListItem, ToolTip$, Counter%
Set Lst = Lijst.HitTest(X, Y)
Lijst.ToolTipText = ""
If mnuOptions(4).Checked = False Then
  picPreview.Visible = False
End If
If mnuOptions(3).Checked = False Then Exit Sub
If Lst Is Nothing Then
  picPreview.Visible = False
  Exit Sub
End If
ToolTip = Trim(UCase(Lst.Text))
If Lst.SubItems(1) <> "" Then
  ToolTip = ToolTip + "   (" + Lst.SubItems(1) + ")"
Else
  If Lst.ForeColor <> vbRed Then
    ToolTip = ToolTip + "   (inside "
    For Counter = Lst.Index To 1 Step -1
      If Lijst.ListItems(Counter).SubItems(1) <> "" Then
        ToolTip = ToolTip + UCase(Lijst.ListItems(Counter)) + ")"
      End If
    Next
  End If
End If
If InStr("bmp gif jpg jpeg cur ico", LCase(Right(Lst.Key, 3))) <> 0 And mnuOptions(4).Checked And Lst.Key <> "" Then
  picPreview.Visible = True
  imgPreview = LoadPicture(Lst.Key)
Else
  picPreview.Visible = False
End If
If Trim(Lst.SubItems(2)) <> "" Then ToolTip = ToolTip + "; " + Lst.SubItems(2)
If Trim(Lst.SubItems(3)) <> "" Then ToolTip = ToolTip + "; " + Lst.SubItems(3)
Lijst.ToolTipText = ToolTip
End Sub

Private Sub Lijst_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lijst.HitTest(X, Y) Is Nothing Then
  If Not (Lijst.SelectedItem Is Nothing) Then Lijst.SelectedItem.Selected = False
End If
End Sub

Private Sub mnuEdit_Click(Index As Integer)
Dim X%
Select Case Index
  Case 0
    For X = 1 To Lijst.ListItems.Count
      If Lijst.ListItems(X).Key <> "" Then
        Lijst.ListItems(X).Selected = True
      End If
    Next
  Case 1
    For X = 1 To Lijst.ListItems.Count
      If Lijst.ListItems(X).Key <> "" Then
        Lijst.ListItems(X).Selected = Not (Lijst.ListItems(X).Selected)
      End If
    Next
End Select
End Sub

Private Sub mnuFile_Click(Index As Integer)
Dim SEI As SHELLEXECUTEINFO
On Error GoTo Fout
Select Case Index
  Case 0
  Case 2
    SaveFileSearchTask
  Case 3
    LoadFileSearchTask
  Case 5
    With SEI
      .cbSize = Len(SEI)
      .lpVerb = "find"
      .lpFile = ""
      .nShow = 5
    End With
    ShellExecuteEX SEI
  Case 7 Or Index = -1
    Unload Me
    Set Form1 = Nothing
    End
End Select
Exit Sub
Fout:
  If Err.Number = 32755 Then
  Else
    Resume Next
  End If
End Sub
Sub LoadFileSearchTask(Optional Path As String)
Dim Text$, X%
With cmd1
  FSTFileName = Path
  If Trim(Path) = "" Then
    .DialogTitle = "Open File Searcher Task"
    .FilterIndex = 0
    .Flags = cdlOFNExtensionDifferent ' Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt
    .ShowOpen
    FSTFileName = .Filename
  End If
  If VBA.dir(FSTFileName) = "" Then Exit Sub
  Close
  Open FSTFileName For Input As 1
  Do
    Input #1, Text
  Loop Until Trim(Text) <> ""
  If InStr(Text, AppName) = 0 Then
    MsgBox .Filename + " is not a valid File Searcher Task File.", vbCritical
    Exit Sub
  End If
  cmbName = SearchItemInTask("name")
  cmbPath.Text = SearchItemInTask("path")
  chkSubmaps.Value = SearchItemInTask("submaps")
  chkZip.Value = SearchItemInTask("zips")
  optDate(SearchItemInTask("optdate")).Value = True
  If optDate(1).Value Then
    optDateSpec(SearchItemInTask("optdatespec")).Value = True
  End If
  DTBetween1 = SearchItemInTask("optdatespec(0).date1")
  DTBetween2 = SearchItemInTask("optdatespec(0).date2")
  txtPrevMonth = SearchItemInTask("optDateSpec(1).txtPrevMonth")
  txtPrevDay = SearchItemInTask("optDateSpec(2).txtPrevday")
  cmbSize.ListIndex = SearchItemInTask("cmbsize")
  txtSize = SearchItemInTask("txtsize")
  For X = 0 To chkCheck.UBound
    chkCheck(X).Value = SearchItemInTask("chkcheck(" + Trim(Str(X)))
  Next
End With
End Sub
Private Function SearchItemInTask(Item As String) As String
Dim Text$
Close
Open FSTFileName For Input As 1
Do
  Input #1, Text
  If Left(Text, 1) <> "'" Then
    If InStr(LCase(Text), LCase(Item)) Then
       SearchItemInTask = LTrim(Mid(Text, InStr(Text, "=") + 1))
       Exit Do
    End If
  End If
Loop Until EOF(1)
Close
End Function
Sub SaveFileSearchTask(Optional Path As String)
Dim Sel%, X%
With cmd1
  If Trim(Path) = "" Then
    .DialogTitle = "Save File Searcher Task"
    .FilterIndex = 0
    .Flags = cdlOFNExtensionDifferent Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt
    .ShowSave
    Path = .Filename
  End If
  Close
  Open Path For Output As 1
  Print #1, AppName + " - File Searcher Task"
  Print #1, "'-- Name and Location --"
  Print #1, "name= " + cmbName
  Print #1, "path= " + cmbPath.Text
  Print #1, "submaps=" + Str(chkSubmaps.Value)
  Print #1, "zips=" + Str(chkZip.Value)
  Print #1, "'-- Date --"
  If optDate(0).Value Then
    Print #1, "optdate= 0"
  Else
    Sel = 2
    If optDateSpec(0).Value Then Sel = 0
    If optDateSpec(1).Value Then Sel = 1
    Print #1, "optdate= 1"
    Print #1, "  optdatespec=" + Str(Sel)
  End If
  Print #1, "  optdatespec(0).Date1= " + Str(DTBetween1)
  Print #1, "  optdatespec(0).Date2= " + Str(DTBetween2)
  Print #1, "  optdatespec(1).txtprevmonth= " + txtPrevMonth
  Print #1, "  optdatespec(2).txtprevday= " + txtPrevDay
  Print #1, "'-- Advanced --"
  Print #1, "cmbsize=" + Str(cmbSize.ListIndex)
  Print #1, "txtsize=" + Str(Val(txtSize))
  For X = 0 To chkCheck.UBound
    Print #1, "chkCheck(" + Trim(Str(X)) + ")=" + Str(chkCheck(X))
  Next
  Close
End With
End Sub
Private Sub mnuFileSave_Click(Index As Integer)
Dim P$, X%
If Lijst.ListItems.Count = 0 Then
  MsgBox "Nothing to save", vbExclamation
  Exit Sub
End If
Select Case Index
  Case 0
    P = "c:\windows\desktop\"
  Case 1
    P = RightPath(GetFolder("Save search results in ..."))
End Select
If Len(Trim(P)) < 3 Then Exit Sub
P = P + "search results.txt"
Close
Open P For Append As 1
Print #1, "Save created on " + Date$ + ", " + Time$
For X = 1 To Lijst.ListItems.Count
  If Lijst.ListItems(X).Key <> "" And Lijst.ListItems(X).ForeColor <> vbRed Then
    Print #1, Lijst.ListItems(X).Key
  End If
Next
Close
End Sub

Private Sub mnuHelp_Click()
Form2.Show 1
End Sub

Private Sub mnuOptions_Click(Index As Integer)
mnuOptions(Index).Checked = Not (mnuOptions(Index).Checked)
SaveSetting AppName, "Options", mnuOptions(Index).Tag, mnuOptions(Index).Checked
End Sub

Private Sub mnuSel_Click(Index As Integer)
If GetLijstCount = 0 Then Exit Sub
Select Case Index
  Case 0
    Lijst_DblClick
  Case 1
    Lijst.StartLabelEdit
  Case 2
    ShowProps
  Case 3
  Case 4
    MsgBox "Cut code needed"
  Case 5
    MsgBox "Copy code needed"
  Case 6
    EraseFile
  Case 8
    ShellExecute hwnd, "explore", RightPath(Left(Lijst.SelectedItem.Key, InStrRev(Lijst.SelectedItem.Key, "\") - 1)), "", RightPath(Left(Lijst.SelectedItem.Key, InStrRev(Lijst.SelectedItem.Key, "\") - 1)), 1
  Case 10
    picPreview = LoadPicture(Lijst.SelectedItem.Key)
    picPreview.Visible = True
End Select
End Sub

Private Sub mnuView_Click(Index As Integer)
Dim X%
If Index > 2 Then Exit Sub
For X = 0 To mnuView.UBound
  mnuView(X).Checked = False
Next
mnuView(Index).Checked = True
Tabs.Tabs(Index + 1).Selected = True
End Sub

Private Sub mnuViewList_Click(Index As Integer)
Dim X%
For X = 0 To mnuViewList.UBound
  mnuViewList(X).Checked = False
Next
mnuViewList(Index).Checked = True
Lijst.View = Index
SaveSetting AppName, "Options", "View", Index
End Sub

Private Sub optDate_Click(Index As Integer)
Dim X%
If Index = 0 Then
  optDateSpec_Click -1
  For X = 0 To optDateSpec.UBound
    optDateSpec(X).Value = False
  Next
End If
If Index = 1 Then
  optDateSpec(0).Value = True
  If optDateSpec(0).Value = True Then
    DTBetween1.Enabled = True
    DTBetween2.Enabled = True
  End If
End If
End Sub

Private Sub optDateSpec_Click(Index As Integer)
If Index <> -1 Then optDate(1).Value = True
If Index <> -1 Then optDateSpec(Index).Value = True
DTBetween1.Enabled = (Index = 0)
DTBetween2.Enabled = (Index = 0)
txtPrevMonth.Enabled = (Index = 1)
UDtxtPrevMonth.Enabled = txtPrevMonth.Enabled
txtPrevDay.Enabled = (Index = 2)
UDtxtPrevDay.Enabled = txtPrevDay.Enabled
SendKeys vbTab
End Sub

Private Sub picPreview_Click()
picPreview.Visible = False
End Sub

Private Sub Tabs_Click()
Static PrevIndex%
Frame1(PrevIndex).Visible = False
Frame1(Tabs.SelectedItem.Index - 1).Visible = True
mnuView(PrevIndex).Checked = False
PrevIndex = Tabs.SelectedItem.Index - 1
mnuView(PrevIndex).Checked = True
End Sub
Sub SearchInSub(ByVal Folder As String)
'recursive search in a folder
Dim Files As String, Folders As String
Dim X%, Lst As ListItem, Size As Double, SizeFormat$
Dim LstImg As ListImage
If StopClicked Then 'cancel the search?
  Exit Sub
End If
On Error GoTo Err

Folder = RightPath(Folder) 'correct the folder

If GetMapNiveau(Folder) < 4 Then status.Panels(1).Text = Folder  'set only if 4 niveaus

If Folder = "" Then Exit Sub

dir.Path = RightPath(Folder) 'set the folder-->update file
If chkCheck(0).Value = 0 Then
  For X = 0 To file.ListCount - 1
    'check every file
    If IsCorrectFile(Folder + file.List(X), DateValue(FileDateTime(Folder + file.List(X))), FileLen(Folder + file.List(X))) Then
      'file is ok, with all chosen settings
      'add to the list
      Set Lst = Lijst.ListItems.Add(, Folder + file.List(X), file.List(X))
      With Lst
        .SubItems(1) = RightPath(Folder)
        .SubItems(2) = FormatSize(FileLen(Folder + file.List(X)))
        .SubItems(4) = DateValue(FileDateTime(Folder + file.List(X)))
        .SubItems(3) = GetFileType(file.List(X))
        SetIconForFile file.List(X), Lst
'        If ExtractPicture(file.List(X)) Then
'          imglstListViewSmall.ListImages.Add , , pic.Picture
'          imglstListViewBig.ListImages.Add , , pic.Picture
''          Stop
'          .Icon = imglstListViewBig.ListImages.Count
'          .SmallIcon = imglstListViewSmall.ListImages.Count
'        End If
      End With
      ItemsFound = ItemsFound + 1
    End If
  Next
  
  If chkZip.Value Then
    'search in zipfiles, if option is selected
    For X = 0 To fileZip.ListCount - 1
      status.Panels(2) = fileZip.List(X) + "(" + Trim(Str(Round(FileLen(Folder + fileZip.List(X)) / 1024 / 1024))) + " Mb)"
      SearchInZip Folder + fileZip.List(X) 'search for matching files in zip
      If StopClicked Then Exit Sub
    Next
  End If
End If
status.Panels(2) = ""

dir.Path = Folder
For X = 0 To dir.ListCount - 1
  Folders = RightPath(dir.List(X))
  Folders = Mid(Folders, InStrRev(Folders, "\", Len(Folders) - 1) + 1)
  Folders = Left(Folders, Len(Folders) - 1)
  If LCase(Folders) Like "*" + LCase(cmbName.Text) + "*" Then
    Set Lst = Lijst.ListItems.Add(, RightPath(dir.List(X)), Folders, "folder", "folder")
    Lst.SubItems(1) = ParsePath(dir.List(X))
    Lst.SubItems(3) = "Folder"
    ItemsFound = ItemsFound + 1
    
  End If
Next
If chkSubmaps.Value = 0 Then Exit Sub  'search in subdirectories?
If Unloading Then Exit Sub

For X = 0 To dir.ListCount - 1
  dir.Path = Folder
  SearchInSub dir.List(X) 'recursive search in every folder
Next
DoEvents
Exit Sub

Err:
  If Err = 68 Then
    Lijst.ListItems.Add , , "Error: " + "Drive " + UCase(Left(Folder, 2)) + " not ready"
    Lijst.ListItems(Lijst.ListItems.Count).ForeColor = vbRed
    Exit Sub
  ElseIf Err = 76 Then
    Lijst.ListItems.Add , , "Error: " + Err.Description + "(" + Folder + ")"
    Lijst.ListItems(Lijst.ListItems.Count).ForeColor = vbRed
    Exit Sub
  Else
    Lijst.ListItems.Add , , "Error: " + Err.Description
  End If
  Lijst.ListItems(Lijst.ListItems.Count).ForeColor = vbRed
  Lijst.ListItems(Lijst.ListItems.Count).EnsureVisible
'  Stop
  Resume Next
End Sub
Sub SetIconForFile(Ext As String, LstItem As ListItem)
If mnuOptions(0).Checked = False Then Exit Sub
On Error GoTo Err
Dim PicPath$, IconInList As Boolean, Index%
'search for the defaulticon in the register, and draw it on Pic
PicPath = ExtractPicture(Ext)
If PicPath <> "" Then
  'icon found
  IconInList = True
  'check if the icon is already in the list; if not it gives an error
  'and IconInList will be set to false, if yes, just continue
  'this is done to reduce the amount of memory used by the icons
  Index = imglstListViewSmall.ListImages(PicPath).Index
  If IconInList = False Then
    'icon is not present, add it to the imagelists
    imglstListViewSmall.ListImages.Add , PicPath, pic.Picture
    imglstListViewBig.ListImages.Add , PicPath, pic.Picture
    'set icons on listview
    LstItem.Icon = imglstListViewBig.ListImages.Count
    LstItem.SmallIcon = imglstListViewSmall.ListImages.Count
  Else
    'icon is already in listimage
    'set icons on listview
    LstItem.Icon = Index
    LstItem.SmallIcon = Index
  End If
Else
  LstItem.Icon = imglstListViewBig.ListImages("unknown").Index
  LstItem.SmallIcon = imglstListViewSmall.ListImages("unknown").Index
End If
Exit Sub
Err:
  If Err.Number = 35601 Then
    'icon is not found in the imagelist
    IconInList = False
    Resume Next
  Else
    Lijst.ListItems.Add , , "Error: " + Err.Description
  End If
End Sub
Function IsCorrectFile(file As String, Datum As String, FileLength As Double) As Boolean
'is the file ok? (=good date and size?)
Dim L As Double
On Error GoTo Fout

'date options set?
If optDate(1).Value Then
  If IsCorrectDate(Datum) = False Then
    'no good date
    Exit Function
  End If
End If
'size options set?
If cmbSize.ListIndex <> 0 Then
  L = Round(FileLength / 1024, 2) 'convert length in Kb
  If cmbSize.ListIndex = 1 Then 'minimum
    If L <= Val(txtSize) Then
      Exit Function
    End If
  End If
  If cmbSize.ListIndex = 2 Then 'maximum
    If L >= Val(txtSize) Then
      Exit Function
    End If
  End If
  If cmbSize.ListIndex = 3 Then 'equal to=1 Kb less, or 1 Kb more
    If L <= Val(txtSize - 1) Or L >= Val(txtSize + 1) Then
      Exit Function
    End If
  End If
End If
IsCorrectFile = True

Exit Function

Fout:
  Lijst.ListItems.Add , , "Error:" + Err.Description
  Lijst.ListItems(Lijst.ListItems.Count).ForeColor = vbRed
End Function
Function IsCorrectDate(Datum As String) As Boolean
'is the date in correspondance with the datesettings?
Dim OK As Boolean
If optDateSpec(0).Value Then 'between 2 dates
  If DateDiff("d", DTBetween1, Datum) >= 0 And DateDiff("d", DTBetween2, Datum) <= 0 Then
    OK = True
  End If
End If
If optDateSpec(1).Value Then 'x months ago
  If Abs(DateDiff("m", Date, Datum)) <= Val(txtPrevMonth) Then
    OK = True
  End If
End If
If optDateSpec(2).Value Then 'x days ago
  If Abs(DateDiff("d", Date, Datum)) <= Val(txtPrevDay) Then
    OK = True
  End If
End If
IsCorrectDate = OK
End Function

Function GetMapNiveau(Folder As String) As Integer
'how many niveaus?
Dim X%
GetMapNiveau = -1
For X = 1 To Len(Folder)
  X = InStr(X, Folder, "\")
  GetMapNiveau = GetMapNiveau + 1
Next
End Function
Sub SearchInZip(ZipFile As String)
'search in a zipfile (only 1 level deep)
Dim X%, Lst As ListItem, ItemAdded As Boolean, Index%

ItemAdded = True
On Error GoTo ItemIsAdded
Index = Lijst.ListItems(ZipFile).Index + 1

On Error GoTo Fout

Filename = ZipFile 'read the contents of the zip

If StopClicked Then Exit Sub

For X = 1 To Archive.Count
  If FileMatch(ParseFilename(Archive.Item(X).Filename)) Then  'a filematch?
    If Archive(X).FileDateTime <> "" Then
      If IsCorrectFile(file, DateValue(Archive(X).FileDateTime), Archive.Item(X).UncompressedSize) Then
        If Not (ItemAdded) Then
          'add the zip, if it is not added yet
          Set Lst = Lijst.ListItems.Add(, ZipFile, ParseFilename(ZipFile))
          Lst.SubItems(1) = ParsePath(ZipFile)
          Lst.SubItems(2) = FormatSize(FileLen(ZipFile))
          Lst.SubItems(3) = GetFileType(ZipFile)
          Lst.SubItems(4) = DateValue(FileDateTime(ZipFile))
          ItemAdded = True
          SetIconForFile ZipFile, Lst
        End If
        ItemsFound = ItemsFound + 1
        Set Lst = Lijst.ListItems.Add(Index, , String(10, " ") + Archive.Item(X).Filename)
        Lst.SubItems(2) = FormatSize(Archive.Item(X).UncompressedSize)
        Lst.SubItems(3) = GetFileType(Archive(X).Filename)
        Lst.SubItems(4) = DateValue(Archive(X).FileDateTime)
        Index = Index + 1
        Lst.ForeColor = RGB(100, 100, 100)
        SetIconForFile Archive(X).Filename, Lst
      End If
    End If
  End If
  If Right(X, 1) = "0" Then 'every 10 loops, a Doevents and check for a Stop Clicked
    DoEvents
    If StopClicked Or Unloading Then Exit Sub
  End If
Next
Exit Sub
Fout:
  Lijst.ListItems.Add , , "Error in zip(" + ParseFilename(ZipFile) + "):" + Err.Description
  Lijst.ListItems(Lijst.ListItems.Count).ForeColor = vbRed
  Resume Next
ItemIsAdded:
  ItemAdded = False
  Index = Lijst.ListItems.Count + 2
  Resume Next
End Sub
Function FileMatch(file As String) As Boolean
'does the filematch correspond (with jokers)
FileMatch = LCase(file) Like LCase(Me.file.Pattern)
End Function
Function RightPath(Path As String)
'correct path
RightPath = Path
If Right(Path, 1) <> "\" Then
  RightPath = Path + "\"
End If
End Function
Function GetDrive(DriveName As String) As String
'get drive type, and set the name
Dim Drv As Long
DriveName = Left(DriveName, 2)
Drv = GetDriveType(UCase(RightPath(DriveName)))
Select Case Drv
  Case DRIVE_FIXED
    GetDrive = "drive"
  Case DRIVE_CDROM
    GetDrive = "cdrom"
  Case DRIVE_REMOVABLE
    GetDrive = "floppy"
End Select
End Function
Function FormatSize(ByVal FileSize As Double) As String
'format the size to the most logical
Dim SizeString$
SizeString = "b"
If FileSize > 1024 Then
  FileSize = FileSize / 1024
  SizeString = "Kb"
End If
If FileSize > 1024 Then
  FileSize = FileSize / 1024
  SizeString = "Mb"
End If
FileSize = Round(FileSize, 2)
FormatSize = Trim(Str(FileSize)) + " " + SizeString
End Function
Private Function GetFolder(ByVal sTitle As String) As String
'get folder
Dim bInf As BROWSEINFO
Dim RetVal As Long
Dim PathID As Long
Dim RetPath As String
Dim Offset As Integer
'Set the properties of the folder dialog
bInf.hOwner = Me.hwnd
bInf.lpszTitle = sTitle
bInf.ulFlags = BIF_RETURNONLYFSDIRS
'Show the Browse For Folder dialog
PathID = SHBrowseForFolder(bInf)
RetPath = Space$(512)
RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
If RetVal Then
  'Trim off the null chars ending the path
  'and display the returned folder
  Offset = InStr(RetPath, Chr$(0))
  GetFolder = Left$(RetPath, Offset - 1)
End If
End Function

Private Sub EraseFile()
'delete a file, according to the windows delete
On Error GoTo Fout
Dim intResult As Integer
Dim filop As SHFILEOPSTRUCT
Dim Filename$, X%, Item$
'simulate behavior of windows explorer
'if shift is pressed, don't send to recycle bin
status.Panels(1) = "Deleting selected files ..."
For X = 1 To Lijst.ListItems.Count
  If X <= Lijst.ListItems.Count Then
    If Lijst.ListItems(X).Selected Then
      Item = Lijst.ListItems(X).Key
      If Right(Item, 1) = "\" Then Item = Left(Item, Len(Item) - 1)
      Filename = Filename + Item + Chr(0)
      ItemsFound = ItemsFound - 1
    End If
  End If
Next
Filename = Filename + Chr(0)
If GetKeyState(VK_SHIFT) < 0 Then
  With filop
    .wFunc = FO_DELETE
    .pFrom = Filename
  End With
  SHFileOperation filop
Else
  With filop
    .fFlags = FOF_ALLOWUNDO  'send to recycle bin
    .wFunc = FO_DELETE
    .pFrom = Filename
  End With
  SHFileOperation filop
End If
If filop.fAnyOperationsAborted Then Exit Sub
status.Panels(1) = Trim(Str(ItemsFound)) + " item" + IIf(ItemsFound <> 1, "s", "") + " found" + IIf(StopClicked, "(search cancelled)", "")
For X = 1 To Lijst.ListItems.Count
  If X <= Lijst.ListItems.Count Then
    If Lijst.ListItems(X).Selected Then
      Lijst.ListItems.Remove X
      X = X - 1
    End If
  End If
Next
Exit Sub
Fout:
End Sub
Function GetLijstCount() As Integer
Dim X%
For X = 1 To Lijst.ListItems.Count
  If Lijst.ListItems(X).Selected Then GetLijstCount = GetLijstCount + 1
Next
End Function

Private Sub txtPrevDay_GotFocus()
txtPrevDay.SelStart = 0
txtPrevDay.SelLength = Len(txtPrevDay)
End Sub

Private Sub txtPrevMonth_GotFocus()
txtPrevMonth.SelStart = 0
txtPrevMonth.SelLength = Len(txtPrevMonth)
End Sub

Private Sub txtSize_Change()
cmbSize_Click
End Sub

Private Sub txtSize_GotFocus()
txtSize.SelStart = 0
txtSize.SelLength = Len(txtSize)
End Sub
