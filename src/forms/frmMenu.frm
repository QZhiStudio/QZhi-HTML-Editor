VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMenu 
   Caption         =   "QZhi HTML Editor"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12555
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   12555
   StartUpPosition =   3  '窗口缺省
   Begin ComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ilsMain"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   16
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Object.ToolTipText     =   "新建"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Open"
            Object.ToolTipText     =   "打开"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.ToolTipText     =   "保存"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "SaveAs"
            Object.ToolTipText     =   "另存为"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "PageSetup"
            Object.ToolTipText     =   "页面设置"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "PrintPreview"
            Object.ToolTipText     =   "打印预览"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            Object.ToolTipText     =   "打印"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Undo"
            Object.ToolTipText     =   "撤消"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Redo"
            Object.ToolTipText     =   "重做"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "SelectAll"
            Object.ToolTipText     =   "全选"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Cut"
            Object.ToolTipText     =   "剪切"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Copy"
            Object.ToolTipText     =   "复制"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Paste"
            Object.ToolTipText     =   "粘贴"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin ComctlLib.Toolbar tlbBrowse 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ilsBrowse"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "GoBack"
            Object.ToolTipText     =   "后退"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "GoForward"
            Object.ToolTipText     =   "前进"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "刷新"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Stop"
            Object.ToolTipText     =   "停止"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Home"
            Object.ToolTipText     =   "主页"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ZoomIn"
            Object.ToolTipText     =   "放大"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ZoomOut"
            Object.ToolTipText     =   "缩小"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ZoomOriginalSize"
            Object.ToolTipText     =   "缩放至正常大小"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin VB.ComboBox cboZoom 
         Height          =   300
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   45
         Width           =   3135
      End
   End
   Begin ComctlLib.Toolbar tlbFormat 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ilsFormat"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   20
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Bold"
            Object.ToolTipText     =   "粗体"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Italic"
            Object.ToolTipText     =   "斜体"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Underline"
            Object.ToolTipText     =   "下划线"
            Object.Tag             =   ""
            ImageIndex      =   17
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "StrikeThrough"
            Object.ToolTipText     =   "删除线"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Subscript"
            Object.ToolTipText     =   "下标"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Superscript"
            Object.ToolTipText     =   "上标"
            Object.Tag             =   ""
            ImageIndex      =   16
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   3600
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ForegroundColor"
            Object.ToolTipText     =   "前景色"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "BackgroundColor"
            Object.ToolTipText     =   "背景色"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "JustifyLeft"
            Object.ToolTipText     =   "左对齐"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "JustifyCenter"
            Object.ToolTipText     =   "居中"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "JustifyRight"
            Object.ToolTipText     =   "右对齐"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "JustifyFull"
            Object.ToolTipText     =   "两端对齐"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "JustifyNone"
            Object.ToolTipText     =   "无对齐"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Outdent"
            Object.ToolTipText     =   "减少缩进量"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Indent"
            Object.ToolTipText     =   "增加缩进量"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "CreateHyperlink"
            Object.ToolTipText     =   "创建超链接"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
      EndProperty
      Begin VB.ComboBox cboFontSize 
         Height          =   300
         Left            =   4680
         TabIndex        =   5
         Text            =   "cboFontSize"
         Top             =   45
         Width           =   975
      End
      Begin VB.ComboBox cboFontName 
         Height          =   300
         Left            =   2280
         TabIndex        =   4
         Text            =   "cboFontName"
         Top             =   45
         Width           =   2310
      End
   End
   Begin SHDocVwCtl.WebBrowser brwWebControl 
      Height          =   510
      Left            =   9000
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2400
      Width           =   510
      ExtentX         =   900
      ExtentY         =   900
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
   Begin VB.Timer tmrQueryDocStat 
      Interval        =   40
      Left            =   12120
      Top             =   1920
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   9720
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   12000
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgFGColor 
      Height          =   240
      Left            =   11040
      Picture         =   "frmMenu.frx":4072
      Top             =   2040
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgBGColor 
      Height          =   240
      Left            =   10680
      Picture         =   "frmMenu.frx":4119
      Top             =   2040
      Visible         =   0   'False
      Width           =   240
   End
   Begin ComctlLib.ImageList ilsTemp 
      Left            =   11400
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList ilsFormat 
      Left            =   11400
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   17
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":41BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":450F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":4861
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":4BB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":4F05
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":5257
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":55A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":58FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":5C4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":5F9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":62F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":6643
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":6995
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":6CE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":7039
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":76DD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ilsBrowse 
      Left            =   10800
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":7A2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":7D81
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":80D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":8425
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":8777
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":8AC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":8E1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":916D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ilsMain 
      Left            =   10200
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":94BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":9811
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":9B63
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":9EB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":A207
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":A559
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":A8AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":ABFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":AF4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":B2A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":B5F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":B945
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenu.frx":BC97
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileNew 
         Caption         =   "新建(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "打开(&O)..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "另存为(&A)..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "页面设置(&U)"
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "撤消(&U)"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "重做(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "全选(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "剪切(&T)"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "复制(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "粘贴(&P)"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "工具(&T)"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "内容(&C)..."
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于 QZhi HTML Editor(&A)"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 2024 QZhi Studio

' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.

Option Explicit

Private lngCurrentBGColor As Long
Private lngCurrentFGColor As Long

Private lngNewBGColor As Long
Private lngNewFGColor As Long

Private blnIsInitialized As Boolean

Private hBGColorButtonIcon As Long
Private hFGColorButtonIcon As Long

Public brwWebControl_Initialized As Boolean

Private Sub brwWebControl_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    brwWebControl_Initialized = True
End Sub

Private Sub brwWebControl_DownloadBegin()
    brwWebControl.Silent = True
End Sub

Private Sub brwWebControl_DownloadComplete()
    brwWebControl.Silent = True
End Sub

Private Sub cboFontName_Click()
    frmCode.hdocHTMLDesignDocument.execCommand "FontName", False, cboFontName.List(cboFontName.ListIndex)
End Sub

Private Sub cboFontSize_Click()
    frmCode.hdocHTMLDesignDocument.execCommand "FontSize", False, cboFontSize.List(cboFontSize.ListIndex)
End Sub

Private Sub cboZoom_Click()
    Dim vntZoom As Variant
    If App.LogMode = 0 Or glngIEVersion < 7 Then    ' 在 IDE 中或 IE 版本低于 7
        Select Case cboZoom.List(cboZoom.ListIndex)
            Case "最小"
                vntZoom = CLng(0)
            Case "较小"
                vntZoom = CLng(1)
            Case "中"
                vntZoom = CLng(2)
            Case "较大"
                vntZoom = CLng(3)
            Case "最大"
                vntZoom = CLng(4)
        End Select
        frmCode.brwQuickView.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, vntZoom, 0
    ElseIf glngIEVersion >= 7 Then
        '.Document.Body.Style = "zoom:100%"
        vntZoom = CLng(Replace(cboZoom.List(cboZoom.ListIndex), "%", ""))
        frmCode.brwQuickView.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DODEFAULT, vntZoom, 0
    End If
End Sub

Private Sub Form_Load()
    LoadFonts
    frmMsg.PrintLog Time, "App.StatusText", "字体加载成功"
    frmMsg.PrintLog Time, "App.StatusText", "主界面已初始化"

    brwWebControl.Move -2400, -2400 ' 该控件 Visible 不可设置为 False

    brwWebControl.Navigate "about:blank"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = Not CloseCurrentFile
    If Cancel = False Then
        AtExit
    End If
End Sub

Private Sub mnuEditCopy_Click()
    If gapmMode = AppDesignMode Then
        frmCode.brwDesign.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT
    ElseIf gapmMode = AppEditMode Then
        frmCode.brwSource.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT
    End If
End Sub

Private Sub mnuEditCut_Click()
    If gapmMode = AppDesignMode Then
        frmCode.brwDesign.ExecWB OLECMDID_CUT, OLECMDEXECOPT_DODEFAULT
    ElseIf gapmMode = AppEditMode Then
        frmCode.brwSource.ExecWB OLECMDID_CUT, OLECMDEXECOPT_DODEFAULT
    End If
End Sub

Private Sub mnuEditPaste_Click()
    If gapmMode = AppDesignMode Then
        frmCode.brwDesign.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DODEFAULT
    ElseIf gapmMode = AppEditMode Then
        frmCode.brwSource.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DODEFAULT
    End If
End Sub

Private Sub mnuEditRedo_Click()
    If gapmMode = AppDesignMode Then
        frmCode.brwDesign.ExecWB OLECMDID_REDO, OLECMDEXECOPT_DODEFAULT
    ElseIf gapmMode = AppEditMode Then
        frmCode.brwSource.ExecWB OLECMDID_REDO, OLECMDEXECOPT_DODEFAULT
    End If
End Sub

Private Sub mnuEditSelectAll_Click()
    If gapmMode = AppDesignMode Then
        frmCode.brwDesign.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT
    ElseIf gapmMode = AppEditMode Then
        frmCode.eEditor.Select
    End If
End Sub

Private Sub mnuEditUndo_Click()
    If gapmMode = AppDesignMode Then
        frmCode.brwDesign.ExecWB OLECMDID_UNDO, OLECMDEXECOPT_DODEFAULT
    ElseIf gapmMode = AppEditMode Then
        frmCode.brwSource.ExecWB OLECMDID_UNDO, OLECMDEXECOPT_DODEFAULT
    End If
End Sub

Private Sub mnuFileExit_Click()
    Dim blnRet As Boolean
    blnRet = CloseCurrentFile
    If blnRet = True Then Unload Me
End Sub

Private Sub mnuFileNew_Click()
    CreateNewFile
End Sub

Private Sub mnuFileOpen_Click()
    OpenNewFile
End Sub

Private Sub mnuFilePageSetup_Click()
    brwWebControl.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub mnuFilePrint_Click()
    brwWebControl_Initialized = False
    brwWebControl.Navigate "about:blank"

    While brwWebControl_Initialized = False
        DoEvents
    Wend

    If gapmMode = AppEditMode Then
        WriteHTML brwWebControl, "<font face=""Courier New"">" & StringtoEntity(frmCode.eEditor.Value) & "</font>"
    Else
        WriteHTML brwWebControl, GetCurrentDocHTML
    End If
    
    While brwWebControl_Initialized = False
        DoEvents
    Wend
    
    brwWebControl.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub mnuFilePrintPreview_Click()

    brwWebControl_Initialized = False
    brwWebControl.Navigate "about:blank"

    While brwWebControl_Initialized = False
        DoEvents
    Wend
    
    If gapmMode = AppEditMode Then
        WriteHTML brwWebControl, "<font face=""Courier New"">" & StringtoEntity(frmCode.eEditor.Value) & "</font>"
    Else
        brwWebControl_Initialized = False
        WriteHTML brwWebControl, GetCurrentDocHTML
        
    End If
    
    While brwWebControl_Initialized = False
        DoEvents
    Wend

    brwWebControl.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER

    
End Sub

Private Sub mnuFileSave_Click()
    SaveCurrentFile
End Sub

Private Sub mnuFileSaveAs_Click()
    SaveCurrentFileAs
End Sub

Private Sub mnuHelpAbout_Click()
    Dim strText As String
    Randomize
    Select Case Int((7 - 0 + 1) * Rnd + 0)
        Case 1, 6
            strText = "越过长城，走向世界。"
            
        Case 2
            strText = "Across the Great Wall we can reach every corner in the world."
        
        Case 3
            strText = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    End Select
    
    If strText <> "" Then strText = strText & vbCrLf
    strText = strText & "当前浏览器内核：Microsoft" & ChrW(&HAE) & " Internet Explorer" & ChrW(&HAE) & " " & glngIEVersion
    
    ShellAboutW Me.hwnd, StrPtr(App.ProductName), StrPtr(strText), Me.Icon
End Sub

Private Sub mnuHelpContents_Click()
    If App.LogMode = 1 Then
        HtmlHelpW Me.hwnd, StrPtr(App.HelpFile), &H0, 0
    End If
End Sub

Private Function LoadFonts()
    Dim i As Long
    Dim lngStep As Long
    
    For i = 0 To Screen.FontCount - 1
        cboFontName.AddItem Screen.Fonts(i)
    Next
    
    For i = 1 To 7
        cboFontSize.AddItem CStr(i)
    Next i
End Function

Private Sub tlbBrowse_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "GoBack"
            frmCode.brwQuickView.GoBack
            
        Case "GoForward"
            frmCode.brwQuickView.GoForward
            
        Case "Refresh"
            frmCode.brwQuickView.Refresh
            
        Case "Stop"
            frmCode.brwQuickView.Stop
            
        Case "Home"
            frmCode.brwQuickView.GoHome
            
        Case "ZoomIn"
            If cboZoom.ListIndex < cboZoom.ListCount - 1 Then cboZoom.ListIndex = cboZoom.ListIndex + 1
            cboZoom_Click
            
        Case "ZoomOut"
            If cboZoom.ListIndex <> 0 Then cboZoom.ListIndex = cboZoom.ListIndex - 1
            cboZoom_Click
            
        Case "ZoomOriginalSize"
            Dim i As Long
            If App.LogMode = 0 Or glngIEVersion < 7 Then    ' 在 IDE 中或 IE 版本低于 7
                cboZoom.ListIndex = CInt(glngDefaultZoom - gintMinZoom)
            Else
                For i = 0 To cboZoom.ListCount - 1
                    If cboZoom.List(i) = "100%" Then cboZoom.ListIndex = i
                Next i
            End If
            cboZoom_Click
    End Select
End Sub

Private Sub tlbFormat_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "Bold", "Italic", "StrikeThrough", "Subscript", "Superscript", _
            "Underline", "JustifyCenter", "JustifyFull", "JustifyLeft", "JustifyNone", "JustifyRight", _
            "Indent", "Outdent"
            
            frmCode.hdocHTMLDesignDocument.execCommand Button.Key, False, 0
            
        Case "CreateHyperlink"
            frmCode.hdocHTMLDesignDocument.execCommand "CreateLink", True, 0
            
        Case "BackgroundColor"
            UpdateBackgroundColor
            
        Case "ForegroundColor"
            UpdateForegroundColor
    End Select
    
    ' 这一行非常重要，强制重绘必须这样写，不能用 .refresh 代替
    tlbFormat.buttons(Button.Key).Image = tlbFormat.buttons(Button.Key).Image
End Sub

Private Function UpdateBackgroundColor()
    Dim clsCP As New clsColorPicker
    Dim lngColorRef As Long
    
    lngColorRef = clsCP.GetColor(lngCurrentBGColor, lngNewBGColor)
    
    If lngColorRef <> -1 Then
        lngNewBGColor = lngColorRef
        frmCode.hdocHTMLDesignDocument.execCommand "BackColor", False, lngColorRef
    End If
    
    UpdateBGColorButton
    
    Set clsCP = Nothing
End Function

Private Function UpdateForegroundColor()
    Dim clsCP As New clsColorPicker
    Dim lngColorRef As Long
    
    lngColorRef = clsCP.GetColor(lngCurrentFGColor, lngNewFGColor)
    
    If lngColorRef <> -1 Then
        lngNewFGColor = lngColorRef
        frmCode.hdocHTMLDesignDocument.execCommand "ForeColor", False, lngColorRef
    End If
    
    UpdateFGColorButton
    
    Set clsCP = Nothing
End Function

Public Function SetToolbarStat()

    Dim i As Long

    With gdsDocStat
        tlbFormat.buttons("Bold").Value = CBool(.vntBold) And &H1
        tlbFormat.buttons("Italic").Value = CBool(.vntItalic) And &H1
        tlbFormat.buttons("StrikeThrough").Value = CBool(.vntStrikeThrough) And &H1
        tlbFormat.buttons("Subscript").Value = CBool(.vntSubscript) And &H1
        tlbFormat.buttons("Superscript").Value = CBool(.vntSuperscript) And &H1
        tlbFormat.buttons("Underline").Value = CBool(.vntUnderline) And &H1
        
        tlbFormat.buttons("JustifyCenter").Value = CBool(.vntJustifyCenter) And &H1
        tlbFormat.buttons("JustifyFull").Value = CBool(.vntJustifyFull) And &H1
        tlbFormat.buttons("JustifyLeft").Value = CBool(.vntJustifyLeft) And &H1
        tlbFormat.buttons("JustifyNone").Value = CBool(.vntJustifyNone) And &H1
        tlbFormat.buttons("JustifyRight").Value = CBool(.vntJustifyRight) And &H1
        
        If IsNull(.vntFontName) = False Then
            cboFontName.Text = .vntFontName
        Else
            cboFontName.Text = ""
        End If
        
        If IsNull(.vntFontSize) = False Then
            cboFontSize.Text = .vntFontSize
        Else
            cboFontSize.Text = ""
        End If
        
        tlbBrowse.buttons("GoBack").Enabled = .blnNavigateBack
        tlbBrowse.buttons("GoForward").Enabled = .blnNavigateForward
    End With
    
End Function

Private Function UpdateBGColorButton()

    Dim lngCLRTemp As Long
    
    picTemp.BackColor = RGB(255, 0, 255)
    ilsTemp.MaskColor = RGB(255, 0, 255)

    If lngNewBGColor = RGB(255, 0, 255) Then
        lngCLRTemp = RGB(255, 0, 254)
    Else
        lngCLRTemp = lngNewBGColor
    End If

    picTemp.PaintPicture imgBGColor.Picture, 0, 0
    picTemp.Line (1, 7)-(8, 14), lngCLRTemp, BF
    
    picTemp.Refresh

    ilsTemp.ListImages.Add 1, , picTemp.Image

    hBGColorButtonIcon = ImageList_GetIcon(ilsTemp.hImageList, 0, 0)
    If hBGColorButtonIcon <> 0 Then
        ImageList_ReplaceIcon ilsFormat.hImageList, 0, hBGColorButtonIcon
    End If
    
    ilsTemp.ListImages.Remove 1

    tlbFormat.Refresh
    
End Function

Private Function UpdateFGColorButton()
    Dim lngCLRTemp As Long
    
    picTemp.BackColor = RGB(255, 0, 255)
    ilsTemp.MaskColor = RGB(255, 0, 255)

    If lngNewFGColor = RGB(255, 0, 255) Then
        lngCLRTemp = RGB(255, 0, 254)
    Else
        lngCLRTemp = lngNewFGColor
    End If

    picTemp.PaintPicture imgFGColor.Picture, 0, 0
    picTemp.Line (3, 8)-(7, 12), lngCLRTemp, BF
    
    picTemp.Refresh

    ilsTemp.ListImages.Add 1, , picTemp.Image

    hFGColorButtonIcon = ImageList_GetIcon(ilsTemp.hImageList, 0, 0)
    If hFGColorButtonIcon <> 0 Then
        ImageList_ReplaceIcon ilsFormat.hImageList, 3, hFGColorButtonIcon
    End If
    
    ilsTemp.ListImages.Remove 1

    tlbFormat.Refresh
End Function

Private Sub tlbMain_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "New"
            mnuFileNew_Click
        
        Case "Open"
            mnuFileOpen_Click
            
        Case "Save"
            mnuFileSave_Click
            
        Case "SaveAs"
            mnuFileSaveAs_Click
            
        Case "PageSetup"
            mnuFilePageSetup_Click
        
        Case "PrintPreview"
            mnuFilePrintPreview_Click
            
        Case "Print"
            mnuFilePrint_Click
            
        Case "Cut"
            mnuEditCut_Click
            
        Case "Copy"
            mnuEditCopy_Click
            
        Case "Paste"
            mnuEditPaste_Click
            
        Case "SelectAll"
            mnuEditSelectAll_Click
            
        Case "Undo"
            mnuEditUndo_Click
            
        Case "Redo"
            mnuEditRedo_Click
    End Select
End Sub

Private Sub tmrQueryDocStat_Timer()
    QueryDocStat
    
    If blnIsInitialized = False Then
        If IsNull(gdsDocStat.vntBackgroundColor) = False Then
            lngNewBGColor = CLng(gdsDocStat.vntBackgroundColor)
        Else
            lngNewBGColor = vbWhite
        End If
        If IsNull(gdsDocStat.vntForegroundColor) = False Then
            lngNewFGColor = CLng(gdsDocStat.vntForegroundColor)
        Else
            lngNewFGColor = vbBlack
        End If
        
        UpdateBGColorButton
        UpdateFGColorButton
        blnIsInitialized = True
    End If
    
    If IsNull(gdsDocStat.vntBackgroundColor) = False Then
        lngCurrentBGColor = CLng(gdsDocStat.vntBackgroundColor)
    Else
        lngCurrentBGColor = vbWhite
    End If
    If IsNull(gdsDocStat.vntForegroundColor) = False Then
        lngCurrentFGColor = CLng(gdsDocStat.vntForegroundColor)
    Else
        lngCurrentFGColor = vbBlack
    End If
    
    SetToolbarStat
End Sub
