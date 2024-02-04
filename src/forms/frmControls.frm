VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Begin VB.Form frmControls 
   Caption         =   "控件"
   ClientHeight    =   9015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   ControlBox      =   0   'False
   Icon            =   "frmControls.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin ComctlLib.ListView lvwControls 
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   14208
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ilsControls"
      SmallIcons      =   "ilsControls"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ImageList ilsControls 
      Left            =   3840
      Top             =   8280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmControls.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmControls.frx":035E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmControls.frx":06B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmControls.frx":0A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmControls.frx":0D54
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmControls.frx":10A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmControls.frx":13F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmControls.frx":174A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmControls.frx":1A9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmControls.frx":1DEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmControls.frx":2140
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmControls.frx":2492
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    With lvwControls
        .LabelEdit = lvwManual
        .ListItems.Clear
        .ColumnHeaders.Clear
        .view = lvwList
        
        .ListItems.Add , "Pointer", "Pointer", 1, 1
        .ListItems.Add , "Button", "Button", 2, 2
        .ListItems.Add , "CheckBox", "CheckBox", 3, 3
        .ListItems.Add , "Dropdown", "Dropdown", 9, 9
        .ListItems.Add , "FieldSet", "FieldSet", 4, 4
        .ListItems.Add , "FileUpload", "FileUpload", 5, 5
        .ListItems.Add , "Image", "Image", 6, 6
        .ListItems.Add , "ListBox", "ListBox", 10, 10
        .ListItems.Add , "Password", "Password", 7, 7
        .ListItems.Add , "Radio Button", "Radio Button", 8, 8
        .ListItems.Add , "Reset Button", "Reset Button", 2, 2
        .ListItems.Add , "Submit Button", "Submit Button", 2, 2
        .ListItems.Add , "Text Area", "Text Area", 11, 11
        .ListItems.Add , "TextBox", "TextBox", 12, 12
    End With

    frmMsg.PrintLog Time, "App.StatusText", "控件列表已初始化"
    
End Sub

Private Sub Form_Resize()
    lvwControls.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub lvwControls_ItemClick(ByVal Item As ComctlLib.ListItem)
    
    Dim strCtlName As String
    Dim blnRetVal As Boolean
    
    If Not apmMode = AppDesignMode Then Exit Sub
    If frmCode.brwDesign_Initialized = False Then Exit Sub
    
    Select Case Item.Text
        Case "Pointer"
            Exit Sub
            
        Case "Button", "CheckBox", "FileUpload", "Image", "Password"
            strCtlName = "Input" & Item.Text
            
        Case "Dropdown", "ListBox"
            strCtlName = "Select" & Item.Text
            
        Case "Text Area"
            strCtlName = "TextArea"
            
        Case "Reset Button"
            strCtlName = "InputReset"
            
        Case "Submit Button"
            strCtlName = "InputSubmit"
            
        Case "Radio Button"
            strCtlName = "InputRadio"
            
        Case "TextBox"
            strCtlName = "InputText"
            
        Case "FieldSet"
            strCtlName = "FieldSet"
    End Select

    blnRetVal = frmCode.hdocHTMLDesignDocument.execCommand("Insert" & strCtlName, False, 0)
    
    If blnRetVal = True Then
        frmMsg.PrintLog Time, "App.StatusText", "成功插入 """ & Item.Text & """ 控件"
    Else
        frmMsg.PrintLog Time, "App.StatusText", "插入 """ & Item.Text & """ 控件失败"
    End If
    
End Sub
