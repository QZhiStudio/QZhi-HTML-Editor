VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Begin VB.Form frmWidgets 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "组件"
   ClientHeight    =   9015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   ControlBox      =   0   'False
   Icon            =   "frmWidgets.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.TreeView treWidgets 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   14208
      _Version        =   327682
      Indentation     =   0
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ilsWidgets"
      Appearance      =   1
   End
   Begin ComctlLib.ImageList ilsWidgets 
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
         NumListImages   =   19
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWidgets.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWidgets.frx":035E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWidgets.frx":06B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWidgets.frx":0A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWidgets.frx":0D54
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWidgets.frx":10A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWidgets.frx":13F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWidgets.frx":174A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWidgets.frx":1A9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWidgets.frx":1DEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWidgets.frx":2140
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWidgets.frx":2492
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWidgets.frx":27E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWidgets.frx":2B36
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWidgets.frx":2E88
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWidgets.frx":31DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWidgets.frx":352C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWidgets.frx":387E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWidgets.frx":3BD0
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmWidgets"
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

Private Sub Form_Load()

    InitWidgetList
    frmMsg.PrintLog Time, "App.StatusText", "控件列表已初始化"
    Me.Refresh
    
End Sub

Public Function InitWidgetList()

    Dim nodX As Node
    
    treWidgets.Nodes.Clear

    With treWidgets
        
        Set nodX = .Nodes.Add(, , "Widgets", "组件", 1)
        nodX.ExpandedImage = 2
        
        Set nodX = .Nodes.Add("Widgets", tvwChild, "Regular", "常规", 1)
        nodX.ExpandedImage = 2
        
        .Nodes.Add "Regular", tvwChild, "Regular.Pointer", "Pointer", 3
        
        .Nodes.Add "Regular", tvwChild, "Button", "Button", 9
        .Nodes.Add "Regular", tvwChild, "HorizontalRule", "Horizontal Rule", 4
        .Nodes.Add "Regular", tvwChild, "IFrame", "IFrame", 5
        .Nodes.Add "Regular", tvwChild, "Image", "Image", 13
        .Nodes.Add "Regular", tvwChild, "OrderedList", "Ordered List", 6
        .Nodes.Add "Regular", tvwChild, "Paragraph", "Paragraph", 7
        .Nodes.Add "Regular", tvwChild, "UnorderedList", "Unordered List", 8
        
        Set nodX = .Nodes.Add("Widgets", tvwChild, "Controls", "表单组件", 1)
        nodX.ExpandedImage = 2
        
        .Nodes.Add "Controls", tvwChild, "Controls.Pointer", "Pointer", 3

        .Nodes.Add "Controls", tvwChild, "CheckBox", "CheckBox", 10
        .Nodes.Add "Controls", tvwChild, "Dropdown", "Dropdown", 16
        .Nodes.Add "Controls", tvwChild, "FieldSet", "FieldSet", 11
        .Nodes.Add "Controls", tvwChild, "FileUpload", "FileUpload", 12
        .Nodes.Add "Controls", tvwChild, "InputButton", "Input Button", 9
        .Nodes.Add "Controls", tvwChild, "InputImage", "Input Image", 13
        .Nodes.Add "Controls", tvwChild, "ListBox", "ListBox", 17
        .Nodes.Add "Controls", tvwChild, "Password", "Password", 14
        .Nodes.Add "Controls", tvwChild, "RadioButton", "Radio Button", 15
        .Nodes.Add "Controls", tvwChild, "ResetButton", "Reset Button", 9
        .Nodes.Add "Controls", tvwChild, "SubmitButton", "Submit Button", 9
        .Nodes.Add "Controls", tvwChild, "TextArea", "Text Area", 18
        .Nodes.Add "Controls", tvwChild, "TextBox", "TextBox", 19
        
        Set nodX = Nothing
    
        .Nodes("Widgets").Expanded = True
        .Nodes("Regular").Expanded = True
        .Nodes("Controls").Expanded = True
        
        .Refresh
        
    End With
    
End Function

Private Sub Form_Resize()
    treWidgets.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub treWidgets_NodeClick(ByVal Node As ComctlLib.Node)

    Dim strCmdID As String
    Dim blnRetVal As Boolean
    Dim vntID As Variant

    If gapmMode <> AppDesignMode Then Exit Sub
    If frmCode.brwDesign_Initialized = False Then Exit Sub
    
    Select Case Node.Key
        Case "Button", "HorizontalRule", "IFrame", "OrderedList", "UnorderedList", "InputButton", "Image", "InputImage", "FieldSet", "TextArea", "Paragraph"
            strCmdID = "Insert" & Node.Key
            
        Case "CheckBox", "FileUpload", "Password"
            strCmdID = "InsertInput" & Node.Key
            
        Case "Dropdown", "ListBox"
            strCmdID = "InsertSelect" & Node.Key
            
        Case "ResetButton"
            strCmdID = "InsertInputReset"
            
        Case "SubmitButton"
            strCmdID = "InsertInputSubmit"
            
        Case "TextBox"
            strCmdID = "InsertInputText"
            
        Case "RadioButton"
            strCmdID = "InsertInputRadio"
            
        Case Else
            Exit Sub
    End Select
    
    vntID = GenID(Node.Key)
    blnRetVal = frmCode.hdocHTMLDesignDocument.execCommand(strCmdID, True, vntID)

    If blnRetVal = True Then
        frmMsg.PrintLog Time, "App.StatusText", "成功插入 """ & Node.Text & """ 组件"
    Else
        frmMsg.PrintLog Time, "App.StatusText", "插入 """ & Node.Text & """ 组件失败"
    End If
    
End Sub

' 原理：
' 1. 循环，通过 ID 获取元素
' 2. 将元素的 tagName 赋值给 vntTemp
' 3. 如果没有元素，VB 会抛出错误，IsEmpty(vntTemp) = True，返回 ID
' 4. 跳转到步骤 1
Public Function GenID(ByVal strWidgetType As String) As String

    On Error Resume Next
    
    Dim i As Long
    Dim vntTemp As Variant
    
    vntTemp = 1
    
    While IsEmpty(vntTemp) = False
        i = i + 1
        vntTemp = Empty
        vntTemp = frmCode.hdocHTMLDesignDocument.getElementById(strWidgetType & CStr(i)).tagName
    Wend
    
    GenID = strWidgetType & CStr(i)
    
End Function
