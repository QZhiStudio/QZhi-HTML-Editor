VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Begin VB.Form frmCode 
   Caption         =   "无标题 - QZhi HTML Editor"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360
   ControlBox      =   0   'False
   Icon            =   "frmCode.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin SHDocVwCtl.WebBrowser brwQuickView 
      Height          =   4335
      Left            =   6000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   3015
      ExtentX         =   5318
      ExtentY         =   7646
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
   Begin SHDocVwCtl.WebBrowser brwSource 
      Height          =   4335
      Left            =   3000
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   2895
      ExtentX         =   5106
      ExtentY         =   7646
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
   Begin SHDocVwCtl.WebBrowser brwDesign 
      Height          =   4335
      Left            =   240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   2655
      ExtentX         =   4683
      ExtentY         =   7646
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
   Begin ComctlLib.TabStrip tabMain 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8705
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "设计(&D)"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "设计"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "源码(&S)"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "源码"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "预览(&Q)"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "预览"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCode"
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

Public brwDesign_Initialized As Boolean
Public brwSource_Initialized As Boolean
Public brwQuickView_Initialized As Boolean

Dim idxSelectedTabItem As Long

Public WithEvents eEditor As HTMLTextAreaElement
Attribute eEditor.VB_VarHelpID = -1

Public WithEvents hdocHTMLDesignDocument As HTMLDocument
Attribute hdocHTMLDesignDocument.VB_VarHelpID = -1

Private Sub brwDesign_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    brwDesign_Initialized = True
    brwDesign.Document.execCommand "EditMode", False, vbNullString
    brwDesign.Document.execCommand "LiveResize", False, vbNullString
    Set hdocHTMLDesignDocument = brwDesign.Document
End Sub

Private Sub brwDesign_DownloadBegin()
    brwDesign.Silent = True
End Sub

Private Sub brwDesign_DownloadComplete()
    brwDesign.Silent = True
End Sub

Private Sub brwQuickView_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
     If Command = CSC_NAVIGATEBACK Then
         gdsDocStat.blnNavigateBack = Enable
     End If
     If Command = CSC_NAVIGATEFORWARD Then
         gdsDocStat.blnNavigateForward = Enable
     End If
     frmMenu.SetToolbarStat
End Sub

Private Sub brwQuickView_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    brwQuickView_Initialized = True
End Sub

Private Sub brwQuickView_DownloadBegin()
    brwQuickView.Silent = True
End Sub

Private Sub brwQuickView_DownloadComplete()
    brwQuickView.Silent = True
End Sub

Private Sub brwQuickView_NewWindow2(ppDisp As Object, Cancel As Boolean)
    brwQuickView.Navigate2 brwQuickView.Document.activeElement.href
    Cancel = True
End Sub

Private Sub brwQuickView_StatusTextChange(ByVal Text As String)
    frmMsg.PrintLog Time, "WebBrowser.StatusText", Text
End Sub

Private Sub brwSource_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    brwSource_Initialized = True
    WriteHTML brwSource, StrConv(LoadResData("EDITOR", 23), vbUnicode)
    Set eEditor = brwSource.Document.getElementById("editor")
End Sub

' onchange 不起效啊！！！
Private Sub eEditor_onkeydown()
    UpdateFormCaption gstrFileName
End Sub

Private Sub eEditor_onkeyup()
    UpdateFormCaption gstrFileName
End Sub

Private Sub eEditor_onmousedown()
    UpdateFormCaption gstrFileName
End Sub

Private Sub eEditor_onmouseup()
    UpdateFormCaption gstrFileName
End Sub

Private Sub Form_Load()
    brwSource.Navigate "about:blank"
    brwDesign.Navigate "about:blank"
    brwQuickView.Navigate "about:blank"
    tabMain.Tabs(2).Selected = True
    idxSelectedTabItem = 2
    tabMain_Click
    SetAppMode AppEditMode
    frmMsg.PrintLog Time, "App.StatusText", "HTML 编辑器已初始化"
    
    
End Sub

Private Sub Form_Resize()
    With tabMain
        .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        brwDesign.Move .clientLeft, .clientTop, .clientWidth, .clientHeight
        brwSource.Move .clientLeft, .clientTop, .clientWidth, .clientHeight
        brwQuickView.Move .clientLeft, .clientTop, .clientWidth, .clientHeight
    End With
End Sub

Private Function hdocHTMLDesignDocument_oncontextmenu() As Boolean
    hdocHTMLDesignDocument_oncontextmenu = False
End Function

Private Sub hdocHTMLDesignDocument_onkeydown()
    UpdateFormCaption gstrFileName
End Sub

Private Sub hdocHTMLDesignDocument_onkeyup()
    UpdateFormCaption gstrFileName
End Sub

Private Sub hdocHTMLDesignDocument_onmousedown()
    UpdateFormCaption gstrFileName
End Sub

Private Sub hdocHTMLDesignDocument_onmouseup()
    UpdateFormCaption gstrFileName
End Sub

Public Sub tabMain_Click()
    Select Case tabMain.SelectedItem.Index
        Case 1
            SetAppMode AppDesignMode
            
            brwDesign.ZOrder 0
            
            If brwDesign_Initialized = True Then WriteHTML brwDesign, eEditor.Value
        
        Case 2
            SetAppMode AppEditMode
            
            brwSource.ZOrder 0
            
            If idxSelectedTabItem = 1 Then
                eEditor.Value = brwDesign.Document.documentElement.outerHTML
            End If
        
        Case 3
            SetAppMode AppQuickViewMode
            
            brwQuickView.ZOrder 0
            
            If idxSelectedTabItem = 1 Then
                eEditor.Value = brwDesign.Document.documentElement.outerHTML
            End If
            
            brwQuickView_Initialized = False
            brwQuickView.Navigate "about:blank"
            While brwQuickView_Initialized = False
                DoEvents
            Wend
            WriteHTML brwQuickView, eEditor.Value
        
    End Select
    
    idxSelectedTabItem = tabMain.SelectedItem.Index
    
End Sub

Public Function UpdateFormCaption(strFileName As String)
    Dim strCaption As String
    
    If strFileName = "" Then
        strCaption = "无标题 - " & App.ProductName
    Else
        strCaption = strFileName & " - " & App.ProductName
    End If
    
    If IsDocChanged = True Then
        strCaption = "*" & strCaption
    End If
    
    Me.Caption = strCaption
End Function
