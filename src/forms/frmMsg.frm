VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMsg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "消息"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   ControlBox      =   0   'False
   Icon            =   "frmMsg.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar tlbMsg 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ilsMsg"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   2
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ClearMsgList"
            Object.ToolTipText     =   "清空消息列表"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "SaveMsgList"
            Object.ToolTipText     =   "保存消息列表"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   3360
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser brwMsgList 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3135
      ExtentX         =   5530
      ExtentY         =   3836
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
   Begin ComctlLib.ImageList ilsMsg 
      Left            =   3840
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMsg.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMsg.frx":035E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMsg"
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

Private blnIsInitialized As Boolean
Private blnBufferUsed As Boolean
Private msgMsgBuffer() As AppMsg
Private blnLastItemIsOK As Boolean

Private WithEvents hdocHtmlDocument As HTMLDocument
Attribute hdocHtmlDocument.VB_VarHelpID = -1

Private Sub brwMsgList_DocumentComplete(ByVal pDisp As Object, URL As Variant)

    Set hdocHtmlDocument = brwMsgList.Document

    With brwMsgList.Document
        .open
    End With
    
    blnIsInitialized = True

    If blnBufferUsed = True Then
        Dim i As Long
        
        For i = 0 To UBound(msgMsgBuffer)
            With msgMsgBuffer(i)
                PrintLog .strTime, .strType, .strText
            End With
        Next i
        
        blnBufferUsed = False
    End If
    
End Sub

Private Sub Form_Load()
    blnIsInitialized = False
    brwMsgList.Navigate "about:blank"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.ScaleHeight - tlbMsg.Height Then brwMsgList.Move 0, tlbMsg.Height, Me.ScaleWidth, Me.ScaleHeight - tlbMsg.Height
End Sub

Public Function PrintLog(ByVal strTime As String, ByVal strMsgType As String, ByVal strText As String) As Long

    Dim strColor As String

    If strText = "" Then Exit Function
    
    If strMsgType = "WebBrowser.StatusText" Then
        If (strText = "完成") Or (strText = "完毕") Then
            If blnLastItemIsOK = True Then Exit Function
            blnLastItemIsOK = True
        Else
            blnLastItemIsOK = False
        End If
    Else
        blnLastItemIsOK = False
    End If
    
    strColor = "#00539C"
    
    If InStr(1, strMsgType, "Error") Then
        strColor = "#CC0000"
    End If
    If InStr(1, strMsgType, "Warning") Then
        strColor = "#FFCC00"
    End If

    If blnIsInitialized = True Then
        
        With brwMsgList.Document
            .write "<font face=""Courier New""><p style=""font-size: 16px; white-space: nowrap; margin:4px 0;""><b style=""color: " & strColor & ";"">[" & StringtoEntity(strTime) & "] [" & StringtoEntity(strMsgType) & "] </b>" & StringtoEntity(strText) & "</p></font>"
        End With
        
        hdocHtmlDocument.parentWindow.scrollBy 0, hdocHtmlDocument.body.scrollHeight
        
        PrintLog = 0
    Else
        PushBuffer Time, strMsgType, strText
        PrintLog = -1
    End If
    
End Function

Private Function PushBuffer(ByVal strTime As String, ByVal strMsgType As String, ByVal strText As String)
    If blnBufferUsed = False Then
        ReDim Preserve msgMsgBuffer(0)
    Else
        ReDim Preserve msgMsgBuffer(UBound(msgMsgBuffer) + 1)
    End If
    
    With msgMsgBuffer(UBound(msgMsgBuffer))
            .strTime = strTime
            .strText = strText
            .strType = strMsgType
    End With
    
    blnBufferUsed = True
End Function

Private Sub tlbMsg_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "ClearMsgList"
            blnIsInitialized = False
            brwMsgList.Navigate "about:blank"
            
        Case "SaveMsgList"
            SaveMsgList
    End Select
End Sub

Private Function SaveMsgList()
    On Error GoTo userCanceled
    
    With dlgMain
        .CancelError = True
        .flags = cdlOFNHideReadOnly + cdlOFNExplorer + cdlOFNNoDereferenceLinks + cdlOFNLongNames + cdlOFNOverwritePrompt
        .Filter = "HTML 文档 (*.htm;*.html)|*.htm;*.html|聚合 HTML 文档 (*.mht;*.mhtml)|*.mht;*.mhtml|文本文件 (*.txt)|*.txt"
        .ShowSave
    End With

    WriteToFile dlgMain.FileName, brwMsgList.Document.documentElement.outerHTML
    
    Exit Function
    
userCanceled:
    Exit Function
End Function
