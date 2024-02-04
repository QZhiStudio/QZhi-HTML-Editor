VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmMsg 
   Caption         =   "消息"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   ControlBox      =   0   'False
   Icon            =   "frmMsg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin SHDocVwCtl.WebBrowser brwMsgList 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      ExtentX         =   7646
      ExtentY         =   4895
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
    brwMsgList.Navigate "about:blank"
End Sub

Private Sub Form_Resize()
    brwMsgList.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Public Function PrintLog(ByVal strTime As String, ByVal strMsgType As String, ByVal strText As String) As Long

    If strText = "" Then Exit Function

    If blnIsInitialized = True Then
        
        With brwMsgList.Document
            .write "<p style=""font-size: 16px;""><b style=""color: #00539C;"">[" & strTime & "] [" & strMsgType & "] </b>" & strText & "</p>"
        End With
        
        hdocHtmlDocument.parentWindow.scrollBy 0, hdocHtmlDocument.documentElement.clientHeight
        
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

