Attribute VB_Name = "modDef"
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

Public Type AppMsg
    strType As String
    strTime As String
    strText As String
End Type

Public Enum AppMode
    AppDesignMode = 1
    AppEditMode = 2
    AppQuickViewMode = 3
End Enum

Public Type RGBColor
    R As Integer
    G As Integer
    B As Integer
End Type

Public Type HSLColor
    H As Integer
    S As Integer
    L As Integer
End Type

Public Type DocStat
    vntBold As Variant
    vntItalic As Variant
    vntUnderline As Variant
    vntStrikeThrough As Variant
    vntSubscript As Variant
    vntSuperscript As Variant
    vntFontName As Variant
    vntFontSize As Variant
    
    vntJustifyCenter As Variant
    vntJustifyFull As Variant
    vntJustifyLeft As Variant
    vntJustifyRight As Variant
    vntJustifyNone As Variant
    
    vntBackgroundColor As Variant
    vntForegroundColor As Variant
    
    blnNavigateBack As Boolean
    blnNavigateForward As Boolean
End Type

Public Type List
    idxIndex As Long
    strKey As String
    strText As String
End Type

Public Type DOMNode
    Node As IHTMLElement
    Parent As IHTMLElement
    tagName As String
    parentTag As String
    id As String
    className As String
    idx As Long
    parentIdx As Long
End Type

Public Type DOMTree
    strHash As String
    nDOMNode() As DOMNode
End Type

Public gapmMode As AppMode
Public glngIEVersion As Long

Public gintMinZoom As Integer
Public gintMaxZoom As Integer

Public lngWebSafeColor(215) As Long

Public gdsDocStat As DocStat

Public glngDefaultZoom As Long

Public gstrDocHTML As String
Public gstrFileName As String

Public gDOMTree() As DOMTree

'Public gTick As Long
