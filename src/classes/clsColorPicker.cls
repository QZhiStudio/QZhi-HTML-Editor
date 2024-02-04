VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

Private Sub Class_Initialize()
    '
End Sub

Public Function GetColor(ByVal lngCurrentColor As Long, ByVal lngNewColor As Long) As Long
    Dim frmCP As New frmColorPicker
    Dim i As Long
    Dim blnIsNewColorWebSafe As Boolean
    
    frmCP.picColors.Line (0, 37)-(73, 73), lngCurrentColor, BF
    frmCP.picColors.Refresh
    
    For i = 0 To 215
        If lngNewColor = lngWebSafeColor(i) Then
            frmCP.imgWebSafeColor_Click CInt(i)
            blnIsNewColorWebSafe = True
        End If
    Next i
    
    If blnIsNewColorWebSafe = False Then
        Dim H_ As Integer
        Dim S_ As Integer
        Dim L_ As Integer
        
        ColorRGBToHLS lngNewColor, H_, L_, S_
        
        frmCP.picColorsRectangle_MouseDown 1, 0, CSng(H_), CSng(CSng(240) - S_)
        frmCP.picLuminance_MouseDown 1, 0, 0, CSng(CSng(240) - L_)
        
    End If
    
    frmCP.blnIsInitialized = True
    frmCP.Show vbModal
    
    If frmCP.blnIsSelected = True Then
        GetColor = frmCP.lngColorRef
    Else
        GetColor = -1
    End If
    
    Set frmCP = Nothing
End Function
