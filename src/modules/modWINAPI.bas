Attribute VB_Name = "modWINAPI"
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

Public Enum ICC_FLAGS
    ICC_ANIMATE_CLASS = &H80&
    ICC_BAR_CLASSES = &H4&
    ICC_COOL_CLASSES = &H400&
    ICC_DATE_CLASSES = &H100&
    ICC_HOTKEY_CLASS = &H40&
    ICC_INTERNET_CLASSES = &H800&
    ICC_LINK_CLASS = &H8000&
    ICC_LISTVIEW_CLASSES = &H1&
    ICC_NATIVEFNTCTL_CLASS = &H2000&
    ICC_PAGESCROLLER_CLASS = &H1000&
    ICC_PROGRESS_CLASS = &H20&
    ICC_STANDARD_CLASSES = &H4000&
    ICC_TAB_CLASSES = &H8&
    ICC_TREEVIEW_CLASSES = &H2&
    ICC_UPDOWN_CLASS = &H10&
    ICC_USEREX_CLASSES = &H200&
    ICC_WIN95_CLASSES = &HFF&
End Enum

Public Type tagINITCOMMONCONTROLSEX
    dwSize As Long
    dwICC As ICC_FLAGS
End Type

Public Declare Function InitCommonControlsEx Lib "comctl32" (ByRef picce As tagINITCOMMONCONTROLSEX) As Boolean

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

Public Declare Function ShellAboutW Lib "shell32.dll" (ByVal hwnd As Long, ByVal szApp As Long, ByVal szOtherStuff As Long, ByVal hIcon As Long) As Long
Public Declare Function HtmlHelpW Lib "hhctrl.ocx" (ByVal hwndCaller As Long, ByVal pszFile As Long, ByVal uCommand As Long, ByVal dwData As Long) As Long

Public Declare Sub ColorRGBToHLS Lib "shlwapi" (ByVal clrRGB As Long, ByRef pwHue As Integer, ByRef pwLuminance As Integer, ByRef pwSaturation As Integer)
Public Declare Function ColorHLSToRGB Lib "shlwapi" (ByVal wHue As Integer, ByVal wLuminance As Integer, ByVal wSaturation As Integer) As Long

Public Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal wType As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function GetWindowLongW Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLongW Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function SendMessageW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'    以下 API 的系统要求尚不明确，不同 MSDN 的说明不同
'    MSDN 2001
'        Requirements
'            Windows NT/2000: Requires Windows NT 3.51 or later
'            Windows 95/98/Me: Requires Windows 95 or later
'            Header: Declared in commctrl.h.
'            Import Library: comctl32.lib.
'
'    MSDN Online
'        Minimum supported client    Windows Vista [desktop apps only]
'        Minimum supported server    Windows Server 2003 [desktop apps only]
'        Target Platform             Windows
'        Header                      commctrl.h
'        Library                     comctl32.lib
'        dll                         comctl32.dll
Public Declare Function ImageList_ReplaceIcon Lib "comctl32" (ByVal himl As Long, ByVal i As Long, ByVal hIcon As Long) As Long
Public Declare Function ImageList_GetIcon Lib "comctl32" (ByVal himl As Long, ByVal i As Long, ByVal flags As Long) As Long

Public Function InitCommonControls()
    Dim iccex As tagINITCOMMONCONTROLSEX
    
    With iccex
        .dwSize = LenB(iccex)
        .dwICC = ICC_USEREX_CLASSES
    End With
    
    InitCommonControlsEx iccex
    
    iccex.dwICC = ICC_STANDARD_CLASSES
    
    InitCommonControlsEx iccex
    
    iccex.dwICC = ICC_COOL_CLASSES
    
    InitCommonControlsEx iccex
    
    iccex.dwICC = ICC_WIN95_CLASSES
    
    InitCommonControlsEx iccex
    
    iccex.dwICC = ICC_TREEVIEW_CLASSES
    
    InitCommonControlsEx iccex
End Function


Public Function CLRtoRGB(ByVal clrRGB As Long, ByRef R As Integer, ByRef G As Integer, ByRef B As Integer)
    R = clrRGB And &HFF&
    G = (clrRGB And &HFF00&) \ &H100&
    B = (clrRGB And &HFF0000) \ &H10000
End Function

Public Function CLRtoStr(ByVal RGB As Long)
    Dim i As Long
    Dim bytTemp(5) As Byte
    Dim lngTemp As Long
    
    lngTemp = RGB
    
    For i = 0 To 5
        bytTemp(i) = lngTemp Mod 16
        lngTemp = lngTemp \ 16
    Next i
    
    CLRtoStr = CStr(Hex(bytTemp(1))) & CStr(Hex(bytTemp(0))) & CStr(Hex(bytTemp(3))) & CStr(Hex(bytTemp(2))) & CStr(Hex(bytTemp(5))) & CStr(Hex(bytTemp(4)))
    
End Function

Public Function StrtoCLR(strColor As String) As Long
    If Len(strColor) <> 6 Then
        StrtoCLR = -1
        Exit Function
    End If
    
    StrtoCLR = RGB(CInt("&H" & Mid(strColor, 1, 2)), CInt("&H" & Mid(strColor, 3, 2)), CInt("&H" & Mid(strColor, 5, 2)))
    
End Function
