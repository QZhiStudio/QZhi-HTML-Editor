Attribute VB_Name = "modAPI"
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

Public Function WriteHTML(brwWeb As WebBrowser, strHTML As String)
    On Error GoTo FuncError

    With brwWeb.Document
        .open
        .Clear
        .write strHTML
        .Close
    End With
    
    Exit Function
    
FuncError:
    frmMsg.PrintLog Time, "App.Error", "无法写入 HTML 到指定控件，请立刻保存文档，然后重启程序！"
    Exit Function
    
End Function

Public Function StringtoEntity(ByVal strString As String) As String
    Dim lngStrLen As Long
    Dim arrintString() As Integer
    Dim arrintBuffer() As Integer
    
    Dim arrintStack(4) As Integer   ' 用于存储单字符数据的“栈”
    Dim lngSP As Long
    
    Dim lpBuffer As Long    ' 数组“指针”
    Dim lngTemp As Long
    
    Dim i As Long
    Dim j As Long
    
    lngStrLen = LenB(strString)
    
    If lngStrLen = 0 Then Exit Function   ' 空字符串
    If lngStrLen Mod 2 = 1 Then Exit Function ' 一定不是 Unicode

    ReDim arrintString((lngStrLen / 2) - 1)
    ReDim arrintBuffer(((lngStrLen / 2) * 7) - 1)
    
    CopyMemory VarPtr(arrintString(0)), StrPtr(strString), lngStrLen
    
    For i = 0 To (lngStrLen / 2) - 1    ' 逐字符遍历
        If arrintString(i) <> 0 Then
            lngTemp = arrintString(i) And &HFFFF&
            
            Select Case lngTemp
                Case 13 ' CR
                    ' pass
                    
                Case 10 ' LF
                    arrintBuffer(lpBuffer) = 60 ' "<"
                    arrintBuffer(lpBuffer + 1) = 98 ' "b"
                    arrintBuffer(lpBuffer + 2) = 114 ' "r"
                    arrintBuffer(lpBuffer + 3) = 32 ' <Space>
                    arrintBuffer(lpBuffer + 4) = 47 ' "/"
                    arrintBuffer(lpBuffer + 5) = 62 ' ">"
                    lpBuffer = lpBuffer + 6
                    
                Case 9 ' Tab
                    arrintBuffer(lpBuffer) = 38 ' "&"
                    arrintBuffer(lpBuffer + 1) = 101 ' "e"
                    arrintBuffer(lpBuffer + 2) = 109 ' "m"
                    arrintBuffer(lpBuffer + 3) = 115 ' "s"
                    arrintBuffer(lpBuffer + 4) = 112 ' "p"
                    arrintBuffer(lpBuffer + 5) = 59 ' ";"
                    lpBuffer = lpBuffer + 6
            
                Case Else
                    arrintBuffer(lpBuffer) = 38 ' "&"
                    arrintBuffer(lpBuffer + 1) = 35 ' "#"
                    
                    lpBuffer = lpBuffer + 2
                    
                    lngSP = 0
                    
                    While lngTemp <> 0
                        arrintStack(lngSP) = 48 + (lngTemp Mod 10&)
                        lngTemp = lngTemp \ 10&
                        lngSP = lngSP + 1
                    Wend
                    
                    While lngSP > 0
                        lngSP = lngSP - 1
                        arrintBuffer(lpBuffer) = arrintStack(lngSP)
                        lpBuffer = lpBuffer + 1
                    Wend
            End Select
        End If
    Next i
    
    StringtoEntity = String(lpBuffer, 32)
    CopyMemory StrPtr(StringtoEntity), VarPtr(arrintBuffer(0)), lpBuffer * 2
    
FuncExit:
    Erase arrintString ' 释放内存
    Erase arrintBuffer
End Function

'Public Function StringtoEntity(ByVal strString As String) As String
'    Dim i As Long
'
'    If Len(strString) = 0 Then Exit Function
'
'    For i = 1 To Len(strString)
'        DoEvents
'        Select Case Mid(strString, i, 1)
'            Case vbCr
'                'pass
'
'            Case vbLf
'                StringtoEntity = StringtoEntity & "<br />"
'
'            Case " "
'                StringtoEntity = StringtoEntity & "&nbsp;"
'
'            Case vbTab
'                StringtoEntity = StringtoEntity & "&nbsp;&nbsp;&nbsp;&nbsp;"
'
'            Case Else
'                StringtoEntity = StringtoEntity & "&#" & CLng(&HFFFF& And CInt(AscW(Mid(strString, i, 1))))
'        End Select
'    Next i
'
'End Function

Public Function WriteToFile(ByVal strFileName As String, ByVal strData As String)
    Dim intFileNum As Integer
    
    intFileNum = FreeFile
    
    Open strFileName For Output As #intFileNum
    
        Print #intFileNum, strData
    
    Close #intFileNum
    
End Function
