Attribute VB_Name = "modAppCtrl"
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

Public Function SetAppMode(Mode As AppMode)
    gapmMode = Mode
    Select Case Mode
        Case AppDesignMode
            frmMain.Caption = App.ProductName & " - [设计]"
            frmMain.tlbFormat.Enabled = True
            frmMain.tlbBrowse.Enabled = False

            frmMain.excWidgets.Caption = "组件"
            
        Case AppEditMode
            frmMain.Caption = App.ProductName & " - [编辑]"
            frmMain.tlbFormat.Enabled = False
            frmMain.tlbBrowse.Enabled = False
            
            frmMain.excWidgets.Caption = "组件（不可用）"
            
        Case AppQuickViewMode
            frmMain.Caption = App.ProductName & " - [预览]"
            frmMain.tlbFormat.Enabled = False
            frmMain.tlbBrowse.Enabled = True
            
            'frmStructBrowser.InitDOMTree
            
            frmMain.excWidgets.Caption = "组件（不可用）"
            
            InitZoomCombo
    End Select
    
End Function

Public Function SaveCurrentFileAs() As Boolean
    ' 不用 execWB 因为无法获取文件名
    On Error GoTo userCanceled
    
    With frmMain.dlgMain
        .CancelError = True
        .flags = cdlOFNHideReadOnly + cdlOFNExplorer + cdlOFNNoDereferenceLinks + cdlOFNLongNames
        .Filter = "HTML 文档 (*.htm;*.html)|*.htm;*.html|聚合 HTML 文档 (*.mht;*.mhtml)|*.mht;*.mhtml|文本文件 (*.txt)|*.txt"
        .ShowSave
    End With
    
    gstrFileName = frmMain.dlgMain.FileName
    WriteToFile gstrFileName, GetCurrentDocHTML
    gstrDocHTML = GetCurrentDocHTML
    
    SaveCurrentFileAs = True
    
    frmCode.UpdateFormCaption gstrFileName
    
    Exit Function
    
userCanceled:
    SaveCurrentFileAs = False
    Exit Function
End Function

Public Function SaveCurrentFile() As Boolean
    If gstrFileName = "" Or Dir(gstrFileName) = "" Then
        SaveCurrentFile = SaveCurrentFileAs
    Else
        WriteToFile gstrFileName, GetCurrentDocHTML
        gstrDocHTML = GetCurrentDocHTML
        SaveCurrentFile = True
    End If
    frmCode.UpdateFormCaption gstrFileName
End Function

Public Function OpenHTMLDoc(ByVal strFileName As String) As String

    frmMain.brwWebControl_Initialized = False
    frmMain.brwWebControl.Navigate strFileName

    While frmMain.brwWebControl_Initialized = False
        DoEvents
    Wend

    If frmMain.brwWebControl_Initialized = True Then
        
        OpenHTMLDoc = frmMain.brwWebControl.Document.documentElement.outerHTML
        
    End If
    
End Function

' True 表示执行成功, False 表示取消操作
Public Function CloseCurrentFile() As Boolean
    
    Dim vmbrResult As VbMsgBoxResult    ' 接收 MsgBox 返回值
    Dim blnRet As Boolean   ' 接收 SaveCurrentFile 返回值
    
    Dim strFileName As String
    
    strFileName = "无标题"
    If gstrFileName <> "" Then strFileName = gstrFileName
    
    If IsDocChanged = True Then
        vmbrResult = MsgBox("您想将更改保存到 " & strFileName & " 吗？", vbQuestion + vbYesNoCancel, App.ProductName)
        
        
        Select Case vmbrResult
            Case vbYes
                SaveCurrentFile
                'blnRet = SaveCurrentFile
                
                If blnRet = False Then
                    CloseCurrentFile = False
                    Exit Function
                End If
            
            Case vbNo
                ' pass
                
            Case vbCancel
                CloseCurrentFile = False
                Exit Function
        End Select
    End If

    frmCode.tabMain.Tabs(2).Selected = True
    frmCode.tabMain_Click
    frmCode.eEditor.Value = ""
    gstrDocHTML = ""
    gstrFileName = ""
    frmCode.UpdateFormCaption ""
    
    CloseCurrentFile = True
    
End Function

Public Function GetCurrentDocHTML() As String
    Select Case gapmMode
        Case AppDesignMode
            GetCurrentDocHTML = frmCode.brwDesign.Document.documentElement.outerHTML
            
        Case AppEditMode, AppQuickViewMode
            GetCurrentDocHTML = frmCode.eEditor.Value
    End Select
End Function

Public Function CreateNewFile() As Boolean
    CreateNewFile = CloseCurrentFile
End Function

Public Function OpenNewFile() As Boolean
    OpenNewFile = CloseCurrentFile
    
    If OpenNewFile = False Then Exit Function
    
    On Error GoTo userCanceled
    
    With frmMain.dlgMain
        .CancelError = True
        .flags = cdlOFNHideReadOnly + cdlOFNExplorer + cdlOFNNoDereferenceLinks + cdlOFNLongNames + cdlOFNFileMustExist
        .Filter = "HTML 文档 (*.htm;*.html)|*.htm;*.html|聚合 HTML 文档 (*.mht;*.mhtml)|*.mht;*.mhtml|文本文件 (*.txt)|*.txt"
        .ShowOpen
    End With
    
    gstrFileName = frmMain.dlgMain.FileName
    
    gstrDocHTML = OpenHTMLDoc(gstrFileName)
    
    frmCode.tabMain.Tabs(2).Selected = True
    frmCode.eEditor.Value = gstrDocHTML
    frmCode.tabMain_Click
    frmCode.UpdateFormCaption gstrFileName
    
    OpenNewFile = True
    Exit Function
    
userCanceled:
    OpenNewFile = False
    Exit Function
End Function

Public Function IsDocChanged() As Boolean

    IsDocChanged = False

    If GetCurrentDocHTML <> gstrDocHTML Then IsDocChanged = True
End Function

Public Function QueryDocStat()

    If frmCode.brwDesign_Initialized = False Then Exit Function

    With gdsDocStat
        .vntBold = frmCode.hdocHTMLDesignDocument.queryCommandValue("Bold")
        .vntFontName = frmCode.hdocHTMLDesignDocument.queryCommandValue("FontName")
        .vntFontSize = frmCode.hdocHTMLDesignDocument.queryCommandValue("FontSize")
        .vntItalic = frmCode.hdocHTMLDesignDocument.queryCommandValue("Italic")
        .vntStrikeThrough = frmCode.hdocHTMLDesignDocument.queryCommandValue("StrikeThrough")
        .vntSubscript = frmCode.hdocHTMLDesignDocument.queryCommandValue("Subscript")
        .vntSuperscript = frmCode.hdocHTMLDesignDocument.queryCommandValue("Superscript")
        .vntUnderline = frmCode.hdocHTMLDesignDocument.queryCommandValue("Underline")
        
        .vntJustifyLeft = frmCode.hdocHTMLDesignDocument.queryCommandValue("JustifyLeft")
        .vntJustifyCenter = frmCode.hdocHTMLDesignDocument.queryCommandValue("JustifyCenter")
        .vntJustifyRight = frmCode.hdocHTMLDesignDocument.queryCommandValue("JustifyRight")
        .vntJustifyFull = frmCode.hdocHTMLDesignDocument.queryCommandValue("JustifyFull")
        .vntJustifyNone = frmCode.hdocHTMLDesignDocument.queryCommandValue("JustifyNone")
        
        .vntBackgroundColor = frmCode.hdocHTMLDesignDocument.queryCommandValue("BackColor")
        .vntForegroundColor = frmCode.hdocHTMLDesignDocument.queryCommandValue("ForeColor")
        
        ' Debug.Print "[" & Time & "]" & vbCrLf & _
            "Bold:            " & .vntBold & vbCrLf & _
            "FontName:        " & .vntFontName & vbCrLf & _
            "FontSize:        " & .vntFontSize & vbCrLf & _
            "Italic:          " & .vntItalic & vbCrLf & _
            "StrikeThrough:   " & .vntStrikeThrough & vbCrLf & _
            "Subscript:       " & .vntSubscript & vbCrLf & _
            "Superscript:     " & .vntSuperscript & vbCrLf & _
            "Underline:       " & .vntUnderline & vbCrLf & _
            vbCrLf & _
            "JustifyLeft:     " & .vntJustifyLeft & vbCrLf & _
            "JustifyCenter:   " & .vntJustifyCenter & vbCrLf & _
            "JustifyRight:    " & .vntJustifyRight & vbCrLf & _
            "JustifyFull:     " & .vntJustifyFull & vbCrLf & _
            "JustifyNone:     " & .vntJustifyNone & vbCrLf & _
            vbCrLf & _
            "BackgroundColor: " & .vntBackgroundColor & vbCrLf & _
            "ForegroundColor: " & .vntForegroundColor & vbCrLf
    End With
    
End Function

Public Function InitWebSafeColors()
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim lngTemp(5) As Long
    
    lngTemp(0) = &H0
    lngTemp(1) = &H33
    lngTemp(2) = &H66
    lngTemp(3) = &H99
    lngTemp(4) = &HCC
    lngTemp(5) = &HFF
    
    For i = 0 To 5
        For j = 0 To 5
            For k = 0 To 5
                lngWebSafeColor(i * 36 + j * 6 + k) = RGB(lngTemp(i), lngTemp(j), lngTemp(k))
            Next k
        Next j
    Next i
End Function

Public Function GetIEVersion() As Long
    On Error GoTo OldVer

    Dim vntTemp As Variant
    Dim wshShell As Object
    
    Dim i As Long
    
    Set wshShell = CreateObject("WScript.Shell")
    
    vntTemp = wshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\svcVersion")
    
    GoTo GetVer
    
OldVer:
    i = i + 1
    If i = 2 Then GoTo FuncError    ' 第二次跳转到此标签，必然是没有合适的 IE
    vntTemp = wshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Version")
    
    GoTo GetVer
    
GetVer:
    vntTemp = Split(vntTemp, ".")
    GetIEVersion = CLng(vntTemp(0))
    
    Set wshShell = Nothing
    
    Exit Function
    
FuncError:
    MsgBox "未能检测到您安装有 Microsoft Internet Explorer", vbCritical, App.ProductName
    End
End Function

Private Function InitZoomCombo()

    Dim vntRange As Variant
    Dim lngRange As Long
    
    Dim strText As String
    
    Dim lngStep As Long
    Dim i As Long
    
    Dim vntZoom As Variant
    Dim lngZoom As Variant

    If App.LogMode = 0 Or glngIEVersion < 7 Then    ' 在 IDE 中或 IE 版本低于 7
        frmCode.brwQuickView.ExecWB OLECMDID_GETZOOMRANGE, OLECMDEXECOPT_DODEFAULT, 0, vntRange
        lngRange = CLng(vntRange)
        gintMinZoom = lngRange And &HFFFF&
        gintMaxZoom = (lngRange And &HFFFF0000) \ &H10000
        
        frmMain.cboZoom.Clear
        
        For i = gintMinZoom To gintMaxZoom
            Select Case i
                Case 0
                    strText = "最小"
                Case 1
                    strText = "较小"
                Case 2
                    strText = "中"
                Case 3
                    strText = "较大"
                Case 4
                    strText = "最大"
            End Select
            
            frmMain.cboZoom.AddItem strText
            
        Next i
        
        frmCode.brwQuickView.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DODEFAULT, 0, vntZoom ' 取得当前缩放比
        frmMain.cboZoom.ListIndex = CLng(vntZoom) - gintMinZoom
        glngDefaultZoom = CLng(vntZoom)
        
    ElseIf glngIEVersion >= 7 Then
        
        frmCode.brwQuickView.ExecWB OLECMDID_OPTICAL_GETZOOMRANGE, OLECMDEXECOPT_DODEFAULT, 0, vntRange
        lngRange = CLng(vntRange)
        gintMinZoom = lngRange And &HFFFF&
        gintMaxZoom = (lngRange And &HFFFF0000) \ &H10000
        
        frmMain.cboZoom.Clear
        
        lngStep = 15
        
        i = gintMinZoom
        
        While i <= gintMaxZoom
            frmMain.cboZoom.AddItem i & "%"
            
            Select Case i
                Case 25
                    lngStep = 25

                Case 200
                    lngStep = 50

                Case 400
                    lngStep = 100
            End Select
            
            i = i + lngStep
        Wend

        ' 本来是打算 Get 的，参考了一些资料，这样写：
        ' frmCode.brwQuickView.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DODEFAULT, 0, vntZoom
        ' 但是怎么改都会报错（除非 pvaIn 有实际意义的参数，但这样这个 Get 的过程没有任何意义），就作罢了
        frmCode.brwQuickView.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DODEFAULT, CLng(100), 0    ' 强制设置缩放比为 100%
        glngDefaultZoom = 100
        strText = "100%"

        For i = 0 To frmMain.cboZoom.ListCount - 1
            If frmMain.cboZoom.List(i) = strText Then
                frmMain.cboZoom.ListIndex = i
            End If
        Next i
        
    End If
End Function

