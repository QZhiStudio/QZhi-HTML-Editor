Attribute VB_Name = "modMain"
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

Sub Main()

    If App.LogMode = 1 Then
        If Dir("QZHE.chm", vbNormal) <> "" Then
            App.HelpFile = "QZHE.chm"
        Else
            MsgBox "�����ļ�ȱʧ�������޷�����", vbCritical, App.ProductName
            End
        End If
    End If
    
    glngIEVersion = GetIEVersion()
    
    If glngIEVersion < 5 Then
        MsgBox "���� Microsoft Internet Explorer �汾���ͣ��޷����б����", vbCritical, App.ProductName
        End
    End If

    CreateObject("WScript.Shell").RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION\" & App.EXEName + ".exe", CStr(glngIEVersion * 1000), "REG_DWORD"
    
    InitWebSafeColors
    
'    ����
'    Dim a As New clsColorPicker
'    MsgBox CLRtoStr(a.GetColor(RGB(40, 123, 215)))
'    Exit Sub
'    ����

    
    
    frmMsg.PrintLog Time, "App.StatusText", "��ɫ�������ʼ���ɹ�"
    
    frmMenu.Move 0, 0, Screen.Width, frmMenu.Height - frmMenu.ScaleHeight + frmMenu.tlbMain.Height * 3
    frmMenu.Show
    frmWidgets.Move 0, frmMenu.Height + 120, 3600, Screen.Height - frmMenu.Height - 960
    frmWidgets.Show
    frmCode.Move frmWidgets.Width + 120, frmMenu.Height + 120, Screen.Width - frmWidgets.Width - 120, (Screen.Height - frmMenu.Height - 960 - 120) * 0.7
    frmCode.Show
    frmMsg.Move frmWidgets.Width + 120, frmCode.Top + frmCode.Height + 120, Screen.Width - frmWidgets.Width - 120, (Screen.Height - frmMenu.Height - 960 - 120) * 0.3
    frmMsg.Show
    
    ' frmColorPicker.Show vbModal
    
End Sub

Public Function WriteToFile(ByVal strFileName As String, ByVal strData As String)
    Dim intFileNum As Integer
    
    intFileNum = FreeFile
    
    Open strFileName For Output As #intFileNum
    
        Print #intFileNum, strData
    
    Close #intFileNum
    
End Function

Public Function SaveCurrentFileAs() As Boolean
    ' ���� execWB ��Ϊ�޷���ȡ�ļ���
    On Error GoTo userCanceled
    
    With frmMenu.dlgMain
        .CancelError = True
        .flags = cdlOFNHideReadOnly + cdlOFNExplorer + cdlOFNNoDereferenceLinks + cdlOFNLongNames
        .Filter = "HTML �ĵ� (*.htm;*.html)|*.htm;*.html|�ۺ� HTML �ĵ� (*.mht;*.mhtml)|*.mht;*.mhtml|�ı��ļ� (*.txt)|*.txt"
        .ShowSave
    End With
    
    gstrFileName = frmMenu.dlgMain.FileName
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

    frmMenu.brwWebControl_Initialized = False
    frmMenu.brwWebControl.Navigate strFileName

    While frmMenu.brwWebControl_Initialized = False
        DoEvents
    Wend

    If frmMenu.brwWebControl_Initialized = True Then
        
        OpenHTMLDoc = frmMenu.brwWebControl.Document.documentElement.outerHTML
        
    End If
    
End Function

' True ��ʾִ�гɹ�, False ��ʾȡ������
Public Function CloseCurrentFile() As Boolean
    
    Dim vmbrResult As VbMsgBoxResult    ' ���� MsgBox ����ֵ
    Dim blnRet As Boolean   ' ���� SaveCurrentFile ����ֵ
    
    Dim strFileName As String
    
    strFileName = "�ޱ���"
    If gstrFileName <> "" Then strFileName = gstrFileName
    
    If IsDocChanged = True Then
        vmbrResult = MsgBox("���뽫���ı��浽 " & strFileName & " ��", vbQuestion + vbYesNoCancel, App.ProductName)
        
        
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
    
    With frmMenu.dlgMain
        .CancelError = True
        .flags = cdlOFNHideReadOnly + cdlOFNExplorer + cdlOFNNoDereferenceLinks + cdlOFNLongNames + cdlOFNFileMustExist
        .Filter = "HTML �ĵ� (*.htm;*.html)|*.htm;*.html|�ۺ� HTML �ĵ� (*.mht;*.mhtml)|*.mht;*.mhtml|�ı��ļ� (*.txt)|*.txt"
        .ShowOpen
    End With
    
    gstrFileName = frmMenu.dlgMain.FileName
    
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

Public Sub AtExit()
    On Error Resume Next
    CreateObject("WScript.Shell").RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION\" & App.EXEName + ".exe"
    End
End Sub

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
    frmMsg.PrintLog Time, "App.Error", "�޷�д�� HTML ��ָ���ؼ��������̱����ĵ���Ȼ����������"
    Exit Function
    
End Function

Public Function SetAppMode(Mode As AppMode)
    gapmMode = Mode
    Select Case Mode
        Case AppDesignMode
            frmMenu.Caption = App.ProductName & " - [���]"
            frmMenu.tlbFormat.Enabled = True
            frmMenu.tlbBrowse.Enabled = False
            ' frmWidgets.treWidgets.Enabled = True
            
        Case AppEditMode
            frmMenu.Caption = App.ProductName & " - [�༭]"
            frmMenu.tlbFormat.Enabled = False
            frmMenu.tlbBrowse.Enabled = False
            ' frmWidgets.treWidgets.Enabled = False ' ����ȥ̫������
            
        Case AppQuickViewMode
            frmMenu.Caption = App.ProductName & " - [Ԥ��]"
            frmMenu.tlbFormat.Enabled = False
            frmMenu.tlbBrowse.Enabled = True
            ' frmWidgets.treWidgets.Enabled = False
            
            InitZoomCombo
    End Select
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
    If i = 2 Then GoTo FuncError    ' �ڶ�����ת���˱�ǩ����Ȼ��û�к��ʵ� IE
    vntTemp = wshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Version")
    
    GoTo GetVer
    
GetVer:
    vntTemp = Split(vntTemp, ".")
    GetIEVersion = CLng(vntTemp(0))
    
    Set wshShell = Nothing
    
    Exit Function
    
FuncError:
    MsgBox "δ�ܼ�⵽����װ�� Microsoft Internet Explorer", vbCritical, App.ProductName
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

    If App.LogMode = 0 Or glngIEVersion < 7 Then    ' �� IDE �л� IE �汾���� 7
        frmCode.brwQuickView.ExecWB OLECMDID_GETZOOMRANGE, OLECMDEXECOPT_DODEFAULT, 0, vntRange
        lngRange = CLng(vntRange)
        gintMinZoom = lngRange And &HFFFF&
        gintMaxZoom = (lngRange And &HFFFF0000) \ &H10000
        
        frmMenu.cboZoom.Clear
        
        For i = gintMinZoom To gintMaxZoom
            Select Case i
                Case 0
                    strText = "��С"
                Case 1
                    strText = "��С"
                Case 2
                    strText = "��"
                Case 3
                    strText = "�ϴ�"
                Case 4
                    strText = "���"
            End Select
            
            frmMenu.cboZoom.AddItem strText
            
        Next i
        
        frmCode.brwQuickView.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DODEFAULT, 0, vntZoom ' ȡ�õ�ǰ���ű�
        frmMenu.cboZoom.ListIndex = CLng(vntZoom) - gintMinZoom
        glngDefaultZoom = CLng(vntZoom)
        
    ElseIf glngIEVersion >= 7 Then
        
        frmCode.brwQuickView.ExecWB OLECMDID_OPTICAL_GETZOOMRANGE, OLECMDEXECOPT_DODEFAULT, 0, vntRange
        lngRange = CLng(vntRange)
        gintMinZoom = lngRange And &HFFFF&
        gintMaxZoom = (lngRange And &HFFFF0000) \ &H10000
        
        frmMenu.cboZoom.Clear
        
        lngStep = 15
        
        i = gintMinZoom
        
        While i <= gintMaxZoom
            frmMenu.cboZoom.AddItem i & "%"
            
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

        ' �����Ǵ��� Get �ģ��ο���һЩ���ϣ�����д��
        ' frmCode.brwQuickView.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DODEFAULT, 0, vntZoom
        ' ������ô�Ķ��ᱨ������ pvaIn ��ʵ������Ĳ�������������� Get �Ĺ���û���κ����壩����������
        frmCode.brwQuickView.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DODEFAULT, CLng(100), 0    ' ǿ���������ű�Ϊ 100%
        glngDefaultZoom = 100
        strText = "100%"

        For i = 0 To frmMenu.cboZoom.ListCount - 1
            If frmMenu.cboZoom.List(i) = strText Then
                frmMenu.cboZoom.ListIndex = i
            End If
        Next i
        
    End If
End Function
