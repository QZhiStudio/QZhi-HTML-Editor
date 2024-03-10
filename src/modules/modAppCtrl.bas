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
            frmMain.Caption = App.ProductName & " - [���]"
            frmMain.tlbFormat.Enabled = True
            frmMain.tlbBrowse.Enabled = False

            frmMain.excWidgets.Caption = "���"
            
        Case AppEditMode
            frmMain.Caption = App.ProductName & " - [�༭]"
            frmMain.tlbFormat.Enabled = False
            frmMain.tlbBrowse.Enabled = False
            
            frmMain.excWidgets.Caption = "����������ã�"
            
        Case AppQuickViewMode
            frmMain.Caption = App.ProductName & " - [Ԥ��]"
            frmMain.tlbFormat.Enabled = False
            frmMain.tlbBrowse.Enabled = True
            
            'frmStructBrowser.InitDOMTree
            
            frmMain.excWidgets.Caption = "����������ã�"
            
            InitZoomCombo
    End Select
    
End Function

Public Function SaveCurrentFileAs() As Boolean
    ' ���� execWB ��Ϊ�޷���ȡ�ļ���
    On Error GoTo userCanceled
    
    With frmMain.dlgMain
        .CancelError = True
        .flags = cdlOFNHideReadOnly + cdlOFNExplorer + cdlOFNNoDereferenceLinks + cdlOFNLongNames
        .Filter = "HTML �ĵ� (*.htm;*.html)|*.htm;*.html|�ۺ� HTML �ĵ� (*.mht;*.mhtml)|*.mht;*.mhtml|�ı��ļ� (*.txt)|*.txt"
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
    
    With frmMain.dlgMain
        .CancelError = True
        .flags = cdlOFNHideReadOnly + cdlOFNExplorer + cdlOFNNoDereferenceLinks + cdlOFNLongNames + cdlOFNFileMustExist
        .Filter = "HTML �ĵ� (*.htm;*.html)|*.htm;*.html|�ۺ� HTML �ĵ� (*.mht;*.mhtml)|*.mht;*.mhtml|�ı��ļ� (*.txt)|*.txt"
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
        
        frmMain.cboZoom.Clear
        
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
            
            frmMain.cboZoom.AddItem strText
            
        Next i
        
        frmCode.brwQuickView.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DODEFAULT, 0, vntZoom ' ȡ�õ�ǰ���ű�
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

        ' �����Ǵ��� Get �ģ��ο���һЩ���ϣ�����д��
        ' frmCode.brwQuickView.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DODEFAULT, 0, vntZoom
        ' ������ô�Ķ��ᱨ������ pvaIn ��ʵ������Ĳ�������������� Get �Ĺ���û���κ����壩����������
        frmCode.brwQuickView.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DODEFAULT, CLng(100), 0    ' ǿ���������ű�Ϊ 100%
        glngDefaultZoom = 100
        strText = "100%"

        For i = 0 To frmMain.cboZoom.ListCount - 1
            If frmMain.cboZoom.List(i) = strText Then
                frmMain.cboZoom.ListIndex = i
            End If
        Next i
        
    End If
End Function

