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

' TO DO
' ���ԣ�IHTMLEditDesigner �ػ��¼�
' (ʵ�鷢�ֲ�����)

Option Explicit

Sub Main()

    If App.LogMode = 1 Then
        If Dir(App.Path & "\QZHE.chm", vbNormal) <> "" Then
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

    frmMsg.PrintLog Time, "App.StatusText", "��ɫ�������ʼ���ɹ�"
    
    InitCommonControls
    
    frmMain.Show
    
    If Command <> "" Then
        gstrFileName = Command
        
        gstrDocHTML = OpenHTMLDoc(gstrFileName)
        
        frmCode.tabMain.Tabs(2).Selected = True
        frmCode.eEditor.Value = gstrDocHTML
        frmCode.tabMain_Click
        frmCode.UpdateFormCaption gstrFileName
    End If

End Sub

Public Sub AtExit()
    On Error Resume Next
    CreateObject("WScript.Shell").RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION\" & App.EXEName + ".exe"
    End
End Sub
