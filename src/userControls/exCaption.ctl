VERSION 5.00
Begin VB.UserControl exCaption 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "exCaption.ctx":0000
   Begin VB.PictureBox picCloseButton 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1080
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   120
      Width           =   240
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   1320
      ScaleHeight     =   3615
      ScaleWidth      =   3495
      TabIndex        =   2
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "exCaption"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   810
   End
End
Attribute VB_Name = "exCaption"
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

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private blnIsCapturing As Boolean

'�¼�����:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "���û���һ�������ϰ��²��ͷ���갴ťʱ������"
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "���û���һ�������ϰ��²��ͷ���갴ť���ٴΰ��²��ͷ���갴ťʱ������"
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "���û���ӵ�н���Ķ����ϰ��������ʱ������"
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "���û����º��ͷ� ANSI ��ʱ������"
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "���û���ӵ�н���Ķ������ͷż�ʱ������"
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "���û���ӵ�н���Ķ����ϰ�����갴ťʱ������"
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "���û��ƶ����ʱ������"
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "���û���ӵ�н���Ķ������ͷ���귢����"
Event OnClose()

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "����/���ö������ı���ͼ�εı���ɫ��"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    picCloseButton.BackColor = UserControl.BackColor
    RepaintCtrl
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "����/���ö������ı���ͼ�ε�ǰ��ɫ��"
    ForeColor = lblCaption.ForeColor
    RepaintCtrl
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblCaption.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    RepaintCtrl
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "����/����һ��ֵ������һ�������Ƿ���Ӧ�û������¼���"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    RepaintCtrl
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "����һ�� Font ����"
Attribute Font.VB_UserMemId = -512
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    PropertyChanged "Font"
    RepaintCtrl
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "ǿ����ȫ�ػ�һ������"
    UserControl.Refresh
    RepaintCtrl
End Sub

Private Sub picCloseButton_Click()
    RaiseEvent OnClose
    picCloseButton.BackColor = UserControl.BackColor
    RepaintCtrl
    If blnIsCapturing = True Then ReleaseCapture
    blnIsCapturing = False
End Sub

Private Sub picCloseButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim R As Long
    Dim G As Long
    Dim B As Long
    Dim A As Long
    
    R = UserControl.BackColor And &HFF&
    G = (UserControl.BackColor And &HFF00&) \ &H100&
    B = (UserControl.BackColor And &HFF0000) \ &H10000
    A = 64

    If Button = 0 Or Button = 1 Then
        If (X < 0) Or (X > picCloseButton.ScaleWidth) Or (Y < 0) Or (X > picCloseButton.ScaleHeight) Then
            picCloseButton.BackColor = UserControl.BackColor
            RepaintCtrl
            If blnIsCapturing = True Then ReleaseCapture
            blnIsCapturing = False
        Else
            picCloseButton.BackColor = RGB(&HFF& * A / 255 + R * (255 - A) / 255, &HFF& * A / 255 + G * (255 - A) / 255, &HFF& * A / 255 + B * (255 - A) / 255)
            RepaintCtrl
            If blnIsCapturing = False Then SetCapture picCloseButton.hwnd
            blnIsCapturing = True
        End If
    End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'�Ӵ������м�������ֵ
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCaption.Caption = PropBag.ReadProperty("Caption", "exCaption")
    
    picCloseButton.BackColor = UserControl.BackColor
    picMask.BackColor = UserControl.BackColor
End Sub

Private Sub UserControl_Resize()
    RepaintCtrl
End Sub

Private Sub UserControl_Show()
    RepaintCtrl
End Sub

'������ֵд���洢��
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "exCaption")
End Sub

Private Function RepaintCtrl()
    Dim lngGray As Long
    Dim clrColor As Long

    lblCaption.Move 4, (UserControl.ScaleHeight - lblCaption.Height) / 2

    With picCloseButton
        
        .Cls
        .Move UserControl.ScaleWidth - 20, (UserControl.ScaleHeight - 16) / 2
        
        lngGray = CLng(0.299 * (UserControl.BackColor And &HFF&) + 0.587 * ((UserControl.BackColor And &HFF00&) \ &H100&) + 0.114 * ((UserControl.BackColor And &HFF0000) \ &H10000))
        
        If lngGray > 127 Then
            clrColor = vbBlack
        Else
            clrColor = vbWhite
        End If
        
        ' ����ʡ�� picCloseButton
        picCloseButton.DrawWidth = 2
        picCloseButton.Line (0, 0)-(15, 15), clrColor
        picCloseButton.Line (15, 0)-(0, 15), clrColor
        picCloseButton.DrawWidth = 1
        picCloseButton.Line (0, 0)-(15, 15), .BackColor, B
        picCloseButton.Line (1, 1)-(14, 14), .BackColor, B
        picCloseButton.Line (2, 2)-(13, 13), .BackColor, B
    End With
    
    picMask.BackColor = UserControl.BackColor
    picMask.Move picCloseButton.Left, 0, 32, UserControl.ScaleHeight
    
    UserControl.Refresh
    
End Function
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "����/���ö���ı������л�ͼ��������ı���"
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

