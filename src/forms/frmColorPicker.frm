VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "ComCt232.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Begin VB.Form frmColorPicker 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "颜色选择器"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12705
   Icon            =   "frmColorPicker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   611
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   847
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picHex 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   240
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   33
      Top             =   5640
      Width           =   4335
      Begin VB.TextBox txtHex 
         Height          =   270
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   35
         Text            =   "#000000"
         Top             =   0
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "十六进制(&H)："
         Height          =   180
         Left            =   0
         TabIndex        =   34
         Top             =   45
         Width           =   1170
      End
   End
   Begin VB.PictureBox picHSL 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   120
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   23
      Top             =   7200
      Width           =   4335
      Begin ComCtl2.UpDown updL 
         Height          =   270
         Left            =   4080
         TabIndex        =   32
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   327681
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtL"
         BuddyDispid     =   196613
         OrigLeft        =   2520
         OrigTop         =   840
         OrigRight       =   2775
         OrigBottom      =   1095
         Max             =   240
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtL 
         Height          =   270
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   31
         Text            =   "0"
         Top             =   720
         Width           =   2880
      End
      Begin ComCtl2.UpDown updS 
         Height          =   270
         Left            =   4080
         TabIndex        =   29
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   327681
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtS"
         BuddyDispid     =   196614
         OrigLeft        =   3720
         OrigTop         =   360
         OrigRight       =   3975
         OrigBottom      =   495
         Max             =   240
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtS 
         Height          =   270
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   28
         Text            =   "0"
         Top             =   360
         Width           =   2880
      End
      Begin ComCtl2.UpDown updH 
         Height          =   270
         Left            =   4080
         TabIndex        =   26
         Top             =   0
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   327681
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtH"
         BuddyDispid     =   196615
         OrigLeft        =   165
         OrigTop         =   45
         OrigRight       =   420
         OrigBottom      =   75
         Max             =   240
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtH 
         Height          =   270
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   25
         Text            =   "0"
         Top             =   0
         Width           =   2880
      End
      Begin VB.Label lblL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "亮度(&L)："
         Height          =   180
         Left            =   0
         TabIndex        =   30
         Top             =   765
         Width           =   810
      End
      Begin VB.Label lblS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "饱和度(&S)："
         Height          =   180
         Left            =   0
         TabIndex        =   27
         Top             =   405
         Width           =   990
      End
      Begin VB.Label lblH 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "色调(&H)："
         Height          =   180
         Left            =   0
         TabIndex        =   24
         Top             =   45
         Width           =   810
      End
   End
   Begin VB.PictureBox picRGB 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   240
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   13
      Top             =   4560
      Width           =   4335
      Begin ComCtl2.UpDown updB 
         Height          =   270
         Left            =   4080
         TabIndex        =   22
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   327681
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtB"
         BuddyDispid     =   196620
         OrigLeft        =   3120
         OrigTop         =   1320
         OrigRight       =   3375
         OrigBottom      =   1455
         Max             =   255
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtB 
         Height          =   270
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   21
         Text            =   "0"
         Top             =   720
         Width           =   2880
      End
      Begin ComCtl2.UpDown updG 
         Height          =   270
         Left            =   4080
         TabIndex        =   19
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   327681
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtG"
         BuddyDispid     =   196621
         OrigLeft        =   3360
         OrigTop         =   1200
         OrigRight       =   3615
         OrigBottom      =   1335
         Max             =   255
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtG 
         Height          =   270
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   18
         Text            =   "0"
         Top             =   360
         Width           =   2880
      End
      Begin ComCtl2.UpDown updR 
         Height          =   270
         Left            =   4080
         TabIndex        =   16
         Top             =   0
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   327681
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtR"
         BuddyDispid     =   196622
         OrigLeft        =   2520
         OrigTop         =   240
         OrigRight       =   2775
         OrigBottom      =   375
         Max             =   255
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   15
         Text            =   "0"
         Top             =   0
         Width           =   2880
      End
      Begin VB.Label lblB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "蓝色(&B)："
         Height          =   180
         Left            =   0
         TabIndex        =   20
         Top             =   765
         Width           =   810
      End
      Begin VB.Label lblG 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "绿色(&G)："
         Height          =   180
         Left            =   0
         TabIndex        =   17
         Top             =   405
         Width           =   810
      End
      Begin VB.Label lblR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "红色(&R)："
         Height          =   180
         Left            =   0
         TabIndex        =   14
         Top             =   45
         Width           =   810
      End
   End
   Begin VB.PictureBox picColorMode 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   240
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   10
      Top             =   4200
      Width           =   4335
      Begin VB.ComboBox cboColorMode 
         Height          =   300
         ItemData        =   "frmColorPicker.frx":000C
         Left            =   1200
         List            =   "frmColorPicker.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   0
         Width           =   3135
      End
      Begin VB.Label lblColorMode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "颜色模式(&M)："
         Height          =   180
         Left            =   0
         TabIndex        =   11
         Top             =   45
         Width           =   1170
      End
   End
   Begin VB.PictureBox picLuminance 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   3960
      ScaleHeight     =   241
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   9
      Top             =   480
      Width           =   615
      Begin VB.Image imgCursor2 
         Enabled         =   0   'False
         Height          =   225
         Left            =   480
         Picture         =   "frmColorPicker.frx":0024
         Top             =   0
         Width           =   225
      End
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   495
      Left            =   12120
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picWebSafeColors 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   6360
      ScaleHeight     =   433
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   7
      Top             =   120
      Width           =   4335
      Begin VB.Image imgWebSafeColor 
         Height          =   375
         Index           =   0
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picColors 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   4920
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   4
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox picColorsRectangle 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   240
      Picture         =   "frmColorPicker.frx":009A
      ScaleHeight     =   241
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   241
      TabIndex        =   0
      Top             =   480
      Width           =   3615
      Begin VB.Image imgCursor1 
         Enabled         =   0   'False
         Height          =   225
         Left            =   0
         Picture         =   "frmColorPicker.frx":2AA70
         Top             =   0
         Width           =   225
      End
   End
   Begin ComctlLib.TabStrip tabMain 
      Height          =   6975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   12303
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Web 安全色"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "自定义"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNew 
      Alignment       =   2  'Center
      Caption         =   "新增"
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblCurrent 
      Alignment       =   2  'Center
      Caption         =   "当前"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   6120
      Width           =   1095
   End
End
Attribute VB_Name = "frmColorPicker"
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

Public lngColorRef As Long
Public idxWebSafeColor As Long

Public blnIsInitialized As Boolean

Private RGB_ As RGBColor
Private HSL_ As HSLColor

Private blnIsRGB As Boolean ' 用于判断颜色模式
Private blnIsDragging As Boolean ' 用于判断是否在拖动模式
Private blnIsAutoChanging As Boolean ' 用于判断是否在自动改变

Public blnIsSelected As Boolean ' 用于判断颜色是否选定

Private blnIsBGInitialized As Boolean

Private Sub cboColorMode_Click()
    Select Case cboColorMode.List(cboColorMode.ListIndex)
        Case "RGB"
            blnIsRGB = True
        Case "HSL"
            blnIsRGB = False
    End Select
    
    tabMain_Click
End Sub

Private Sub cmdCancel_Click()
    blnIsSelected = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    blnIsSelected = True
    Unload Me
End Sub

Private Sub Form_Initialize()
    Me.Width = 6400
    Me.Height = 7680
    
    InitWebSafeColorButtons
    
    picWebSafeColors.Move picColorsRectangle.Left, picColorsRectangle.Top
    picHSL.Move picRGB.Left, picRGB.Top
    blnIsRGB = True
    
    cboColorMode.ListIndex = 0
    
    tabMain_Click
End Sub

Private Sub InitWebSafeColorButtons()
    Dim i As Long
    
    For i = 0 To 215
        If i > 0 Then Load imgWebSafeColor(i)
        
        With imgWebSafeColor(i)
            .Stretch = True
            .Move (i - ((i \ 12) * 12)) * 24 + 2, _
                (i \ 12) * 24 + 2, _
                20, _
                20
            picTemp.Line (0, 0)-(32, 32), lngWebSafeColor(i), BF
            picTemp.Refresh
            Set .Picture = picTemp.Image
            .ToolTipText = "#" & CLRtoStr(lngWebSafeColor(i))
            .Visible = True
        End With
    Next i
    
End Sub

Private Sub Form_Paint()
    If blnIsBGInitialized = False Then
        'MsgBox 1
        blnIsBGInitialized = True
        
        tabMain.ZOrder 0
        Me.Refresh
        
        ' 通过重绘一遍背景来实现伪透明
        
        BitBlt picLuminance.hDC, 0, 0, picLuminance.ScaleWidth, picLuminance.ScaleHeight, Me.hDC, picLuminance.Left, picLuminance.Top, vbSrcCopy
        
        BitBlt picColorMode.hDC, 0, 0, picColorMode.ScaleWidth, picColorMode.ScaleHeight, Me.hDC, picColorMode.Left, picColorMode.Top, vbSrcCopy
        
        BitBlt picRGB.hDC, 0, 0, picRGB.ScaleWidth, picRGB.ScaleHeight, Me.hDC, picRGB.Left, picRGB.Top, vbSrcCopy

        BitBlt picHSL.hDC, 0, 0, picHSL.ScaleWidth, picHSL.ScaleHeight, Me.hDC, picHSL.Left, picHSL.Top, vbSrcCopy
        
        BitBlt picHex.hDC, 0, 0, picHex.ScaleWidth, picHex.ScaleHeight, Me.hDC, picHex.Left, picHex.Top, vbSrcCopy
        
        tabMain.ZOrder 1
    End If
End Sub

' 改成 Public 很重要
Public Sub imgWebSafeColor_Click(Index As Integer)


    picWebSafeColors.Line ((idxWebSafeColor - ((idxWebSafeColor \ 12) * 12)) * 24, (idxWebSafeColor \ 12) * 24)-((idxWebSafeColor - ((idxWebSafeColor \ 12) * 12)) * 24 + 24, (idxWebSafeColor \ 12) * 24 + 24), _
        picWebSafeColors.BackColor, BF
    picWebSafeColors.Line ((Index - ((Index \ 12) * 12)) * 24, (Index \ 12) * 24)-((Index - ((Index \ 12) * 12)) * 24 + 23, (Index \ 12) * 24 + 23), 0, BF
    picWebSafeColors.Refresh

    lngColorRef = lngWebSafeColor(Index)
    
    UpdateColorByColorRef lngColorRef
    UpdateLuminanceBar
    UpdateNewColorPicture
    UpdateTextBoxes
    UpdateCursors
    
    idxWebSafeColor = Index
    
    If blnIsInitialized = True Then picWebSafeColors.SetFocus
End Sub

Private Function UpdateNewColorPicture()
    picColors.Line (0, 0)-(73, 36), lngColorRef, BF
    picColors.Refresh
End Function

Private Function UpdateColorByHSL(ByVal wHue As Integer, ByVal wSaturation As Integer, ByVal wLuminance As Integer)
    With HSL_
        .H = wHue
        .S = wSaturation
        .L = wLuminance
        lngColorRef = ColorHLSToRGB(.H, .L, .S)
    End With
    
    With RGB_
        CLRtoRGB lngColorRef, .R, .G, .B
    End With

End Function

Private Function UpdateColorByColorRef(ByVal clrColorRef As Long)
    CLRtoRGB clrColorRef, RGB_.R, RGB_.G, RGB_.B
    UpdateColorByRGB RGB_.R, RGB_.G, RGB_.B
End Function

Private Function UpdateColorByRGB(ByVal R As Integer, ByVal G As Integer, ByVal B As Integer)
    With RGB_
        .R = R
        .G = G
        .B = B
    End With
    
    lngColorRef = RGB(R, G, B)
    
    With HSL_
        ColorRGBToHLS lngColorRef, .H, .L, .S
    End With
    
End Function

Private Function UpdateTextBoxes()
    With RGB_
        txtR.Text = .R
        txtG.Text = .G
        txtB.Text = .B
        
        txtR.Refresh
        txtG.Refresh
        txtB.Refresh
    End With
    
    With HSL_
        txtH.Text = .H
        txtS.Text = .S
        txtL.Text = .L
        
        txtH.Refresh
        txtS.Refresh
        txtL.Refresh
    End With
    
    txtHex = "#" & CLRtoStr(lngColorRef)
    txtHex.Refresh
    
End Function

Private Function UpdateLuminanceBar()
    Dim i As Long
    
    For i = 0 To 240
        picLuminance.Line (0, 240 - i)-(31, 240 - i), ColorHLSToRGB(HSL_.H, i, HSL_.S), BF
    Next i
    
    picLuminance.Refresh
End Function

Private Function UpdateCursors()
    imgCursor1.Move HSL_.H - 7, (240 - HSL_.S) - 7
    imgCursor2.Top = (240 - HSL_.L) - 7
End Function

Private Sub imgWebSafeColor_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub picColorsRectangle_DblClick()
    cmdOK_Click
End Sub

Private Sub picColorsRectangle_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft
            If HSL_.H > 0 Then HSL_.H = HSL_.H - 1
            blnIsDragging = True
            
        Case vbKeyRight
            If HSL_.H < 240 Then HSL_.H = HSL_.H + 1
            blnIsDragging = True
            
        Case vbKeyUp
            If HSL_.S < 240 Then HSL_.S = HSL_.S + 1
            blnIsDragging = True
            
        Case vbKeyDown
            If HSL_.S > 0 Then HSL_.S = HSL_.S - 1
            blnIsDragging = True
    End Select
    
    HSL_.L = 120
    
    UpdateColorByHSL HSL_.H, HSL_.S, HSL_.L
    UpdateLuminanceBar
    UpdateNewColorPicture
    UpdateCursors
    UpdateTextBoxes
End Sub

Private Sub picColorsRectangle_KeyUp(KeyCode As Integer, Shift As Integer)
    blnIsDragging = False
End Sub

' 改成 Public 很重要
Public Sub picColorsRectangle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = 1) And (0 <= X) And (X <= 240) And (0 <= Y) And (Y <= 240) Then
    
        blnIsDragging = True
    
        imgCursor1.Move X - 7, Y - 7
        lngColorRef = picColorsRectangle.Point(X, Y)
        UpdateColorByHSL X, 240 - Y, 120
        
        UpdateLuminanceBar
        UpdateNewColorPicture
        UpdateTextBoxes
        picLuminance_MouseDown 1, 0, 0, 120
    End If
End Sub

Private Sub picColorsRectangle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = 1) And (0 <= X) And (X <= 240) And (0 <= Y) And (Y <= 240) Then
        imgCursor1.Move X - 7, Y - 7
        lngColorRef = picColorsRectangle.Point(X, Y)
        UpdateColorByHSL X, 240 - Y, 120
        
        UpdateLuminanceBar
        UpdateNewColorPicture
        UpdateTextBoxes
    End If
End Sub

Private Sub picColorsRectangle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = 1) And (0 <= X) And (X <= 240) And (0 <= Y) And (Y <= 240) Then
    
        blnIsDragging = False
    
        imgCursor1.Move X - 7, Y - 7
        lngColorRef = picColorsRectangle.Point(X, Y)
        UpdateColorByHSL X, 240 - Y, 120
        
        UpdateLuminanceBar
        UpdateNewColorPicture
        UpdateTextBoxes
    End If
End Sub

Private Sub picLuminance_DblClick()
    cmdOK_Click
End Sub

Private Sub picLuminance_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            If HSL_.L < 240 Then HSL_.L = HSL_.L + 1
            blnIsDragging = True
        
        Case vbKeyDown
            If HSL_.L > 0 Then HSL_.L = HSL_.L - 1
            blnIsDragging = True
    End Select
    
    UpdateColorByHSL HSL_.H, HSL_.S, HSL_.L
    UpdateLuminanceBar
    UpdateNewColorPicture
    UpdateCursors
    UpdateTextBoxes
End Sub

Private Sub picLuminance_KeyUp(KeyCode As Integer, Shift As Integer)
    blnIsDragging = False
End Sub

' 改成 Public 很重要
Public Sub picLuminance_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = 1) And (0 <= X) And (X <= 240) And (0 <= Y) And (Y <= 240) Then
    
        blnIsDragging = True
    
        imgCursor2.Top = Y - 7
        UpdateColorByHSL HSL_.H, HSL_.S, 240 - Y
        lngColorRef = RGB(RGB_.R, RGB_.G, RGB_.B)
        
        UpdateNewColorPicture
        UpdateTextBoxes
    End If
End Sub

Private Sub picLuminance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = 1) And (0 <= X) And (X <= 240) And (0 <= Y) And (Y <= 240) Then
        imgCursor2.Top = Y - 7
        UpdateColorByHSL HSL_.H, HSL_.S, 240 - Y
        lngColorRef = RGB(RGB_.R, RGB_.G, RGB_.B)
        
        UpdateNewColorPicture
        UpdateTextBoxes
    End If
End Sub

Private Sub picLuminance_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = 1) And (0 <= X) And (X <= 240) And (0 <= Y) And (Y <= 240) Then
    
        blnIsDragging = False
        
        imgCursor2.Top = Y - 7
        UpdateColorByHSL HSL_.H, HSL_.S, 240 - Y
        lngColorRef = RGB(RGB_.R, RGB_.G, RGB_.B)
        
        UpdateNewColorPicture
        UpdateTextBoxes
    End If
End Sub

Private Sub picWebSafeColors_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft
            If idxWebSafeColor Mod 12 <> 0 Then
                imgWebSafeColor_Click idxWebSafeColor - 1
            End If
            
        Case vbKeyRight
            If idxWebSafeColor Mod 12 <> 11 Then
                imgWebSafeColor_Click idxWebSafeColor + 1
            End If
            
        Case vbKeyUp
            If idxWebSafeColor \ 12 <> 0 Then
                imgWebSafeColor_Click idxWebSafeColor - 12
            End If
            
        Case vbKeyDown
            If idxWebSafeColor \ 12 <> 17 Then
                imgWebSafeColor_Click idxWebSafeColor + 12
            End If
    End Select
End Sub

Private Sub tabMain_Click()

    Dim i As Long

    Select Case tabMain.SelectedItem.Index
        Case 1
            'picWebSafeColors.ZOrder 0
            picWebSafeColors.Visible = True
            
            picColorsRectangle.Visible = False
            picLuminance.Visible = False
            
            picColorMode.Visible = False
            picHex.Visible = False
            picHSL.Visible = False
            picRGB.Visible = False
            
            picWebSafeColors.Cls
            idxWebSafeColor = 0
            For i = 0 To 215
                If lngColorRef = lngWebSafeColor(i) Then
                    imgWebSafeColor_Click CInt(i)
                End If
            Next i
            
            If blnIsInitialized = True Then picWebSafeColors.SetFocus
            
        Case 2
            picWebSafeColors.Visible = False
            'picColorsRectangle.ZOrder 0
            'picLuminance.ZOrder 0
            
            picColorsRectangle.Visible = True
            picLuminance.Visible = True
            
            picColorMode.Visible = True
            picHex.Visible = True
            
            If blnIsRGB = True Then
                picRGB.Visible = True
                picHSL.Visible = False
            Else
                picHSL.Visible = True
                picRGB.Visible = False
            End If
            
            With RGB_
                txtR.Text = .R
                txtG.Text = .G
                txtB.Text = .B
            End With
            
            With HSL_
                txtH.Text = .H
                txtS.Text = .S
                txtL.Text = .L
            End With
            
            UpdateLuminanceBar
            
            If blnIsInitialized = True Then picColorsRectangle.SetFocus
    End Select
End Sub

Private Sub txtHex_Change()
    If (blnIsDragging = False) And (Len(txtHex.Text) = 7) Then txtHex_LostFocus
End Sub

Private Sub txtHex_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 35, 48 To 57, 65 To 70
            ' pass

        Case Asc(vbCr)
            txtHex_LostFocus

        Case 97 To 102
            KeyAscii = KeyAscii - 32

        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub txtHex_LostFocus()
    If Len(txtHex.Text) <> 7 Then GoTo SubError
    If Len(Replace(txtHex.Text, "#", "")) <> 6 Then GoTo SubError
    If Mid(txtHex.Text, 1, 1) <> "#" Then GoTo SubError

    lngColorRef = StrtoCLR(Replace(txtHex.Text, "#", ""))
    UpdateColorByColorRef lngColorRef

    blnIsAutoChanging = True
    UpdateTextBoxes
    blnIsAutoChanging = False

    UpdateLuminanceBar
    UpdateNewColorPicture
    UpdateCursors

    Exit Sub

SubError:
    txtHex.Text = "#" & CLRtoStr(lngColorRef)
    Exit Sub
End Sub

Private Sub txtR_Change()
    If blnIsAutoChanging = True Then Exit Sub
    If txtR.Text = "" Then
        RGB_.R = 0
    ElseIf CLng(txtR.Text) > 255 Then
        txtR.Text = 255
    End If

    If txtR.Text <> "" Then RGB_.R = txtR.Text
    If blnIsDragging = False Then UpdateColorByRGB RGB_.R, RGB_.G, RGB_.B
    UpdateLuminanceBar
    UpdateNewColorPicture
    UpdateCursors
End Sub

Private Sub txtR_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
            ' pass

        Case Asc(vbCr)
            txtR_LostFocus

        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtR_LostFocus()
    If txtR.Text = "" Then txtR.Text = 0
    UpdateTextBoxes
End Sub

Private Sub txtG_Change()
    If blnIsAutoChanging = True Then Exit Sub
    If txtG.Text = "" Then
        RGB_.G = 0
    ElseIf CLng(txtG.Text) > 255 Then
        txtG.Text = 255
    End If

    If txtG.Text <> "" Then RGB_.G = txtG.Text
    If blnIsDragging = False Then UpdateColorByRGB RGB_.R, RGB_.G, RGB_.B
    UpdateLuminanceBar
    UpdateNewColorPicture
    UpdateCursors
End Sub

Private Sub txtG_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
            ' pass

        Case Asc(vbCr)
            txtG_LostFocus

        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtG_LostFocus()
    If txtG.Text = "" Then txtG.Text = 0
    UpdateTextBoxes
End Sub

Private Sub txtB_Change()
    If blnIsAutoChanging = True Then Exit Sub
    If txtB.Text = "" Then
        RGB_.B = 0
    ElseIf CLng(txtB.Text) > 255 Then
        txtB.Text = 255
    End If

    If txtB.Text <> "" Then RGB_.B = txtB.Text
    If blnIsDragging = False Then UpdateColorByRGB RGB_.R, RGB_.G, RGB_.B
    UpdateLuminanceBar
    UpdateNewColorPicture
    UpdateCursors
End Sub

Private Sub txtB_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
            ' pass

        Case Asc(vbCr)
            txtB_LostFocus

        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtB_LostFocus()
    If txtB.Text = "" Then txtB.Text = 0
    UpdateTextBoxes
End Sub

Private Sub txtH_Change()
    If blnIsAutoChanging = True Then Exit Sub
    If txtH.Text = "" Then
        HSL_.H = 0
    ElseIf CLng(txtH.Text) > 240 Then
        txtH.Text = 240
    End If

    If txtH.Text <> "" Then HSL_.H = txtH.Text
    If blnIsDragging = False Then UpdateColorByHSL HSL_.H, HSL_.S, HSL_.L
    UpdateLuminanceBar
    UpdateNewColorPicture
    UpdateCursors
End Sub

Private Sub txtH_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
            ' pass

        Case Asc(vbCr)
            txtH_LostFocus

        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtH_LostFocus()
    If txtH.Text = "" Then txtH.Text = 0
    UpdateTextBoxes
End Sub

Private Sub txtS_Change()
    If blnIsAutoChanging = True Then Exit Sub
    If txtS.Text = "" Then
        HSL_.S = 0
    ElseIf CLng(txtS.Text) > 240 Then
        txtS.Text = 240
    End If

    If txtS.Text <> "" Then HSL_.S = txtS.Text
    If blnIsDragging = False Then UpdateColorByHSL HSL_.H, HSL_.S, HSL_.L
    UpdateLuminanceBar
    UpdateNewColorPicture
    UpdateCursors
End Sub

Private Sub txtS_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
            ' pass

        Case Asc(vbCr)
            txtS_LostFocus

        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtS_LostFocus()
    If txtS.Text = "" Then txtS.Text = 0
    UpdateTextBoxes
End Sub

Private Sub txtL_Change()
    If blnIsAutoChanging = True Then Exit Sub
    If txtL.Text = "" Then
        HSL_.L = 0
    ElseIf CLng(txtL.Text) > 240 Then
        txtL.Text = 240
    End If

    If txtL.Text <> "" Then HSL_.L = txtL.Text
    If blnIsDragging = False Then UpdateColorByHSL HSL_.H, HSL_.S, HSL_.L
    UpdateLuminanceBar
    UpdateNewColorPicture
    UpdateCursors
End Sub

Private Sub txtL_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
            ' pass

        Case Asc(vbCr)
            txtL_LostFocus

        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtL_LostFocus()
    If txtL.Text = "" Then txtL.Text = 0
    UpdateTextBoxes
End Sub



