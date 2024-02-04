VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Draw Colors Rectangle"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6960
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   321
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   464
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   4320
      Width           =   2055
   End
   Begin VB.PictureBox picColorsRectangle 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   120
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
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

Private Sub cmdDraw_Click()
    Dim h As Long
    Dim s As Long
    
    For h = 0 To 240
        For s = 0 To 240
            picColorsRectangle.PSet (h, 240 - s), ColorHLSToRGB(h, 120, s)
        Next s
    Next h
    
    picColorsRectangle.Refresh

End Sub

Private Sub cmdSave_Click()
    SavePicture picColorsRectangle.Image, "ColorsRectangle.bmp"
End Sub

Private Sub Form_Load()
    Me.Show
    Me.ScaleMode = vbPixels
    picColorsRectangle.ScaleMode = vbPixels
    picColorsRectangle.Width = picColorsRectangle.Width - picColorsRectangle.ScaleWidth + 241 ' Ëõ·Å 100%
    picColorsRectangle.Height = picColorsRectangle.Height - picColorsRectangle.ScaleHeight + 241
    
End Sub

