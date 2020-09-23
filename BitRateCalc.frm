VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DivX Bitrate Calculator"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "BitRateCalc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll1 
      Height          =   255
      Index           =   2
      Left            =   2280
      Max             =   9999
      TabIndex        =   12
      Top             =   1680
      Value           =   9999
      Width           =   255
   End
   Begin VB.TextBox TimeBox 
      Height          =   285
      Index           =   2
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   5
      Text            =   "0"
      Top             =   1680
      Width           =   495
   End
   Begin VB.ComboBox AudioBox 
      Height          =   315
      Index           =   2
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
      Width           =   2895
   End
   Begin VB.ComboBox AudioBox 
      Height          =   315
      Index           =   1
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Width           =   2895
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   255
      Index           =   1
      Left            =   3000
      Max             =   59
      TabIndex        =   9
      Top             =   120
      Value           =   59
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   255
      Index           =   0
      Left            =   2280
      Max             =   999
      TabIndex        =   8
      Top             =   120
      Value           =   999
      Width           =   255
   End
   Begin VB.ComboBox AudioBox 
      Height          =   315
      Index           =   0
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox TimeBox 
      Height          =   285
      Index           =   1
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TimeBox 
      Height          =   285
      Index           =   0
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "(DivX Codec)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "1000 bits/Kbit:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      DrawMode        =   1  'Blackness
      Height          =   1695
      Left            =   120
      Top             =   2040
      Width           =   4575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Video Bitrate"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2040
      Width           =   4575
   End
   Begin VB.Label BitrateBox2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1560
      TabIndex        =   17
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Label BitrateBox3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label BitrateBox1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label Label5 
      Caption         =   "1024 bits/Kbit:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label BytesBox 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Min. / Sec."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Target Size (MB)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Audio Track(s)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Video Length"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim A%, B%, C%, D%, E%
Dim AudioRates$(20)
Dim TimeMax%(2)
Dim ActualBytes!, VBitrate!
Dim Secs!, ABitrate!
Dim ProgramPath$, FF%, Temp$
Dim FormLoading As Boolean

Private Sub AudioBox_Click(Index As Integer)

UpdateVideoBitrate

End Sub

Private Sub Form_Load()

FormLoading = True
ProgramPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")

AudioRates(0) = "448 Kbit/sec."
AudioRates(1) = "384 Kbit/sec."
AudioRates(2) = "320 Kbit/sec."
AudioRates(3) = "256 Kbit/sec."
AudioRates(4) = "224 Kbit/sec."
AudioRates(5) = "192 Kbit/sec."
AudioRates(6) = "160 Kbit/sec."
AudioRates(7) = "128 Kbit/sec."
AudioRates(8) = "112 Kbit/sec."
AudioRates(9) = "96 Kbit/sec."
AudioRates(10) = "64 Kbit/sec."
AudioRates(11) = "56 Kbit/sec."
AudioRates(12) = "48 Kbit/sec."
AudioRates(13) = "40 Kbit/sec."
AudioRates(14) = "32 Kbit/sec."
AudioRates(15) = "24 Kbit/sec."
AudioRates(16) = "20 Kbit/sec."
AudioRates(17) = "18 Kbit/sec."
AudioRates(18) = "16 Kbit/sec."
AudioRates(19) = "8 Kbit/sec."
AudioRates(20) = "No Audio"

For A = 0 To 2
    For B = 0 To 20
        AudioBox(A).AddItem AudioRates(B)
    Next B
    AudioBox(A).ListIndex = 20
Next A

TimeMax(0) = 999
TimeMax(1) = 59
TimeMax(2) = 9999

If Dir(ProgramPath & "BRCalc.dat") <> "" Then
    FF = FreeFile
    Open ProgramPath & "BRCalc.dat" For Input As #FF
        For A = 0 To 2
            Input #FF, Temp
            TimeBox(A).Text = Trim(Temp)
        Next A
        For A = 0 To 2
            Input #FF, B
            AudioBox(A).ListIndex = B
        Next A
    Close #FF
End If

FormLoading = False
UpdateVideoBitrate

End Sub

Private Sub Form_Unload(Cancel As Integer)

SaveSettings

End Sub

Private Sub TimeBox_Change(Index As Integer)

If Index = 1 And Val(TimeBox(1).Text) > 59 Then
    TimeBox(1).Text = "59"
End If

If Index = 2 Then
    ActualBytes = Val(TimeBox(2).Text) * 1048576
    BytesBox.Caption = "(" & FormatNumber(ActualBytes, 0, , , vbTrue) & " Bytes)"
End If

VScroll1(Index).Value = TimeMax(Index) - Val(TimeBox(Index).Text)

UpdateVideoBitrate

End Sub

Private Sub TimeBox_KeyPress(Index As Integer, KeyAscii As Integer)

If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
    KeyAscii = 0
End If

End Sub

Private Sub VScroll1_Change(Index As Integer)

TimeBox(Index).Text = TimeMax(Index) - VScroll1(Index).Value

End Sub

Private Sub VScroll1_GotFocus(Index As Integer)
TimeBox(Index).SetFocus
End Sub

Private Sub UpdateVideoBitrate()

If FormLoading Then Exit Sub

Secs = Val(TimeBox(0).Text) * 60 + Val(TimeBox(1).Text)
ABitrate = 0
For A = 0 To 2
    ABitrate = ABitrate + Val(AudioBox(A).Text)
Next A
ABitrate = ABitrate * 1024

If Secs > 0 Then
    VBitrate = ActualBytes * 8 / Secs - ABitrate
Else
    VBitrate = 0
End If

If VBitrate > 0 Then
    BitrateBox1.Caption = FormatNumber(VBitrate / 1024, 1, , , vbTrue) & " Kbit/sec."
    BitrateBox2.Caption = FormatNumber(VBitrate / 1000, 1, , , vbTrue) & " Kbit/sec."
    BitrateBox3.Caption = "(" & FormatNumber(VBitrate, 1, , , vbTrue) & " bits/sec.)"
Else
    BitrateBox1.Caption = "Impossible"
    BitrateBox2.Caption = ""
    BitrateBox3.Caption = ""
End If

End Sub

Private Sub SaveSettings()

FF = FreeFile
Open ProgramPath & "BRCalc.dat" For Output As #FF
    For A = 0 To 2
        Print #FF, TimeBox(A).Text
    Next A
    For A = 0 To 2
        Print #FF, AudioBox(A).ListIndex
    Next A
Close #FF

End Sub
