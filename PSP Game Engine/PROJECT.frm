VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{E6EC7E0A-5B57-4783-88F1-27C0D50A5B92}#3.0#0"; "POWERB~1.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PSP GAME ENGINE"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9630
   Icon            =   "PROJECT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Height          =   285
      Left            =   7200
      TabIndex        =   17
      Text            =   "?"
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Height          =   285
      Left            =   7200
      TabIndex        =   16
      Text            =   "?"
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Height          =   285
      Left            =   7920
      TabIndex        =   13
      Text            =   "X.YZ"
      Top             =   960
      Width           =   735
   End
   Begin POWERBUTTONLib.PowerButton PowerButton1 
      Height          =   375
      Left            =   7920
      TabIndex        =   12
      Top             =   1320
      Width           =   735
      _Version        =   196608
      _ExtentX        =   1296
      _ExtentY        =   661
      _StockProps     =   71
      Caption         =   "Apply"
      BackColor       =   14933984
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PROJECT.frx":1CCA
      MouseInPicture  =   "PROJECT.frx":1CE6
      DownPicture     =   "PROJECT.frx":1D02
      DisabledPicture =   "PROJECT.frx":1D1E
      MouseIcon       =   "PROJECT.frx":1D3A
      _ClickSoundMemory=   "PROJECT.frx":1D56
      _MouseInSoundMemory=   "PROJECT.frx":1D6E
      BeginProperty MouseInFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "GAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      TabIndex        =   6
      Text            =   "X.YZ"
      Top             =   2520
      Width           =   735
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   360
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Text            =   "2.80"
      Top             =   3360
      Width           =   735
   End
   Begin POWERBUTTONLib.PowerButton Command2 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6735
      _Version        =   196608
      _ExtentX        =   11880
      _ExtentY        =   1296
      _StockProps     =   71
      Caption         =   "Open EBOOT Game File"
      BackColor       =   14933984
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PROJECT.frx":1D86
      MouseInPicture  =   "PROJECT.frx":1DA2
      DownPicture     =   "PROJECT.frx":1DBE
      DisabledPicture =   "PROJECT.frx":1DDA
      MouseIcon       =   "PROJECT.frx":1DF6
      _ClickSoundMemory=   "PROJECT.frx":1E12
      _MouseInSoundMemory=   "PROJECT.frx":1E2A
      BeginProperty MouseInFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin POWERBUTTONLib.PowerButton Command1 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   6735
      _Version        =   196608
      _ExtentX        =   11880
      _ExtentY        =   1296
      _StockProps     =   71
      Caption         =   "Start Converting"
      BackColor       =   14933984
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PROJECT.frx":1E42
      MouseInPicture  =   "PROJECT.frx":1E5E
      DownPicture     =   "PROJECT.frx":1E7A
      DisabledPicture =   "PROJECT.frx":1E96
      MouseIcon       =   "PROJECT.frx":1EB2
      _ClickSoundMemory=   "PROJECT.frx":1ECE
      _MouseInSoundMemory=   "PROJECT.frx":1EE6
      BeginProperty MouseInFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CC 
      Left            =   120
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FileName        =   "EBOOT.PBP"
      Filter          =   "PBP"
   End
   Begin POWERBUTTONLib.PowerButton PowerButton2 
      Height          =   375
      Left            =   7920
      TabIndex        =   15
      Top             =   2640
      Width           =   735
      _Version        =   196608
      _ExtentX        =   1296
      _ExtentY        =   661
      _StockProps     =   71
      Caption         =   "Apply"
      BackColor       =   14933984
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PROJECT.frx":1EFE
      MouseInPicture  =   "PROJECT.frx":1F1A
      DownPicture     =   "PROJECT.frx":1F36
      DisabledPicture =   "PROJECT.frx":1F52
      MouseIcon       =   "PROJECT.frx":1F6E
      _ClickSoundMemory=   "PROJECT.frx":1F8A
      _MouseInSoundMemory=   "PROJECT.frx":1FA2
      BeginProperty MouseInFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin POWERBUTTONLib.PowerButton PowerButton3 
      Height          =   375
      Left            =   7920
      TabIndex        =   18
      Top             =   3960
      Width           =   735
      _Version        =   196608
      _ExtentX        =   1296
      _ExtentY        =   661
      _StockProps     =   71
      Caption         =   "Apply"
      BackColor       =   14933984
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PROJECT.frx":1FBA
      MouseInPicture  =   "PROJECT.frx":1FD6
      DownPicture     =   "PROJECT.frx":1FF2
      DisabledPicture =   "PROJECT.frx":200E
      MouseIcon       =   "PROJECT.frx":202A
      _ClickSoundMemory=   "PROJECT.frx":2046
      _MouseInSoundMemory=   "PROJECT.frx":205E
      BeginProperty MouseInFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   19
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game Tittle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   14
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   11
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6960
      TabIndex        =   10
      Top             =   120
      Width           =   2655
   End
   Begin VB.Line Line3 
      X1              =   6960
      X2              =   6960
      Y1              =   120
      Y2              =   4560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00008000&
      X1              =   120
      X2              =   6840
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Last Update Wanted"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   2160
      Width           =   6975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New Update Setting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      Width           =   6975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      X1              =   120
      X2              =   6840
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Not Selected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OO As Integer

Private Sub Command1_Click()
On Error GoTo 455
If Not Text2.Text Like "#.##" Then
Text2.Text = Format(Text2.Text, "0.00")
End If
Open App.Path + "\Temp\File_" + Format(OO, "0000000000") + ".Txt" For Binary As #1
Put #1, 1, Text2.Text
Close #1

Dim CHAR As Byte
Dim CHA() As Byte
Dim X As String
Dim ST As Long
Dim I As Integer
ST = 1
File1.Pattern = "*.*"
File1.Path = App.Path + "\Temp\"
File1.Refresh
X = ".PBP"
Open App.Path + "\Temp" + X For Binary As #1
For I = 0 To File1.ListCount - 1
Open File1.Path + "\" + File1.List(I) For Binary As #2
X = Mid(File1.List(I), 6, Len(File1.List(I)) - 9)
ST = Int(X)
DoEvents
If LOF(2) = 0 Then
GoTo 6
End If
ReDim CHA(LOF(2) - 1)
Get #2, 1, CHA
Put #1, ST, CHA
Close #2
6:
Next
Close #1
On Error GoTo 455
CC.FileName = "EBOOT"
CC.ShowSave
If Right(UCase(CC.FileName), Len(".PBP")) = UCase(".PBP") Then
FileCopy App.Path + "\Temp" + ".PBP", CC.FileName
Else
FileCopy App.Path + "\Temp" + ".PBP", CC.FileName + ".PBP"
End If
455:
Exit Sub
End Sub

Private Sub Command2_Click()
On Error GoTo 5
CC.ShowOpen
FileCopy CC.FileName, App.Path + "\Temp\File_0000000001.Dat"
Open App.Path + "\Temp\File_0000000001.Dat" For Binary As #2
Dim CHAR As Byte
Get #2, OO, CHAR
Text1.Text = Chr(CHAR)
Get #2, OO + 1, CHAR
Text1.Text = Text1.Text + Chr(CHAR)
Get #2, OO + 2, CHAR
Text1.Text = Text1.Text + Chr(CHAR)
Get #2, OO + 3, CHAR
Text1.Text = Text1.Text + Chr(CHAR)
If Option1.Value = True Then
Get #2, 325, CHAR
Text3.Text = Chr(CHAR)
Get #2, 326, CHAR
Text3.Text = Text3.Text + Chr(CHAR)
Get #2, 327, CHAR
Text3.Text = Text3.Text + Chr(CHAR)
Get #2, 328, CHAR
Text3.Text = Text3.Text + Chr(CHAR)

Text4.Text = ""
For I = 0 To 50
Get #2, 352 + I, CHAR
Text4.Text = Text4.Text + Chr(CHAR)
Next

Text5.Text = ""
For I = 0 To 14
Get #2, 309 + I, CHAR
Text5.Text = Text5.Text + Chr(CHAR)
Next

End If
Close #2
5:
End Sub

Private Sub Form_Load()
OO = 337
End Sub

Private Sub Option1_Click()
OO = 337
End Sub

Private Sub Option2_Click()
OO = 2113
End Sub

Private Sub PowerButton1_Click()
If Not Text3.Text Like "#.##" Then
Text3.Text = Format(Text3.Text, "0.00")
End If
Open App.Path + "\Temp\File_" + Format(325, "0000000000") + ".Txt" For Binary As #1
Put #1, 1, Text3.Text
Close #1
End Sub

Private Sub PowerButton2_Click()
Open App.Path + "\Temp\File_" + Format(352, "0000000000") + ".Txt" For Binary As #1
Put #1, 1, Text4.Text
Close #1
End Sub

Private Sub PowerButton3_Click()
Open App.Path + "\Temp\File_" + Format(309, "0000000000") + ".Txt" For Binary As #1
Put #1, 1, Text5.Text
Close #1
End Sub
