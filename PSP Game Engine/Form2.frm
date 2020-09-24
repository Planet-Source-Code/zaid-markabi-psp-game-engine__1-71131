VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2085
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2760
      Top             =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      X1              =   -120
      X2              =   120
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Scanning Files ..."
      BeginProperty Font 
         Name            =   "MS Dialog Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ver 1.0"
      BeginProperty Font 
         Name            =   "MS Dialog Light"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   135
      Left            =   5160
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PSP GAME ENGINE"
      BeginProperty Font 
         Name            =   "MS Dialog Light"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
On Error Resume Next
If Line1.X2 < Me.Width Then
Line1.X2 = Line1.X2 + (Rnd * Rnd * 300)
Else
Kill App.Path + "\Temp\File_0000000001.Dat"
Kill App.Path + "\Temp\File_0000000352.Txt"
Kill App.Path + "\Temp\File_0000000337.Txt"
Kill App.Path + "\Temp\File_0000000309.Txt"
Kill App.Path + "\Temp\File_0000000325.Txt"
Form1.Show
Unload Me
End If
End Sub
