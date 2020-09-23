VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto Complete"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5745
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin Autocomplete.ExtedListbox EXT 
      Height          =   2850
      Left            =   2985
      TabIndex        =   3
      Top             =   480
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   5027
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      List0           =   "Pankaj"
      ListIndex       =   -1
      BackStyle       =   0
   End
   Begin VB.TextBox Text1 
      Height          =   3105
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":038A
      Top             =   990
      Width           =   2460
   End
   Begin VB.Line Line4 
      X1              =   3015
      X2              =   3210
      Y1              =   1035
      Y2              =   1020
   End
   Begin VB.Line Line3 
      X1              =   3210
      X2              =   3180
      Y1              =   1020
      Y2              =   1170
   End
   Begin VB.Line Line2 
      X1              =   2505
      X2              =   3225
      Y1              =   1380
      Y2              =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STEP 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   3015
      TabIndex        =   5
      Top             =   0
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STEP 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   165
      TabIndex        =   4
      Top             =   30
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Now type any word here (EX: Word1)"
      Height          =   195
      Left            =   3015
      TabIndex        =   2
      Top             =   225
      Width           =   2640
   End
   Begin VB.Line Line1 
      X1              =   2880
      X2              =   2880
      Y1              =   30
      Y2              =   4410
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type here differemt word. Seperating each word by a ""Enter"" Key"
      Height          =   630
      Left            =   225
      TabIndex        =   1
      Top             =   300
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim Mys As Mystring ' Type variable decleared in Module1
Dim Temp As String

EXT.Txtheight 315
EXT.Clear

Temp = Me.Text1.Text
Temp = Replace$(Temp, Chr(13), "~")
Temp = Replace$(Temp, Chr(10), "")

Mys = ToArray(Temp, "~")

Dim Loop1 As Integer
For Loop1 = 1 To Mys.Myarray.Count
    EXT.AddItem "" & Mys.Myarray.item(Loop1) & ""
    
Next Loop1

End Sub

Private Sub Text1_LostFocus()
Dim Mys As Mystring ' Type variable decleared in Module1
Dim Temp As String

EXT.Clear

Temp = Me.Text1.Text
Temp = Replace$(Temp, Chr(13), "~")
Temp = Replace$(Temp, Chr(10), "")

Mys = ToArray(Temp, "~")

Dim Loop1 As Integer
For Loop1 = 1 To Mys.Myarray.Count
    EXT.AddItem "" & Mys.Myarray.item(Loop1) & ""
    
Next Loop1
End Sub
