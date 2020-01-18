VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Bit - Byte Converter"
   ClientHeight    =   5805
   ClientLeft      =   5310
   ClientTop       =   2940
   ClientWidth     =   10335
   Icon            =   "Bit - Byte Converter.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   10335
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   11
      Text            =   "Select Unit"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "GE Curviture"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Convert"
      BeginProperty Font 
         Name            =   "GE Curviture"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   3840
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   3240
      Width           =   4095
   End
   Begin VB.OptionButton Option4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "TB"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.OptionButton Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "GB"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "MB"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "KB"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "GE Curviture"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   9
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "GE Curviture"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   " Bit - Byte Converter "
      BeginProperty Font 
         Name            =   "GE Curviture"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox ("Enter Value")
Else
If Option1 = True Then
Text2.Text = Text1.Text * 1024 / 8
Label2.Caption = "    Bytes"
Label3.Caption = "  Kilo Bit"
ElseIf Option2 = True Then
Text2.Text = Text1.Text * 1024 / 8
Label2.Caption = " Kilo Bytes"
Label3.Caption = "  Mega Bit"
ElseIf Option3 = True Then
Text2.Text = Text1.Text * 1024 / 8
Label2.Caption = " Mega Bytes"
Label3.Caption = "   Giga Bit"
ElseIf Option4 = True Then
Text2.Text = Text1.Text * 1024 / 8
Label2.Caption = "  Giga Bytes"
Label3.Caption = "  Tera Bit"
End If

End If
End Sub

Private Sub Command2_Click()
If Option1 = True Then
Text2.Text = Text1.Text / 8
Label2.Caption = " Kilo Bytes"
Label3.Caption = "  Kilo Bit"
ElseIf Option2 = True Then
Text2.Text = Text1.Text / 8
Label2.Caption = " Mega Bytes"
Label3.Caption = "  Mega Bit"
ElseIf Option3 = True Then
Text2.Text = Text1.Text / 8
Label2.Caption = " Giga Bytes"
Label3.Caption = "  Giga Bit"
ElseIf Option4 = True Then
Text2.Text = Text1.Text / 8
Label2.Caption = " Tera Bytes"
Label3.Caption = "  Tera Bit"
End If
End Sub

Private Sub Command3_Click()
If Combo1.Text = KB Then
Text2.Text = Text1.Text * 1024 / 8
Label2.Caption = "    Bytes"
Label3.Caption = "  Kilo Bit"
End If
End Sub

Private Sub Form_Load()
Combo1.AddItem ("KB")
Combo1.AddItem ("MB")
Combo1.AddItem ("GB")
Combo1.AddItem ("TB")
End Sub
