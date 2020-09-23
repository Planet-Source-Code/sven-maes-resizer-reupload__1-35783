VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3120
      TabIndex        =   13
      Top             =   3960
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   2760
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   195
      Left            =   1680
      TabIndex        =   9
      Top             =   3120
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      Height          =   540
      Left            =   1560
      TabIndex        =   8
      Top             =   4080
      Width           =   855
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   615
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   3600
      Width           =   1695
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   1080
      TabIndex        =   5
      Top             =   3120
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Left            =   1680
      Top             =   480
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   3240
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2475
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   600
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   840
      Y1              =   4200
      Y2              =   4680
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   3840
      TabIndex        =   11
      Top             =   2880
      Width           =   735
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   3000
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   2160
      Picture         =   "test.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r As New resizer
    
Private Sub Form_Load()
    'required
    Set r.frmToSize = Me

End Sub

Private Sub Form_Resize()
    'required
    r.AdjustSize
End Sub
