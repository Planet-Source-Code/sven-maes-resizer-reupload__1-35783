VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Width           =   2415
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   1200
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2535
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   255
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   5880
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   495
      Left            =   6720
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   495
      Left            =   6720
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   4680
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.OLE OLE1 
      Height          =   975
      Left            =   840
      TabIndex        =   13
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   7920
      Y1              =   480
      Y2              =   2880
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   1680
      Y1              =   960
      Y2              =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   4680
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
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
    
    'let the fonts be resized to (before the frmtosize)
    r.ResizeFonts = True
    
    Set r.frmToSize = Me
    'optionally add objects not to resize
    'when this line is removed, command1 will also be resized
    r.AddItemNotToResize "command1"
End Sub

Private Sub Form_Resize()
    r.AdjustSize
End Sub
