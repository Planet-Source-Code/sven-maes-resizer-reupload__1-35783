VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "resizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private frmSize As Object
Private Type objSize
    percTop As Single
    percLeft As Single
    percHeight As Single
    percWidth As Single
    name As String
    fontsize As Single
End Type
Private collSize() As objSize
Private sngOrigHeight As Single
Private sngOrigWidth As Single
Dim frmset As Boolean
Dim doFonts As Boolean

    '*******************************************************************************
    ' Resizer (Class)
    '
    ' Author : Sven Maes
    '          mail.me@vt4.net
    '
    ' DESCRIPTION:
    ' This class resizes all controls on the form
    '
    ' usage :
    ' in the form, add a reference to clsResizer
    ' in the form, declare a new resizerclass
    ' in form_load : set resizer.frmtosize = me
    ' in form_resize : resizer.adjustSize
    '  --> 1 resizer variable for each form !
    '*******************************************************************************
Public Sub ResetForm()
    setInitialFormPos
End Sub
Public Property Set frmToSize(frm As Object)
    Set frmSize = frm
    setInitialFormPos
End Property
Public Sub AddItemNotToResize(Itemname As String)
    Dim i As Long

    If Not frmset Then Exit Sub
    
    For i = 0 To UBound(collSize)
        If LCase(collSize(i).name) = LCase(Itemname) Then
            collSize(i).name = ""
            Exit For
        End If
    Next i


End Sub
Public Property Let ResizeFonts(rsf As Boolean)
    doFonts = rsf
End Property


Private Sub setInitialFormPos()
    'remember the relative size and position of the controls on the form
    Dim objContr As Control
    Dim i As Integer
    
    If frmSize.Controls.Count = 0 Then Exit Sub

    sngOrigHeight = frmSize.ScaleHeight
    sngOrigWidth = frmSize.ScaleWidth
    ReDim collSize(0 To frmSize.Controls.Count)
    i = 0
    For Each objContr In frmSize.Controls
        If TTS(objContr) Then 'try to size
            collSize(i).name = objContr.name
            collSize(i).percHeight = objContr.Height / sngOrigHeight
            collSize(i).percLeft = objContr.Left / sngOrigWidth
            collSize(i).percTop = objContr.Top / sngOrigHeight
            collSize(i).percWidth = objContr.Width / sngOrigWidth
            If doFonts And TTF(objContr) Then
                collSize(i).fontsize = objContr.Font.Size / sngOrigHeight
            Else
                collSize(i).fontsize = -1
            End If
            i = i + 1
        Else
            'move the object
            
            'special objects
            Select Case LCase(TypeName(objContr))
                Case "timer"
                    'not possible to resize or to move
                    DoEvents
                Case "line"
                    collSize(i).name = objContr.name
                    collSize(i).percLeft = objContr.X1 / sngOrigWidth
                    collSize(i).percWidth = objContr.X2 / sngOrigWidth
                    collSize(i).percTop = objContr.Y1 / sngOrigHeight
                    collSize(i).percHeight = objContr.Y2 / sngOrigHeight
                    i = i + 1
                    If doFonts And TTF(objContr) Then
                        collSize(i).fontsize = objContr.Font.Size / sngOrigHeight
                    Else
                        collSize(i).fontsize = -1
                    End If
                Case Else
                    collSize(i).name = objContr.name
                    collSize(i).percLeft = objContr.Left / sngOrigWidth
                    collSize(i).percTop = objContr.Top / sngOrigHeight
                    collSize(i).percWidth = -1 'negative number -> move only
                    collSize(i).percHeight = -1
                    i = i + 1
                    If doFonts And TTF(objContr) Then
                        collSize(i).fontsize = objContr.Font.Size / sngOrigHeight
                    Else
                        collSize(i).fontsize = -1
                    End If
            End Select
        End If
    Next
    frmset = True
End Sub
Private Function TTS(obj As Control) As Boolean
    'this function tests if the control is sizeable
    On Error GoTo errhan
    
    obj.Width = obj.Width
    obj.Height = obj.Height
    
    TTS = True
    Exit Function
errhan:
    TTS = False
End Function
Private Function TTF(obj As Control) As Boolean
    'Try To size Font
    On Error GoTo errhan

    obj.Font.Size = obj.Font.Size

    TTF = True
    Exit Function
errhan:
    TTF = False
End Function
Public Sub AdjustSize()

    If Not frmset Then Exit Sub
    
    Dim objContr As Control
    Dim j As Integer
    
    If frmSize.Controls.Count = 0 Then Exit Sub

    sngOrigHeight = frmSize.ScaleHeight
    sngOrigWidth = frmSize.ScaleWidth
    j = 0
    For Each objContr In frmSize.Controls
        For j = 0 To UBound(collSize)
            If collSize(j).name = objContr.name Then
                If LCase(TypeName(objContr)) = "line" Then
                    objContr.X1 = collSize(j).percLeft * sngOrigWidth
                    objContr.X2 = collSize(j).percWidth * sngOrigWidth
                    objContr.Y1 = collSize(j).percTop * sngOrigHeight
                    objContr.Y2 = collSize(j).percHeight * sngOrigHeight
                Else
                    'move should work
                    objContr.Left = collSize(j).percLeft * sngOrigWidth
                    objContr.Top = collSize(j).percTop * sngOrigHeight
                    If collSize(j).percHeight > 0 Then 'size only when possible
                        objContr.Width = collSize(j).percWidth * sngOrigWidth
                        objContr.Height = collSize(j).percHeight * sngOrigHeight
                    End If
                End If
                
                If doFonts And collSize(j).fontsize > 0 Then
                    objContr.Font.Size = IIf(Int(collSize(j).fontsize * sngOrigHeight) = 0, objContr.Font.Size, Int(collSize(j).fontsize * sngOrigHeight))
                End If
            End If
        Next j
    Next
     
End Sub

Private Sub Class_Initialize()
    doFonts = False
    frmset = False
End Sub
