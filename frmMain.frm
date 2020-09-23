VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstFiles 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   1200
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   720
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This demonstrates the basic in Dragging a file from the explorer and
'dropping it into a VB program.
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'(c) Thomas Jung 2000

Private Sub Form_Resize()
    'Always view lstFiles in the whole form
    lstFiles.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub lstFiles_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Count number of files
    Dim numFiles As Integer
    numFiles = Data.Files.Count
    
    'Add all dropped files into the list
    Dim i As Integer
    For i = 1 To numFiles
        'File or directory?
        If (GetAttr(Data.Files(i)) And vbDirectory) = vbDirectory Then
            lstFiles.AddItem "Directory: " & Data.Files(i)
        Else
            lstFiles.AddItem "File.....: " & Data.Files(i)
        End If
    Next i

End Sub

Private Sub lstFiles_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    'Only let files get dropped into the form
    If Data.GetFormat(vbCFFiles) Then
        Effect = vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
    End If
End Sub
