VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   2985
      Left            =   5445
      TabIndex        =   3
      Top             =   1065
      Width           =   6390
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   90
      TabIndex        =   2
      Top             =   1050
      Width           =   5250
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fill List"
      Height          =   495
      Left            =   90
      TabIndex        =   1
      Top             =   465
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   45
      TabIndex        =   0
      Text            =   "c:\Incarnation\Region Builder\Region Graphics"
      Top             =   75
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Given a pathname, the class in this program returns a sorted list of
' directories or filenames in a string array, without the use of the
' drivelistbox, dirlistbox, or filelistbox controls.

'This program illustrates several important programming concepts, including
' quicksort, passing arrays to and from procedures, and handling
' undimensioned arrays. In addition, it's very easy to use.

Private Sub Command1_Click()
    Dim i As Long
    Dim Sorted As New clsSort
    Dim PathName As String, DirArray() As String
    On Error GoTo Err_Init
    'retrieve the enclosing folder name
    PathName = Form1.Text1.Text
    If Dir(PathName, vbDirectory) = "" Then
        MsgBox "Directory not found!"
        Exit Sub
    End If
    'retrieve all sub folders 1 level deep
    DirArray = Sorted.DirectoryList(PathName)
    'fill the listbox with the folder names
    List1.Clear
    'Next line might cause error - if it does, handle in error handler.
    For i = 1 To UBound(DirArray, 1)
        List1.AddItem DirArray(i)
    Next i
    Exit Sub
    
Err_Init:
    If Err.Number = 9 Then
        MsgBox "No subdirectories found!"
    End If
End Sub

Private Sub List1_Click()
    Dim i As Long
    Dim Sorted As New clsSort
    Dim PathName As String, FileArray() As String
    On Error GoTo Err_Init
    
    'Make sure they're sitting on a valid list item
    If List1.ListIndex < 0 Then
        Exit Sub
    End If
    'retrieve the enclosing folder name
    PathName = List1.List(List1.ListIndex)
    'retrieve all sub folders 1 level deep
    FileArray = Sorted.FileList(PathName)
    'fill the listbox with the folder names
    List2.Clear
    'Next line might cause error - if it does, handle in error handler.
    For i = 1 To UBound(FileArray, 1)
        List2.AddItem FileArray(i)
    Next i
    Exit Sub
    
Err_Init:
    If Err.Number = 9 Then
        MsgBox "No files found!"
    End If
End Sub
