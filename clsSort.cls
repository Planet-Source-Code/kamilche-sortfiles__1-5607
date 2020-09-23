VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Friend Function DirectoryList(ByVal PathName As String) As String()
    DirectoryList = List(PathName, vbDirectory)
End Function

Friend Function FileList(ByVal PathName As String) As String()
    FileList = List(PathName, vbNormal)
End Function

Private Function List(ByVal PathName As String, ByVal ObjType As VbFileAttribute) As String()
    Dim DirName As String, SortArray() As String, Max As Long
    'If it's not a valid directory bail out
    If Dir(PathName, vbDirectory) = "" Then
        Exit Function
    End If
    'Fill the sortarray with the subdirectories
    DirName = Dir(PathName & "\*.*", vbDirectory)
    Do
        If DirName = "" Then
            Exit Do
        End If
        If DirName = "." Or DirName = ".." Then
            'skip it
        ElseIf ObjType = vbDirectory And GetAttr(PathName & "\" & DirName) <> vbDirectory Then
            'skip it
        ElseIf ObjType = vbNormal And GetAttr(PathName & "\" & DirName) = vbDirectory Then
            'skip it
        Else
            Max = Max + 1
            ReDim Preserve SortArray(1 To Max)
            SortArray(Max) = PathName & "\" & DirName
        End If
        DirName = Dir()
    Loop
    'Sort the resulting list
    QuickSort SortArray, 1, Max
    List = SortArray
End Function

Private Sub QuickSort(SortArray() As String, ByVal Low As Long, ByVal High As Long)
    Dim i As Long, J As Long, RandIndex As Long, Partition As String
    If Low < High Then
        ' Only two elements in this subdivision; swap them if they are
        ' out of order, then end recursive calls:
        If High - Low = 1 Then
            If UCase(SortArray(Low)) > UCase(SortArray(High)) Then
                SWAP SortArray(Low), SortArray(High)
            End If
        Else
            ' Pick a pivot element at random, then move it to the end:
            RandIndex = Rnd() * (High - Low) + Low ' RandInt%(Low, High)
            SWAP SortArray(High), SortArray(RandIndex)
            Partition = UCase(SortArray(High))
            Do
                ' Move in from both sides towards the pivot element:
                i = Low: J = High
                Do While (i < J) And (UCase(SortArray(i)) <= Partition)
                    i = i + 1
                Loop
                Do While (J > i) And (UCase(SortArray(J)) >= Partition)
                    J = J - 1
                Loop
                ' If we haven't reached the pivot element it means that 2
                ' elements on either side are out of order, so swap them:
                If i < J Then
                    SWAP SortArray(i), SortArray(J)
                End If
            Loop While i < J
            ' Move the pivot element to its proper place in the array:
            SWAP SortArray(i), SortArray(High)
            ' Recursively call the QuickSort procedure (pass the
            ' smaller subdivision first to use less stack space):
            If (i - Low) < (High - i) Then
                QuickSort SortArray, Low, i - 1
                QuickSort SortArray, i + 1, High
            Else
                QuickSort SortArray, i + 1, High
                QuickSort SortArray, Low, i - 1
            End If
        End If
    End If
End Sub

Private Sub SWAP(first As String, second As String)
    Dim temp As String
    temp = first
    first = second
    second = temp
End Sub
