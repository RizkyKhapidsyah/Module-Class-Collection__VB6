Attribute VB_Name = "PsychoColl"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit
Option Base 1

Private mvarCount As Long
Private mvarUKey As Boolean
Private ItemData() As String
Private ItemKey() As String

Public Sub XAdd(item As String, Optional key As String = Empty)

    If mvarUKey Then
        If Not KeyIsUnique(key) Then
            Err.Raise 457               'Key is not unique in collection
            Exit Sub
        End If
    End If
    mvarCount = mvarCount + 1
    ReDim Preserve ItemData(1 To mvarCount)
    ReDim Preserve ItemKey(1 To mvarCount)
    ItemData(mvarCount) = item
    ItemKey(mvarCount) = LCase(key)

End Sub

Public Property Get ReqUniqueKey() As Boolean

    ReqUniqueKey = mvarUKey

End Property

Public Property Let ReqUniqueKey(YesorNo As Boolean)

'If you have added data before with same keys they are not removed!
'The first one will be displayed when you will search by key
    mvarUKey = YesorNo

End Property

Public Property Get XCount() As Long

    XCount = mvarCount

End Property

Public Function XItem(ByVal XIndex As Long, Optional NewData As String = Empty) As String

    On Error Resume Next
    XItem = Empty
    If XIndex > mvarCount Then
        Err.Raise 9
        Exit Function
    End If
    If NewData <> Empty Then
        ItemData(XIndex) = NewData
    Else
        XItem = ItemData(XIndex)
    End If

End Function

Public Function XItembyKey(ByVal KeyData As String) As String
Dim count As Long
Dim elmax As Long, lcount As Long
Dim c1count As Long, c2count As Long
Dim XIndex As Long

    XIndex = -1
    count = 1
    KeyData = LCase(KeyData)
    On Error GoTo EmptyColl         'UBound causes error with an empty array
    lcount = UBound(ItemData)
    c1count = lcount / 2
    c2count = c1count
    Do
        If ItemKey(count) = KeyData Then
            XIndex = count
            Exit Do
        End If
        If ItemKey(lcount) = KeyData Then
            XIndex = lcount
            Exit Do
        End If
        If ItemKey(c1count) = KeyData Then
            XIndex = c1count
            Exit Do
        End If
        If ItemKey(c2count) = KeyData Then
            XIndex = c2count
            Exit Do
        End If
        count = count + 1
        lcount = lcount - 1
        c1count = c1count - 1
        c2count = c2count + 1
    Loop Until (lcount < c2count) And (count > c1count)
    If XIndex > 0 Then
        XItembyKey = ItemData(XIndex)
        Exit Function
    End If

EmptyColl:
    XItembyKey = "-1"

End Function

Public Function XKey(ByVal XIndex As Long, Optional NewKey As String = Empty) As String

    XKey = Empty
    If XIndex > mvarCount Then Exit Function
    If NewKey <> Empty Then
        ItemKey(XIndex) = NewKey
    Else
        XKey = ItemKey(XIndex)
    End If

End Function

Public Sub XRemove(ByVal XIndex As Long)
Dim counter As Long

    If XIndex > mvarCount Then
        Err.Raise 9             'Subscript out of range
        Exit Sub
    End If
    Err.Clear

On Error GoTo NotExists         'UBound causes error with an empty array
    While UBound(ItemData) > XIndex
        ItemData(XIndex) = ItemData(XIndex + 1)
        ItemKey(XIndex) = ItemKey(XIndex + 1)
        XIndex = XIndex + 1
    Wend
    ReDim Preserve ItemData(UBound(ItemData) - 1)
    ReDim Preserve ItemKey(UBound(ItemKey) - 1)
    mvarCount = mvarCount - 1
    Exit Sub

NotExists:
    Err.Raise vbObjectError + 9 'Subscript out of range

End Sub

Private Function KeyIsUnique(KeyData As String) As Boolean
Dim count As Long

    KeyIsUnique = False
    count = 1
On Error GoTo EmptyColl         'UBound causes error with an empty array
    Do
        If ItemKey(count) = KeyData Then Exit Function
        count = count + 1
    Loop Until UBound(ItemData) > count
    GoTo EmptyColl

EmptyColl:
    KeyIsUnique = True

End Function

Private Sub Class_Initialize()

    mvarCount = 0
    mvarUKey = False

End Sub
