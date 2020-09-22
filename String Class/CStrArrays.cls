VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStrArrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'
'    CJDigital Terms of Usage:
'
'    By using this code, you agree to the following terms...
'
'    1) You may use this code in your own programs
'    (and may compile it into a program and distribute it
'    in compiled format for langauges that allow it) freely and with no charge.
'
'    2) If you do use this code for profit, an mention of the Author and
'    Company name would be more than appreciated.
'
'    3) You MAY NOT redistribute this code without written
'    permission from the original author. Failure to do so is a violation of copyright laws.
'
'    4) In Otherwords, Don't Screw ME! It isn't necessary,
'    Im just looking for a LIL Recognition, Wouldn't you?
'
'    Copyright:        © 2002 CJDigital.
'    AUTHOR:           Chris Hoffman
'    AUTHORS EMAIL:    Tech@CJDigital.com   OR   stanleyhoffman@knology.net
'    AUTHORS WEBSITE:  http://www.CJDigital.com
'
'        Project Type: Class Obj
'
'        Dependents  : NONE
'
'        References  : NONE
'
'        Thanx Too   : "David Bue Pedersen" for his submission at PSC:
'               http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=36945&lngWId=1
'
'               See sub "AppendWithBuffer"
'
'        FUTURE PLANS: More and More optimization   :O)
'
'        Please Note: This class is a work in progress, as i use it in nearly all
'                     of my projects, you can guess its very important to me that it
'                     work flawlessly. If you CHANGE anything for the better please by
'                     all means let me know!  :O)


'
'   Class Constants
Const INITIAL_SIZE = 2
Const SIZE_64K = 65536
'
'   Class variables
Private mElements   As Integer
Private mStrArray() As String
Private mSizes()    As Double

Private Sub Class_Initialize( _
)
'
'   Allocate initial amount of memory to the array
'
    ReDim Preserve mStrArray(INITIAL_SIZE)
    ReDim Preserve mSizes(INITIAL_SIZE)
End Sub

Private Sub Class_Terminate( _
)
'
'   Deallocate the entire array
'
    Erase mStrArray
    Erase mSizes
    mElements = 0
End Sub

Public Property Get Size( _
) As Long
'
'   Return the current array size
'
    Size = UBound(mStrArray)
End Property

Public Property Get Count( _
) As Long
'
'   Return current amount of elements in the array
'
    Count = mElements
End Property

Public Property Get Item( _
    Index As Integer _
) As String
'
'   Return the value found at the specified location
'
Dim sTmp As String
Dim lPos As Long

    If Index < Size() Then
        lPos = InStr(1, mStrArray(Index), Chr(0))
        If (lPos > 0) Then
            Item = Left$(mStrArray(Index), lPos - 1)
        Else
            Item = mStrArray(Index)
        End If
    End If
End Property

Public Property Let Item( _
    Index As Integer, _
    newValue As String _
)
'
'   Assign a new value to the specified item
'
    If Index < Size() Then
        mStrArray(Index) = newValue
        mSizes(Index) = Len(newValue)
        If Index >= mElements Then mElements = Index + 1
    End If
End Property

Public Sub Add( _
    newValue As String, _
    Optional ByRef NewIndex As Integer _
)
'
'   Add a new item to the array Returning an INDEX value of the element
'
    '
    '   Check the size of the array
    Resize
    '
    '   Add the value at the next index
    mStrArray(mElements) = newValue
    mSizes(mElements) = Len(newValue)
    NewIndex = mElements
    mElements = mElements + 1
End Sub

Public Function AddNew( _
    newValue As String _
) As Integer
'
'   Add a new item to the array only if it doesn't already exist
'   and return an INDEX value of the NEW element.
'
    On Error GoTo ER
    AddNew = -1
    If newValue = vbNullString Then Exit Function
    If IsInArray(newValue) Then
        Exit Function
        Else
        Call Add(newValue, AddNew)
    End If
    Exit Function
ER:
End Function

Public Function RemoveString( _
    KillString As String _
) As Boolean
    
    Dim iItem As Integer
    iItem = Find(KillString)
    If Not iItem >= 0 Then
        Exit Function
        Else
        Let Item(iItem) = vbNullString
        Call Resize
        RemoveString = True
    End If
End Function

Public Function IsInArray( _
    FindValue As String _
) As Boolean

    IsInArray = InStr(1, vbNullChar & Join(mStrArray, vbNullChar) & _
        vbNullChar, vbNullChar & FindValue & vbNullChar) > 0
End Function

Private Sub Resize( _
)
'
'   Resize the array to fit more items
'
    Dim arraySize As Long
    arraySize = Size()
    If Count() = arraySize Then
        ReDim Preserve mStrArray(arraySize + INITIAL_SIZE)
        ReDim Preserve mSizes(arraySize + INITIAL_SIZE)
    End If
End Sub

Public Sub Clear( _
)
'
'   Erase the array
'
    Erase mStrArray
    Erase mSizes
    mElements = 0
    ReDim Preserve mStrArray(INITIAL_SIZE)
    ReDim Preserve mSizes(INITIAL_SIZE)
End Sub

Public Function Find( _
    strItem As String _
) As Integer
'
'   Find the specified item in the list;
'   return its index if found, -1 otherwise
'
    '
    '   Find the specified item
    Dim i As Integer
    If (IsInArray(strItem)) Then
        For i = 0 To Count - 1
            If LCase(Item(i)) = LCase(strItem) Then
                Find = i
                Exit Function
            End If
        Next
    End If
    Find = -1
End Function

Public Sub AppendToItem( _
    Index As Integer, _
    ByVal AppendString As String, _
    Optional bForceNoBuffer As Boolean _
)
    If (Len(mStrArray(Index)) >= SIZE_64K) And Not bForceNoBuffer Then
        Call AppendWithBuffer(Index, AppendString)
    Else
        Call AppendNoBuffer(Index, AppendString)
    End If
End Sub

Private Sub AppendNoBuffer( _
    Index As Integer, _
    ByVal sAppend As String _
)
    Item(Index) = mStrArray(Index) & sAppend
End Sub

Private Sub AppendWithBuffer( _
    Index As Integer, _
    ByVal sAppend As String _
)
    Dim strTemp As String
    Dim lngLoop As Long
    '
    '   Empty strings will cause a fatal error if not eliminated
    If Not Trim$(sAppend) = "" Then
        '
        '   Does the BUFFER of the Current Item need to be INCREASED?
        If mSizes(Index) + Len(sAppend) > Len(mStrArray(Index)) Then
            '
            '   STORE mStrArray(index) current value
            strTemp = mStrArray(Index)
            '
            '   Increase memory storage bytes
            Do
                '
                '   Increment for Multiplier
                lngLoop = lngLoop + 1
                If ((Len(mStrArray(Index)) + (SIZE_64K * lngLoop)) >= (mSizes(Index) + Len(sAppend))) Then
                    Exit Do
                End If
            Loop
            '
            '   Resize buffer
            mStrArray(Index) = String$(Len(mStrArray(Index)) + (SIZE_64K * lngLoop), Chr(0))
            '
            '   RESTORE mStrArray(index)'s Value
            Mid$(mStrArray(Index), 1, mSizes(Index)) = strTemp
        End If
        '
        '   Insert the NEW data BEFORE the NULLS
        Mid$(mStrArray(Index), mSizes(Index) + 1, Len(sAppend)) = sAppend
        '
        '   Reset the Buffers length for current item.
        mSizes(Index) = mSizes(Index) + Len(sAppend)
    End If
End Sub
