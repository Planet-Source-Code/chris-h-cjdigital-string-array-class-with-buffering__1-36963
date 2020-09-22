VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "String Arrays CLASS with Buffering "
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Also in array class but FORCING no buffer."
      Height          =   735
      Left            =   2520
      TabIndex        =   1
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "In array class and uses string buffering."
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":0000
      Height          =   975
      Index           =   1
      Left            =   345
      TabIndex        =   4
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":010E
      Height          =   615
      Index           =   0
      Left            =   345
      TabIndex        =   3
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "This demonstration is based on 2 things:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   3450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objStr As CStrArrays
Dim objStr2 As CStrArrays

Private Sub Command1_Click()
    Dim lTimer      As Double
    Dim i           As Integer
    Dim iIndex      As Integer
    Dim bRemoved    As Boolean
    Dim sTmp        As String
    '
    '   Start timing
    lTimer = Timer
    '
    '   Set methods and append some stuff
    With objStr
        iIndex = .AddNew("[This is a test String with 100 APPENDS]." & vbCrLf)
        '
        '   Loop and add a few diffrent items
        For i = 1 To 1000
            .AppendToItem (0), " This is a really long test to vreate as many characters as possible so i can try and see what may or may not happen to my string class as it processses an striing greater than the possible 64000 characters#" & i & vbCrLf
        Next i
    End With
    '
    '   Calc time taken
    lTimer = Format(Timer - lTimer, "00.00")
    sTmp = Format(CStr(Len(objStr.Item(0))), "000,000")
    MsgBox "Time taken to add string(s): " & lTimer & "(s)" & vbCrLf & _
    "The length of the ENTIRE element (String): " & sTmp, , "STRING BUFFERING"

End Sub

Private Sub Command2_Click()
    Dim lTimer      As Double
    Dim i           As Integer
    Dim iIndex      As Integer
    Dim bRemoved    As Boolean
    Dim sTmp        As String
    '
    '   show warning
    MsgBox "This may or may not take a while depending on your system," & vbCrLf & _
    "for example purposes BOTH iterations are 1000.", vbCritical
    '
    '   Start timing
    lTimer = Timer
    '
    '   Set methods and append some stuff
    With objStr2
        iIndex = .AddNew("[This is a test String with 100 APPENDS]." & vbCrLf)
        '
        '   Loop and add a few diffrent items
        For i = 1 To 1000
            .AppendToItem (0), " This is a really long test to vreate as many characters as possible so i can try and see what may or may not happen to my string class as it processses an striing greater than the possible 64000 characters#" & i & vbCrLf, True
        Next i
    End With
    '
    '   Calc time taken
    lTimer = Format(Timer - lTimer, "00.00")
    sTmp = Format(CStr(Len(objStr2.Item(0))), "000,000")
    MsgBox "Time taken to add string(s): " & lTimer & "(s)" & vbCrLf & _
    "The length of the ENTIRE element (String): " & sTmp, , "NO STRING BUFFERING"
End Sub

Private Sub Form_Load()
    '
    '   Create cls Object
    Set objStr = New CStrArrays
    Set objStr2 = New CStrArrays
    '
    '   Show info
    MsgBox "The STRINGS are NOT being emptied when after clicking" & vbCrLf & _
    "so be carefull as the more you click an button the LARGER the string!"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objStr2 = Nothing
    Set objStr = Nothing
End Sub




