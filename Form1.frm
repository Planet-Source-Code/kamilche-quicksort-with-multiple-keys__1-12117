VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3645
      TabIndex        =   3
      Text            =   "3000"
      Top             =   2400
      Width           =   945
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   2145
      Left            =   45
      TabIndex        =   2
      Top             =   60
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   3784
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sort Array"
      Height          =   495
      Left            =   2070
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fill Array"
      Height          =   495
      Left            =   420
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaration necessary for timing routine only.
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Type typeUDT
    'Keys - let's put them at the top, to make it clear what the
    'keys are. Note that you don't have to define the UDT this way -
    'the key values to be checked can actually be placed anywhere.
    Y As Integer                 'Primary key
    Z As Integer                 'Secondary key
    X As Integer                 'Third key.
    ItemNo As Long               'Misc junk to pad the UDT.
    OtherStuff(1 To 10) As Long  'More misc junk to pad the UDT.
    Name As String               'Yeah, it's junk. But typical of
                                 ' 'additional info' you'll likely have
                                 ' stored in your UDT, stuff you don't
                                 ' need to SORT by, but is a part
                                 ' of the data structure nonetheless.
End Type

'The array containing the information to be sorted.
Private MyArrayOfUDTs() As typeUDT

Private Sub Command1_Click()
    'Fill array with 3000 random entries.
    Dim TempItem As typeUDT, i As Long
    ReDim Preserve MyArrayOfUDTs(1 To Text2.Text)
    For i = 1 To Text2.Text
        With TempItem
            .X = Random(1, 600)
            .Y = Random(100, 455)
            .Z = Random(10, 20)
            .Name = "Line " & i & " - " & Format(.Y, "000") & Format(.Z, "000") & Format(.X, "000")
            MyArrayOfUDTs(i) = TempItem
        End With
    Next i
    Display "Array filled with " & UBound(MyArrayOfUDTs, 1) & " random elements."
End Sub

Private Sub Command2_Click()
    Dim s As String, StartTime As Long
    StartTime = timeGetTime
    QuickSort MyArrayOfUDTs, LBound(MyArrayOfUDTs, 1), UBound(MyArrayOfUDTs, 1)
    s = "Sort completed in " & Format((timeGetTime - StartTime) / 1000, "#.###") & " seconds."
    Display s
'    Dim i As Long
'    For i = 1 To Text2.Text
'        With MyArrayOfUDTs(i)
'            Debug.Print Format(.Y, "000") & vbTab & Format(.Z, "000") & vbTab & Format(.X, "000") & vbTab & .ItemNo & vbTab & .Name
'        End With
'    Next i
End Sub

Private Sub Display(ByVal s As String)
    Text1.SelStart = Len(Text1.Text)
    Text1.SelText = s & vbCrLf
End Sub

Private Function Random(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    Random = Int((UpperBound - LowerBound + 1) * Rnd + LowerBound)
End Function

Private Function QuickSort(Item() As typeUDT, LowerBound As Long, UpperBound As Long)
'This routine actually performs the sort.
'It calls the method 'LessThan' to actually perform the UDT comparison.
'Modify 'LessThan' to sort on different or additional keys.

    Dim MidpointItem As typeUDT
    Dim TempItem As typeUDT
    Dim CurLow As Long
    Dim CurHigh As Long
    Dim CurMidpoint As Long
    
    CurLow = LowerBound
    CurHigh = UpperBound
    
    If UpperBound <= LowerBound Then Exit Function
    CurMidpoint = (LowerBound + UpperBound) \ 2
    
    MidpointItem = Item(CurMidpoint)
    Do While (CurLow <= CurHigh)

        Do While LessThan(Item(CurLow), MidpointItem)
            CurLow = CurLow + 1
            If CurLow = UpperBound Then Exit Do
        Loop
        
        Do While LessThan(MidpointItem, Item(CurHigh))
            CurHigh = CurHigh - 1
            If CurHigh = LowerBound Then Exit Do
        Loop

        If (CurLow <= CurHigh) Then
            TempItem = Item(CurLow)
            Item(CurLow) = Item(CurHigh)
            Item(CurHigh) = TempItem
            CurLow = CurLow + 1
            CurHigh = CurHigh - 1
        End If
        
    Loop

    If LowerBound < CurHigh Then
        QuickSort Item(), LowerBound, CurHigh
    End If

    If CurLow < UpperBound Then
        QuickSort Item(), CurLow, UpperBound
    End If
    
End Function

Private Function LessThan(Item1 As typeUDT, Item2 As typeUDT) As Boolean
    'This is the routine to modify, to change the sort keys.
    'The goal of this routine, is to return 'true' if item1 < item2.
    'For a single key sort, checking the value of two variables is enough.
    'For a multi-key sort, it has to be more complex -
    ' if key1 of item1 is < key1 of item2, well, ok then.
    ' but! if it's equal, you have to progress to the NEXT key.
    ' So  if key2 of item1 < key2 of item2, ok then.
    ' Otherwise! Check key 3.
    'Proceed in this manner, for MyArrayOfUDTs the keys you need to check.
    'In this example, the sort keys in ascending order are:
    ' Y
    ' Z
    ' X
    
    LessThan = False
    If Item1.Y < Item2.Y Then
        LessThan = True
    ElseIf Item1.Y = Item2.Y Then
        If Item1.Z < Item2.Z Then
            LessThan = True
        ElseIf Item1.Z = Item2.Z Then
            If Item1.X < Item2.X Then
                LessThan = True
            End If
        End If
    End If
End Function

