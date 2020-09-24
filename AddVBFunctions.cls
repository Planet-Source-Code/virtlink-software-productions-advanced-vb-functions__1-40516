VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddVBFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private b1() As Byte
Private b2() As Byte
'By: Arthur Chaparyan
Public Function IsDuplicatedArray(ArrayCheck As Variant) As Boolean
    If Not IsArray(ArrayCheck) Then
        Err.Raise 13
    End If
    Dim a As Long, b As Long
    For a = LBound(ArrayCheck) To UBound(ArrayCheck)
        For b = a + 1 To UBound(ArrayCheck)
            If ArrayCheck(a) = ArrayCheck(b) Then IsDuplicatedArray = True: Exit Function
        Next
    Next
End Function
'By: Jorma Heikkinen
Public Sub SortArray(ArraySort As Variant)
    Dim n As Double
    Dim d As Double, k As Double, l As Double, rr As Double
    Dim i As Double, j As Double, Temp As Double, F As String, K1 As Double
    Dim TempString$
    Application.EnableCancelKey = xlErrorHandler
    n = UBound(ArraySort)
    ReDim H(1 To n) As Double
    Randomize Timer
    d = 10
    k = 1
    l = 1
    rr = n
Sort1:
    If rr - l + 1 > d Then
        GoTo Sort4
    Else
        GoSub Sort2
    End If
    GoTo Sort3
Sort4:
    i = l
    j = rr
    Temp = Rnd * (rr - l) + 0.5
    Temp = Int(Temp + l)
    F = ArraySort(Temp)
Sort5:
    If ArraySort(i) >= F Then
        GoTo Sort6
    Else
        i = i + 1
        GoTo Sort5
    End If
Sort6:
    If ArraySort(j) <= F Then
        GoTo Sort7
    Else
        j = j - 1
        GoTo Sort6
    End If
Sort7:
    If i > j Then
        GoTo Sort8
    Else
        TempString$ = ArraySort(i)
        ArraySort(i) = ArraySort(j)
        ArraySort(j) = TempString$
        i = i + 1
        j = j - 1
    End If
Sort8:
    If i <= j Then
        GoTo Sort5
    ElseIf j - 1 >= rr - i Then
        GoTo Sort9
    ElseIf i >= rr Then
        GoTo Sort10
    Else
        H(k) = i
        k = k + 1
        H(k) = rr
        k = k + 1
    End If
Sort10:
    rr = j
    GoTo Sort12
Sort9:
    If l >= j Then
        GoTo Sort11
    Else
        H(k) = l
        k = k + 1
        H(k) = j
        k = k + 1
    End If
Sort11:
    l = i
Sort12:
    If l < rr Then
        GoTo Sort4
    End If
Sort3:
    If k <= 1 Then
        GoTo Sort13
    Else
        k = k - 1
        rr = H(k)
        k = k - 1
        l = H(k)
        GoTo Sort1
    End If
Sort13:
    Exit Sub
Sort2:
    For i = l To rr - 1
        K1 = i
        For j = i + 1 To rr
            If ArraySort(j) < ArraySort(K1) Then
                K1 = j
            End If
        Next j
        If K1 = i Then
            TempString$ = ArraySort(i)
            ArraySort(i) = ArraySort(K1)
            ArraySort(K1) = TempString$
        End If
    Next i
    Return
End Sub
'By: CoderUX
Public Function GetDaysInMonth(MonthCheck As Integer, Optional YearCheck As Integer) As Double
    On Error GoTo Handler:
    If YearCheck = 0 Then
        GetDaysInMonth = DateDiff("D", "01/" & MonthCheck & "/" & Format(Now, "YYYY"), DateAdd("M", 1, "01/" & MonthCheck & "/" & Format(Now, "YYYY")))
    Else
        GetDaysInMonth = DateDiff("D", "01/" & MonthCheck & "/" & YearCheck, DateAdd("M", 1, "01/" & MonthCheck & "/" & YearCheck))
    End If
Handler:
    Exit Function
End Function
'By: John Buzzurro
Public Function InStrCount(StringCheck As String, StringMatch As String, Optional Start As Long = 0, Optional Compare As VbCompareMethod = vbBinaryCompare) As Integer
    Dim nPos As Integer
    Dim nCount As Integer
    nCount = 0
    nPos = InStr(Start, StringCheck, StringMatch, Compare)

    While nPos
        nCount = nCount + 1
        nPos = InStr(nPos + 1, StringCheck, StringMatch, Compare)
    Wend
    InStrCount = nCount
End Function
'By: John Buzzurro
Public Function TrimNull(StringCheck As String) As String
    Dim nPos As Integer
    nPos = InStr(StringCheck, Chr$(0))
    
    Select Case nPos
        Case Is > 1
            sTrimNulls = Left(StringCheck, nPos - 1)
        Case 1
            sTrimNulls = ""
        Case Else
            sTrimNulls = StringCheck
    End Select
End Function
'By: John Buzzurro
Public Function InsertString(StringSource As String, StringInsert As String, Optional Position As Long = 0) As String
    Dim sLeft As String, sRight As String
    Dim nLen As Integer
    
    nLen = Len(StringSource)
    
    If Position >= 0 And Position <= nLen Then
        If Position > 1 And nLen > 0 Then
            sLeft = Left(StringSource, Position - 1)
        Else
            sLeft = ""
        End If
        sRight = Right(StringSource, nLen - Position + 1)
        InsertString = sLeft & StringInsert & sRight
    Else
        InsertString = StringSource
    End If
End Function
'By: Arthur Chaparyan
Public Function IsDuplicated(ParamArray ArrayCheck()) As Boolean
    Dim a As Long, b As Long
    For a = LBound(ArrayCheck) To UBound(ArrayCheck)
        For b = a + 1 To UBound(ArrayCheck)
            If ArrayCheck(a) = ArrayCheck(b) Then IsDuplicated = True: Exit Function
        Next
    Next
End Function
'By: Brad Amsberry
Public Function AdvStrComp(StringCheck As String, StringMatch As String) As Double
    Dim l1 As Long
    Dim l2 As Long
    Dim l As Long
    Dim r As Double
    If UCase(StringCheck) = UCase(StringMatch) Then
        r = 1
    Else
        l1 = Len(StringCheck)
        l2 = Len(StringMatch)
        If l1 = 0 Or l2 = 0 Then
            r = 0
        Else
            ReDim b1(1 To l1): ReDim b2(1 To l2)
            For l = 1 To l1
                b1(l) = Asc(UCase(Mid(StringCheck, l, 1)))
            Next
            For l = 1 To l2
                b2(l) = Asc(UCase(Mid(StringMatch, l, 1)))
            Next
            r = SubSim(1, l1, 1, l2) / (l1 + l2) * 2
        End If
    End If
    AdvStrComp = r
    Erase b1
    Erase b2
End Function
'By: Brad Amsberry
Private Function SubSim(st1 As Long, end1 As Long, st2 As Long, end2 As Long) As Long
    Dim c1 As Long
    Dim c2 As Long
    Dim ns1 As Long
    Dim ns2 As Long
    Dim i As Long
    Dim max As Long
    If st1 > end1 Or st2 > end2 Or st1 <= 0 Or st2 <= 0 Then Exit Function
    For c1 = st1 To end1
        For c2 = st2 To end2
            i = 0
            Do Until b1(c1 + i) <> b2(c2 + i)
                i = i + 1
                If i > max Then
                    ns1 = c1
                    ns2 = c2
                    max = i
                End If
                If c1 + i > end1 Or c2 + i > end2 Then Exit Do
            Loop
        Next
    Next
    max = max + SubSim(ns1 + max, end1, ns2 + max, end2)
    max = max + SubSim(st1, ns1 - 1, st2, ns2 - 1)
    SubSim = max
End Function
'By: Daniël Pelsmaeker
Public Function GetPath(Filename As String) As String
    On Error Resume Next
    GetPath = Left(Filename, InStrRev(Filename, "\") - 1)
End Function
'By: Daniël Pelsmaeker
Public Function GetExt(strFilename As String) As String
    On Error Resume Next
    GetExt = LCase(Right(strFilename, Len(strFilename) - _
        InStrRev(strFilename, ".")))
End Function
'By: Daniël Pelsmaeker
Public Function GetFileTitle(Filename As String) As String
    On Error Resume Next
    
    Dim strTmp As String
    
    strTmp = Right(Filename, Len(Filename) - _
        InStrRev(Filename, "\"))
    strTmp = Left(strTmp, InStrRev(strTmp, ".") - 1)
    GetFileTitle = strTmp
End Function
'By: Daniël Pelsmaeker
Public Function GetFileName(Filename As String) As String
    On Error Resume Next
    
    Dim strTmp As String
    
    strTmp = Right(Filename, Len(Filename) - _
        InStrRev(Filename, "\"))
    GetFileName = strTmp
End Function
'By: Daniël Pelsmaeker
Public Function InArray(ArrayCheck, StringMatch As String, Optional Start As Long = 0, Optional Compare As VbCompareMethod = vbBinaryCompare) As Integer
    If Not IsArray(ArrayCheck) Then
        Err.Raise 13
    End If
    For q = Start To UBound(ArrayCheck)
        If StrComp(ArrayCheck(q), StringMatch, Compare) = 0 Then
            InArray = q
            Exit Function
        End If
    Next
    InArray = Null
End Function
'By: Daniël Pelsmaeker
Public Function InArrayRev(ArrayCheck, StringMatch As String, Optional Start As Long = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As Integer
    If Not IsArray(ArrayCheck) Then
        Err.Raise 13
    End If
    If Start = -1 Then Start = UBound(ArrayCheck)
    For q = Start To LBound(ArrayCheck) Step -1
        If StrComp(ArrayCheck(q), StringMatch, Compare) = 0 Then
            InArrayRev = q
            Exit Function
        End If
    Next
    InArrayRev = Null
End Function
'By: Daniël Pelsmaeker
Public Function OnReturn(Value, ParamArray RetValue())
    If IsNumeric(Value) Then
        If Value >= LBound(RetValue) And Value <= UBound(RetValue) Then
            OnReturn = RetValue(Value)
        Else
            OnReturn = Null
        End If
    Else
        Dim q As Integer
        For q = LBound(RetValue) To UBound(RetValue)
            If Value = RetValue(q) Then
                OnReturn = q
                Exit Function
            End If
        Next
        OnReturn = Null
    End If
End Function
'By: Daniël Pelsmaeker
Public Function OnReturnArray(Value, RetValue As Variant)
    If Not IsArray(RetValue) Then
        Err.Raise 13
    End If
    If IsNumeric(Value) Then
        If Value >= LBound(RetValue) And Value <= UBound(RetValue) Then
            OnReturnArray = RetValue(Value)
        Else
            OnReturnArray = Null
        End If
    Else
        Dim q As Integer
        For q = LBound(RetValue) To UBound(RetValue)
            If Value = RetValue(q) Then
                OnReturnArray = q
                Exit Function
            End If
        Next
        OnReturnArray = Null
    End If
End Function
'By: Brad Amsberry
Public Function InStrLike(StringCheck As String, StringLike As String, Optional Start As Long = 1, Optional Compare As VbCompareMethod = vbBinaryCompare) As Variant
    On Error GoTo err_InStrLike
    Dim intPos As Integer
    Dim intLength As Integer
    Dim strBuffer As String
    Dim blnFound As Boolean
    Dim varReturn As Variant
    
    If IsNull(StringCheck) Or IsNull(StringLike) Then
        varReturn = Null
        GoTo exit_InStrLike
    End If
    
    If Left(StringLike, 1) = "*" Then
        Err.Raise vbObjectError + 2600, "InStrLike", "Comparison mask cannot start With '*' since a start position cannot be determined."
        Exit Function
    End If
    
    For intPos = Start To Len(StringCheck) - Len(StringLike) + 1
        If InStr(1, StringLike, "*", vbTextCompare) Then
            For intLength = 1 To Len(StringCheck) - intPos + 1
                strBuffer = Mid(StringCheck, intPos, intLength)
                If strBuffer Like StringLike Then
                    blnFound = True
                    GoTo done
                End If
            Next intLength
        Else
            strBuffer = Mid(StringCheck, intPos, Len(StringLike))
            If strBuffer Like StringLike Then
                blnFound = True
                GoTo done
            End If
        End If
    Next intPos
done:
    If blnFound = False Then
        varReturn = 0
    Else
        varReturn = intPos
    End If
exit_InStrLike:
    InStrLike = varReturn
    Exit Function
err_InStrLike:
    varReturn = Null
    MsgBox Err.Description, vbCritical, "Error #" & Err.Number & " (InStrLike)"
    GoTo exit_InStrLike
End Function

