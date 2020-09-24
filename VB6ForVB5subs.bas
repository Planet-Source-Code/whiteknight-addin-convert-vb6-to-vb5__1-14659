Attribute VB_Name = "basVB6ForVB5"
Option Explicit

Public Enum CompareMethod
    BinaryCompare
    TextCompare
End Enum
Public Function InStrRevVB5(ByVal StringCheck As String, ByVal StringMatch As String, Optional ByVal Start As Long = -1, Optional ByVal Compare As CompareMethod = BinaryCompare) As Long

'StringCheck:   The string to search.
'StringMatch:   The string to find.
'Start:         -1 = search entire string. Positive number = search only up to that position.
'Compare:       The compare method (binary or text)

'Returns:       The last position of StringMatch within StringCheck.

Dim lPos        As Long
Dim lSavePos    As Long
 
    If Start = -1 Then Start = Len(StringCheck)
    
    'Find the last instance of StringMatch within StringCheck.
    lPos = InStr(1, StringCheck, StringMatch, Compare)
    While lPos > 0 And lPos < Start
        lSavePos = lPos
        lPos = InStr(lPos + 1, StringCheck, StringMatch, Compare)
    Wend
    
    InStrRevVB5 = lSavePos
        
End Function

Public Function JoinVB5(SourceArray As Variant, Optional ByVal Delimiter As String = " ") As String

'SourceArray:   The array of strings to join.
'Delimiter:     The delimiter used in the join.

Dim lIdx    As Long
Dim lLower  As Long
Dim lUpper  As Long
Dim sRet    As String

    On Error GoTo LocalError
    'Return nothing if array has no lower or upper bounds.
    lLower = LBound(SourceArray)
    lUpper = UBound(SourceArray)
    
    'Concatinate the strings.
    For lIdx = lLower To lUpper
        sRet = sRet & SourceArray(lIdx) & Delimiter
    Next
    
    'Remove last delimiter.
    If Len(sRet) > 0 Then
        sRet = Left$(sRet, Len(sRet) - Len(Delimiter))
    End If
    
    'Return joined strings.
    JoinVB5 = sRet
    
NormalExit:
    Exit Function

LocalError:
    Resume NormalExit
    
End Function

Public Function SplitVB5(Expression As String, Optional ByVal Delimiter As String = "  ", Optional ByVal Limit As Long = -1, Optional ByVal Compare As CompareMethod = BinaryCompare) As Variant

'Expression:    The string to split.
'Delimiter:     The delimiter used for the split.
'Limit:         The max number of elements to return (-1 = all elements).
'Compare:       The compare method (binary or text).

'Returns:       A zero-based variant array of substrings or
'               entire expression as element(0) if no delimiter found.

Dim lPos1   As Long
Dim lPos2   As Long
Dim lIdx    As Long
Dim lCnt    As Long
Dim saTmp() As String

    'Initialize the variables
    lCnt = 0
    lPos1 = 1
    ReDim saTmp(99)
    
    'Search for the delimiter.
    lPos2 = InStr(1, Expression, Delimiter, Compare)
    While lPos2 > 0 And ((lCnt <= Limit) Or (Limit = -1))
        'Delimiter found, extract the substring between the delimiters.
        saTmp(lCnt) = Mid$(Expression, lPos1, lPos2 - lPos1)
        lCnt = lCnt + 1
        If (lCnt Mod 100) = 0 Then
            'Increase array size if needed.
            ReDim Preserve saTmp(UBound(saTmp) + 100)
        End If
        'Move to end of last delimiter found.
        lPos1 = lPos2 + Len(Delimiter)
        'Search for the next delimiter.
        lPos2 = InStr(lPos1, Expression, Delimiter, Compare)
    Wend
    
    If lPos1 < Len(Expression) Then
        'Extract last substring.
        saTmp(lCnt) = Mid$(Expression, lPos1)
        lCnt = lCnt + 1
    End If
    
    'Resize the array to correct size.
    If lCnt > 0 Then
        ReDim Preserve saTmp(lCnt - 1)
    Else
        ReDim saTmp(-1 To -1)
    End If
    
    'Return the array.
    SplitVB5 = saTmp
    
End Function

