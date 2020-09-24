Attribute VB_Name = "Globals"
Option Explicit

Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOSIZE = &H1
Global Const FLAGS = SWP_SHOWWINDOW Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global folder As String
Global Const STRVB6R As String = "Retained=0"
Global Const STRVB6D As String = "DebugStartupOption=0"

Public Sub TopMost(f As Form, Optional ontop As Boolean = True)
    Dim success As Long
    Dim x2 As Long, y2 As Long
    x2 = f.Width / Screen.TwipsPerPixelX
    y2 = f.Height / Screen.TwipsPerPixelY
    If ontop Then
        success = SetWindowPos(f.hWnd, HWND_TOPMOST, f.Left, f.Top, x2, y2, FLAGS)
    Else
        success = SetWindowPos(f.hWnd, HWND_NOTOPMOST, f.Left, f.Top, x2, y2, 0)
    End If
End Sub
Public Function ReadFile(path As String) As String
    Dim TempStr As String
    Open path$ For Input As #1
    Do While Not EOF(1)
        Input #1, TempStr$
        ReadFile$ = ReadFile$ & vbCrLf & TempStr$
    Loop
    Close #1
End Function

Public Sub WriteFile(path As String, DataStr As String)
    Dim TempData As String, TempHoldStr As String
    Open path For Output As #1
    Print #1, DataStr$
    Close #1
End Sub

Public Function IsVB6(path As String) As Boolean
    Dim Data As String
    Data$ = ReadFile(path)
    If InStr(1, Data, STRVB6R$) <> 0 Or InStr(1, Data$, STRVB6D$) <> 0 Then
        IsVB6 = True
    Else
        IsVB6 = False
    End If

End Function

Public Function InvalidMsg(path As String) As String
    Dim Data As String
    Data$ = ReadFile(path)
    If InStr(1, Data, STRVB6R$) <> 0 And InStr(1, Data$, STRVB6D$) <> 0 Then
        InvalidMsg$ = "This Project contains invalid keys." & vbCrLf & STRVB6R$ & " is an invalid key." & vbCrLf & STRVB6D$ & " is an invalid key." & vbCrLf & vbCrLf & "Do you want to convert this file now?"
        Exit Function
    End If
    If InStr(1, Data, STRVB6R$) <> 0 Then
        InvalidMsg$ = "This Project contains invalid keys." & vbCrLf & STRVB6R$ & " is an invalid key." & vbCrLf & vbCrLf & "Do you want to convert this file now?"
        Exit Function
    End If
    If InStr(1, Data, STRVB6D$) <> 0 Then
        InvalidMsg$ = "This Project contains invalid keys." & vbCrLf & STRVB6D$ & " is an invalid key." & vbCrLf & vbCrLf & "Do you want to convert this file now?"
        Exit Function
    End If
End Function

Public Function Convert(path As String) As Boolean
    Dim Data As String, newData As String
    On Error GoTo fail
    If IsVB6(path) = True Then
        Data$ = ReadFile(path$)
        If LCase(Mid(Data$, 1, 3)) = vbCrLf & "t" Then Data$ = ReplaceString(Data$, vbCrLf & "Type", "Type")
        If InStr(1, Data$, vbCrLf & STRVB6R$) <> 0 Then Data$ = ReplaceString(Data$, vbCrLf & STRVB6R$, "")
        If InStr(1, Data$, vbCrLf & STRVB6D$) <> 0 Then Data$ = ReplaceString(Data$, vbCrLf & STRVB6D$, "")
        newData$ = Data$
        WriteFile path, newData$
    End If

    Convert = True
    frmmain.Caption = "Six 2 Five : File Converted with success."
    Exit Function
fail:
    Convert = False
    frmmain.Caption = "Six 2 Five : File Convert Failed."
End Function
Public Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String) As String
    Dim Spot As Long, NewSpot As Long, LeftString As String
    Dim RightString As String, NewString As String
    Spot& = InStr(LCase(MyString$), LCase(ToFind))
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
            Else
                RightString = ""
            End If
            NewString$ = LeftString$ & ReplaceWith$ & RightString$
            MyString$ = NewString$
        Else
            NewString$ = MyString$
        End If
        Spot& = NewSpot& + Len(ReplaceWith$)
        If Spot& > 0 Then
            NewSpot& = InStr(Spot&, LCase(MyString$), LCase(ToFind$))
        End If
    Loop Until NewSpot& < 1
    ReplaceString$ = NewString$
End Function

