Attribute VB_Name = "basComboControl"
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
       
Const CB_SHOWDROPDOWN = &H14F
Const CB_FINDSTRING = &H14C
Const CB_GETLBTEXTLEN = &H149
Const CB_GETDROPPEDWIDTH = &H15F
Const CB_SETDROPPEDWIDTH = &H160

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long

Type SIZE
    cx As Long
    cy As Long
End Type

Public Sub LockWindow(ByVal hwnd As Long)
Dim lRet As Long
    lRet = LockWindowUpdate(hwnd)
End Sub
Public Sub ReleaseWindow()
Dim lRet As Long
    lRet = LockWindowUpdate(0)
End Sub

Public Sub ComboDropdown(ByRef comboObj As ComboBox)
    Call SendMessage(comboObj.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Public Sub ComboRetract(ByRef comboObj As ComboBox)
    Call SendMessage(comboObj.hwnd, CB_SHOWDROPDOWN, 0, ByVal 0&)
End Sub

Public Function ComboAutoComplete(ByRef comboObj As ComboBox) As Boolean
Dim lngItemNum As Long
Dim lngSelectedLength As Long
Dim lngMatchLength As Long
Dim strCurrentText As String
Dim strSearchText As String
Dim sTypedText As String
Const CB_LOCKED = &H255

    With comboObj
        If .Text = Empty Then
            Exit Function
        End If
        Call LockWindow(.hwnd)
        If ((InStr(1, .Text, .Tag, vbTextCompare) <> 0 And Len(.Tag) = Len(.Text) - 1) Or (Left(.Text, 1) <> Left(.Tag, 1) And .Tag <> "")) And .Tag <> CStr(CB_LOCKED) Then
        
            strSearchText = .Text
            lngSelectedLength = Len(strSearchText)
        
            lngItemNum = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strSearchText)
            ComboAutoComplete = Not (lngItemNum = -1)
        
            If ComboAutoComplete Then
                lngMatchLength = Len(.List(lngItemNum)) - lngSelectedLength
                .Tag = CB_LOCKED
                sTypedText = strSearchText
                .Text = .Text & Right(.List(lngItemNum), lngMatchLength)
                .Tag = sTypedText
                .SelStart = lngSelectedLength
                .SelLength = lngMatchLength
            End If
        ElseIf .Tag <> CStr(CB_LOCKED) Then
            .Tag = .Text
        End If
        Call ReleaseWindow
    End With
End Function

Public Sub ComboDropWidth(ByRef comboObj As ComboBox)
Dim nCount As Long
Dim lNewDropDownWidth As Long
Dim lLongestString As Long

    On Error GoTo e_Trap
    For nCount = 0 To comboObj.ListCount - 1
        lNewDropDownWidth = comboObj.Parent.TextWidth(comboObj.List(nCount))
        If comboObj.Parent.ScaleMode = vbTwips Then
            lNewDropDownWidth = lNewDropDownWidth / Screen.TwipsPerPixelX  ' if twips change to pixels
        End If
        If lNewDropDownWidth > lLongestString Then
            lLongestString = lNewDropDownWidth
        End If
    Next nCount
    Call SendMessage(comboObj.hwnd, CB_SETDROPPEDWIDTH, lLongestString + 25, 0)
    Exit Sub
e_Trap:
    Exit Sub
End Sub

Public Sub ComboAddToHistory(ByRef comboObj As ComboBox, Optional ByVal bAllowDuplicates As Boolean = False, Optional ByVal nMaxEntries As Long = 100)
Dim nCount As Integer
Dim InList As Boolean

    '
    ' Combo_AddToHistory: adds current ComboBox's text to the dropdown list.
    '                     By default, this does not allow duplicates in the list.
    '                     Pass True to AllowDuplicates if needed.
    '

    With comboObj

        ' Don't add nulls
        If .Text = Empty Then Exit Sub

        If Not bAllowDuplicates Then
            For nCount = 0 To .ListCount - 1
                If .Text = .List(nCount) Then
                    ' Name is already in history. Don't add.
                    InList = True
                    Exit For
                End If
            Next nCount
        End If

        ' Don't maintain a list greater than 100 items.
        If nCount > nMaxEntries Then
            ' Remove 1st (oldest) entry...
            .RemoveItem 0
        End If

        If Not InList Then
            ' Add
            .AddItem .Text
            Call ComboDropWidth(comboObj)
        End If

    End With

End Sub

Public Sub ComboSaveHistory(ByRef comboObj As ComboBox)
Dim nCount As Integer
    
    '
    ' Combo_SaveHistory: saves current ComboBox's drop-down list to Registry
    '


    For nCount = 0 To comboObj.ListCount - 1
        Call SaveSetting(App.Title, "History", comboObj.Name & Format(nCount), comboObj.List(nCount))
    Next nCount
    ' Mark End
    On Local Error Resume Next
    DeleteSetting App.Title, "History", comboObj.Name & Format(nCount)

End Sub
Public Sub ComboLoadHistory(ByRef comboObj As ComboBox)
Dim Temp As String
Dim nCount As Integer
    
    '
    ' Combo_LoadHistory: loads current ComboBox's drop-down list with List
    '                    from Registry
    '

    comboObj.Clear
    Do
        On Error GoTo e_Trap
        Temp = GetSetting(App.Title, "History", comboObj.Name & Format(nCount), Default:=Chr$(255))
        If Not Temp = Chr$(255) Then
            ' Add item to ComboBox list
            comboObj.AddItem Temp
        Else
            Exit Do
        End If
        nCount = nCount + 1
    Loop
    Call ComboDropWidth(comboObj)
    Exit Sub
e_Trap:
    Exit Sub
End Sub


