Attribute VB_Name = "CAFT"
'**********************************************************************************
'* Setup Menu
'**********************************************************************************
Sub SetupMenu()
    Dim menuGroup As CommandBarControl
    
    'Create Custom Menu Group
    Set menuGroup = Application.CommandBars("Cell").Controls.Add(Type:=msoControlPopup)
    menuGroup.Caption = "Copy as ..."
    
    
    Dim buttons(5) As Variant
    buttons(0) = Array("Markdown", "CopyByMarkDown")
    buttons(1) = Array("Trac", "CopyByTrac")
    buttons(2) = Array("PukiWiki", "CopyByPukiWiki")
    buttons(3) = Array("XPlanner-plus", "CopyByXPlanner")
    'buttons(4) = Array("Reset Menu", "DelAllMenu")
    
    'Create Button
    Dim btn As Variant
    For Each btn In buttons
        If Not IsEmpty(btn) Then
            With menuGroup.Controls.Add(Type:=msoControlButton)
                .Caption = btn(0)
                .OnAction = btn(1)
            End With
        End If
    Next btn
End Sub

Sub DeleteMenu()
    Application.CommandBars("Cell").Reset
End Sub


'**********************************************************************************
'* Copy(MarkDown)
'**********************************************************************************
Private Sub CopyByMarkDown()
    
    '選択範囲の取得し、フォーマット変換
    If Selection.Count <> 0 Then
        '各列の最大文字列長を計算
        Dim lenMap As Variant
        lenMap = CalcLength()
        
        'TODO 各列の最大文字列長からきれいに整形したい
        Dim result As String
        result = ""
        
        Dim rStart As Integer
        Dim rEnd As Integer
        Dim cStart As Integer
        Dim cEnd As Integer
        
        rStart = Selection(1).Row
        rEnd = Selection(Selection.Count).Row
        cStart = Selection(1).Column
        cEnd = Selection(Selection.Count).Column
    
        '列数からヘッダの作成
        result = result & "|"
        For c = cStart To cEnd
            result = result & PaddingText(" " & Cells(rStart, c).Value, " ", " |", lenMap(c - cStart) + 1)
        Next c
        result = result & vbCrLf
        
        
        '列数からヘッダしたを作成
        result = result & "|"
        For c = cStart To cEnd
            result = result & PaddingText("-", "-", ":|", lenMap(c - cStart) + 1)
        Next c
        result = result & vbCrLf
    
        
        '選択範囲からボディを作成
        For r = rStart + 1 To rEnd
            result = result & "|"
            For c = cStart To cEnd
                result = result & PaddingText(" " & Cells(r, c).Value, " ", " |", lenMap(c - cStart) + 1)
            Next c
            result = result & vbCrLf
        Next r
    
        'クリップボードに保存
        SaveToClipboard (result)

    End If
End Sub


'**********************************************************************************
'* Copy(Trac)
'**********************************************************************************
Private Sub CopyByTrac()
    
    '選択範囲の取得し、フォーマット変換
    If Selection.Count <> 0 Then
        '各列の最大文字列長を計算
        Dim lenMap As Variant
        lenMap = CalcLength()
        
        'TODO 各列の最大文字列長からきれいに整形したい
        Dim result As String
        result = ""
        
        Dim rStart As Integer
        Dim rEnd As Integer
        Dim cStart As Integer
        Dim cEnd As Integer
        
        rStart = Selection(1).Row
        rEnd = Selection(Selection.Count).Row
        cStart = Selection(1).Column
        cEnd = Selection(Selection.Count).Column
    
        '列数からヘッダの作成
        result = result & "||"
        For c = cStart To cEnd
            result = result & PaddingText(" **" & Cells(rStart, c).Value & "**", " ", " ||", lenMap(c - cStart) + 5)
        Next c
        result = result & vbCrLf
    
        
        '選択範囲からボディを作成
        For r = rStart + 1 To rEnd
            result = result & "||"
            For c = cStart To cEnd
                result = result & PaddingText(" " & Cells(r, c).Value, " ", " ||", lenMap(c - cStart) + 5)
            Next c
            result = result & vbCrLf
        Next r
    
        'クリップボードに保存
        SaveToClipboard (result)

    End If
End Sub

'**********************************************************************************
'* Copy(PukiWiki)
'**********************************************************************************
Private Sub CopyByPukiWiki()
    
    '選択範囲の取得し、フォーマット変換
    If Selection.Count <> 0 Then
        '各列の最大文字列長を計算
        Dim lenMap As Variant
        lenMap = CalcLength()
        
        'TODO 各列の最大文字列長からきれいに整形したい
        Dim result As String
        result = ""
        
        Dim rStart As Integer
        Dim rEnd As Integer
        Dim cStart As Integer
        Dim cEnd As Integer
        
        rStart = Selection(1).Row
        rEnd = Selection(Selection.Count).Row
        cStart = Selection(1).Column
        cEnd = Selection(Selection.Count).Column
    
        '列数からヘッダの作成
        result = result & "|~"
        For c = cStart To cEnd
            result = result & PaddingText(" " & Cells(rStart, c).Value, " ", " |", lenMap(c - cStart) + 1)
        Next c
        result = result & vbCrLf
    
        
        '選択範囲からボディを作成
        For r = rStart + 1 To rEnd
            result = result & "| "
            For c = cStart To cEnd
                result = result & PaddingText(" " & Cells(r, c).Value, " ", " |", lenMap(c - cStart) + 1)
            Next c
            result = result & vbCrLf
        Next r
    
        'クリップボードに保存
        SaveToClipboard (result)

    End If
End Sub


'**********************************************************************************
'* Copy(XPlanner)
'**********************************************************************************
Private Sub CopyByXPlanner()
    
    '選択範囲の取得し、フォーマット変換
    If Selection.Count <> 0 Then
        '各列の最大文字列長を計算
        Dim lenMap As Variant
        lenMap = CalcLength()
        
        'TODO 各列の最大文字列長からきれいに整形したい
        Dim result As String
        result = ""
        
        Dim rStart As Integer
        Dim rEnd As Integer
        Dim cStart As Integer
        Dim cEnd As Integer
        
        rStart = Selection(1).Row
        rEnd = Selection(Selection.Count).Row
        cStart = Selection(1).Column
        cEnd = Selection(Selection.Count).Column
    
        '列数からヘッダの作成
        result = result & "|"
        For c = cStart To cEnd
            result = result & PaddingText(" *" & Cells(rStart, c).Value & "*", " ", " |", lenMap(c - cStart) + 3)
        Next c
        result = result & vbCrLf
    
        
        '選択範囲からボディを作成
        For r = rStart + 1 To rEnd
            result = result & "|"
            For c = cStart To cEnd
                result = result & PaddingText(" " & Cells(r, c).Value, " ", " |", lenMap(c - cStart) + 3)
            Next c
            result = result & vbCrLf
        Next r
    
        'クリップボードに保存
        SaveToClipboard (result)

    End If
End Sub


'**********************************************************************************
'* SaveToClipBoard
'**********************************************************************************
Private Sub SaveToClipboard(ByVal str As String)
    Dim CB As Object
    Set CB = New DataObject
    With CB
        .SetText str
        .PutInClipboard  'クリップボードに反映する
    End With
End Sub


'**********************************************************************************
'* CalcLength
'**********************************************************************************
Private Function CalcLength() As Variant
    Dim lenMap() As Integer
    Dim colCount As Integer
    
    colCount = Selection(Selection.Count).Column - Selection(1).Column
    
    ReDim lenMap(colCount)
    
    Dim maxLength As Integer
    Dim colLength As Integer
    Dim cIndex As Integer
    For c = Selection(1).Column To Selection(Selection.Count).Column
        maxLength = 0
        colLength = 0
        For r = Selection(1).Row To Selection(Selection.Count).Row
            colLength = LenB(StrConv(Cells(r, c).Value, vbFromUnicode))
            
            If maxLength < colLength Then
                maxLength = colLength
            End If
        Next r
        
        lenMap(c - Selection(1).Column) = maxLength
    Next c
    
    CalcLength = lenMap
End Function


'**********************************************************************************
'* PaddingText
'**********************************************************************************
Private Function PaddingText(ByVal text As String, ByVal rep As String, ByVal tail As String, ByVal maxLen As Integer)
    Dim ret As String
    Dim repLen As Integer
    
    repLen = maxLen - LenB(StrConv(text, vbFromUnicode))
    
    ret = text
    If repLen > 0 Then
        For i = 1 To repLen
            ret = ret & rep
        Next i
    End If
    ret = ret & tail
    
    PaddingText = ret
End Function


'**********************************************************************************
'* Utils
'**********************************************************************************
Private Sub DelAllMenu()
    DeleteMenu
    
    SetupMenu
End Sub

Private Sub SetDummy()
    '選択範囲の取得し、フォーマット変換
    Dim result As String
    
    For r = Selection(1).Row To Selection(Selection.Count).Row
        For c = Selection(1).Column To Selection(Selection.Count).Column
            Cells(r, c) = "あ" & r & "__" & c
        Next c
    Next r
End Sub
'**********************************************************************************


