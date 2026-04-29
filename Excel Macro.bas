Attribute VB_Name = "Module1"
Sub T_Space()
Attribute T_Space.VB_ProcData.VB_Invoke_Func = " \n14"
' 공백 나누기 매크로

    Selection.TextToColumns Destination:=ActiveCell, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
End Sub

Sub T_Comma()
' Comma 나누기 매크로

    Selection.TextToColumns Destination:=ActiveCell, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
End Sub

Sub T_Diamond()
' ◇ 나누기 매크로

    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=ActiveCell, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="◇", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
End Sub

Sub T_Star()
' ☆ 나누기 매크로

    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=ActiveCell, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="☆", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
End Sub

Sub T_SS()
' § 나누기 매크로

    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=ActiveCell, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="§", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
End Sub


Sub T_UdBar()
' _ 나누기 매크로

    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=ActiveCell, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="_", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
End Sub

Sub T_Dash()
' - 나누기 매크로

    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=ActiveCell, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
End Sub
Sub T_Cancel()
Attribute T_Cancel.VB_ProcData.VB_Invoke_Func = " \n14"
' 나누기 매크로 초기화

    Selection.TextToColumns Destination:=ActiveCell, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
End Sub

Sub 셀스타일삭제()

   Dim 스타일 As Style
   Dim 개수 As Long
   For Each 스타일 In ActiveWorkbook.Styles
       If 스타일.BuiltIn = False Then
           On Error Resume Next
           스타일.Delete
           개수 = 개수 + 1
           On Error GoTo 0
       End If
   Next
   MsgBox 개수 & "개의 불필요한 셀 스타일 제거 완료"
End Sub

Sub PreviewRGB()
    Dim i As Long
    For i = 2 To Cells(Rows.Count, "A").End(xlUp).Row
        Cells(i, "D").Interior.Color = RGB( _
            Cells(i, "A"), Cells(i, "B"), Cells(i, "C"))
    Next i
End Sub

