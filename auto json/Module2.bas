Attribute VB_Name = "Module2"
Option Explicit

' 各シートの最後の列をコピーし、指定されたPIDで上書き(レコード追加)
Function addData(pid As String)

    Dim WS_Count As Integer
    Dim I As Integer

    '画面更新停止
    Application.ScreenUpdating = False
    
    WS_Count = ActiveWorkbook.Worksheets.Count

    For I = 2 To WS_Count

       '1行でもデータがあるかどう� 
        With ActiveWorkbook.Worksheets(I)
            If .Range("O6") <> "" Then
                '最終列を取� 
                Dim lastColNum As Integer
                lastColNum = .Range("O6").End(xlToRight).Column
                
                '行数を取� 
                Dim lastRowNum As Integer
                lastRowNum = .Range("B5").End(xlDown).Row
                
                
                '最終列を隣の列にコピー
                .Columns(lastColNum + 1).Insert Shift:=xlToRight
                .Range(.Cells(4, lastColNum), .Cells(lastRowNum, lastColNum)).Copy .Range(.Cells(4, lastColNum + 1), .Cells(lastRowNum, lastColNum + 1))
                
                
                'referenceNoとnationalidを書き換� 
                Dim rowCnt As Integer
                For rowCnt = 5 To lastRowNum
                    If .Range("C" & rowCnt) = "referenceNo" Then
                        .Cells(rowCnt, lastColNum + 1) = Left(.Cells(rowCnt, lastColNum), 2) & pid & Mid(.Cells(rowCnt, lastColNum), 13, 9)
                    ElseIf .Range("C" & rowCnt) = "nationalid" Then
                        .Cells(rowCnt, lastColNum + 1) = pid
                    End If
                Next rowCnt
              
        End If
       End With
    Next I

    '画面更新再開
    Application.ScreenUpdating = True

    MsgBox ("データを追加しました")

End Function


