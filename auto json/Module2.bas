Attribute VB_Name = "Module2"
Option Explicit

' ŠeƒV[ƒg‚ÌÅŒã‚Ì—ñ‚ğƒRƒs[‚µAw’è‚³‚ê‚½PID‚Åã‘‚«(ƒŒƒR[ƒh’Ç‰Á)
Function addData(pid As String)

    Dim WS_Count As Integer
    Dim I As Integer

    '‰æ–ÊXV’â~
    Application.ScreenUpdating = False
    
    WS_Count = ActiveWorkbook.Worksheets.Count

    For I = 2 To WS_Count

       '1s‚Å‚àƒf[ƒ^‚ª‚ ‚é‚©‚Ç‚¤‚ 
        With ActiveWorkbook.Worksheets(I)
            If .Range("O6") <> "" Then
                'ÅI—ñ‚ğæ“ 
                Dim lastColNum As Integer
                lastColNum = .Range("O6").End(xlToRight).Column
                
                's”‚ğæ“ 
                Dim lastRowNum As Integer
                lastRowNum = .Range("B5").End(xlDown).Row
                
                
                'ÅI—ñ‚ğ—×‚Ì—ñ‚ÉƒRƒs[
                .Columns(lastColNum + 1).Insert Shift:=xlToRight
                .Range(.Cells(4, lastColNum), .Cells(lastRowNum, lastColNum)).Copy .Range(.Cells(4, lastColNum + 1), .Cells(lastRowNum, lastColNum + 1))
                
                
                'referenceNo‚Ænationalid‚ğ‘‚«Š·‚ 
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

    '‰æ–ÊXVÄŠJ
    Application.ScreenUpdating = True

    MsgBox ("ƒf[ƒ^‚ğ’Ç‰Á‚µ‚Ü‚µ‚½")

End Function


