Attribute VB_Name = "Module2"
Option Explicit

' �e�V�[�g�̍Ō�̗���R�s�[���A�w�肳�ꂽPID�ŏ㏑��(���R�[�h�ǉ�)
Function addData(pid As String)

    Dim WS_Count As Integer
    Dim I As Integer

    '��ʍX�V��~
    Application.ScreenUpdating = False
    
    WS_Count = ActiveWorkbook.Worksheets.Count

    For I = 2 To WS_Count

       '1�s�ł��f�[�^�����邩�ǂ�� 
        With ActiveWorkbook.Worksheets(I)
            If .Range("O6") <> "" Then
                '�ŏI����� 
                Dim lastColNum As Integer
                lastColNum = .Range("O6").End(xlToRight).Column
                
                '�s������ 
                Dim lastRowNum As Integer
                lastRowNum = .Range("B5").End(xlDown).Row
                
                
                '�ŏI���ׂ̗�ɃR�s�[
                .Columns(lastColNum + 1).Insert Shift:=xlToRight
                .Range(.Cells(4, lastColNum), .Cells(lastRowNum, lastColNum)).Copy .Range(.Cells(4, lastColNum + 1), .Cells(lastRowNum, lastColNum + 1))
                
                
                'referenceNo��nationalid��������� 
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

    '��ʍX�V�ĊJ
    Application.ScreenUpdating = True

    MsgBox ("�f�[�^��ǉ����܂���")

End Function


