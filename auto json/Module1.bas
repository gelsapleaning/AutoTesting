Attribute VB_Name = "Module1"
Option Explicit

'------------------------------
' ’è”’è‹`
'------------------------------
Const ColPosLevel As Integer = 1            'Level
Const ColPosItem As Integer = 2             'Item Name
Const ColPosDescription As Integer = 3      'Description
Const ColPosMandatory As Integer = 4        'Mandatory
Const ColPosDataType As Integer = 5         'DataType
Const ColPosArray As Integer = 6            'Array
Const ColPosMinLength As Integer = 7        'MinLength
Const ColPosMaxLength As Integer = 8        'MaxLength
Const ColPosFormat As Integer = 9           'Format
Const ColPosAllowable As Integer = 10       'AllowableStrings
Const ColPosDefaultValue As Integer = 11    'DefaultValue
Const ColPosSampleValue As Integer = 12     'SampleValue
Const ColPosJPNote As Integer = 13          'JPNote
Const ColPosValue As Integer = 14           'Value

'------------------------------
' •Ï”’è‹`
'------------------------------
Dim myRange As Range


Function createJsonParent(param As String)

    Dim parentFolderName As String


    ' •Û‘¶æ‘I‘ğƒ_ƒCƒAƒƒO•\ 
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = ActiveWorkbook.Path
        If .Show = True Then
            parentFolderName = .SelectedItems(1)
        Else
            Exit Function
        End If
            
    End With


    If param = "ALL" Then
        Dim WS_Count As Integer
        Dim I As Integer
    
        ' Set WS_Count equal to the number of worksheets in the active
        ' workbook.
        WS_Count = ActiveWorkbook.Worksheets.Count
    
        ' Begin the loop.
        For I = 2 To WS_Count
    
           ' Insert your code here.
           ' The following line shows how to reference a sheet within
           ' the loop by displaying the worksheet name in a dialog box.
           If ActiveWorkbook.Worksheets(I).Range("O6") <> "" Then
               Call createJson(ActiveWorkbook.Worksheets(I).Name, parentFolderName)
           End If
        Next I
    Else
        Call createJson(param, parentFolderName)
    End If

End Function




'------------------------------
' [JSON¶¬]ƒ{ƒ^ƒ“
'------------------------------
Function createJson(sheetName As String, parentFolderName As String)

    Dim itemData As String

    Dim rowIdx As Long
    Dim columnIdx As Long
    Dim enCloseCnt As Integer
    Dim preLevel As Integer
    Dim manyFlg As Boolean
    Dim breakFlg As Boolean
    
    Dim fileName As String
    Dim folderName As String
    Dim fileNo  As Integer
    
    
    
    
    
    ' ‘ÎÛƒV[ƒg‘I‘ 
    Worksheets(sheetName).Activate

    ' ‘ÎÛƒf[ƒ^—Ìˆæ‘I‘ 
    With ActiveSheet
        Set myRange = .Range("B5", .Range("A1").SpecialCells(xlLastCell))
        fileName = .Range("C3") & ".txt"
        folderName = .Range("B3")
    End With
    
    '--------------------
    ' •ÒW‰Šú‰ 
    '--------------------
    itemData = ""
        
    '--------------------
    ' JSONŒ`®•ÒWˆ—
    '--------------------
    columnIdx = ColPosValue
    Do While Not myRange.Cells(2, columnIdx) = ""
    
        enCloseCnt = 0
        preLevel = 0
        manyFlg = False
        breakFlg = False
        
        ' n‚Ü‚è‚Ì‚­‚­‚èŠ‡ŒÊ"{"
        If itemData = "" Then
            itemData = "{"
        Else
            itemData = itemData & vbCrLf & "{"
        End If
    
        For rowIdx = 1 To myRange.Rows.Count
        
            With myRange.Rows(rowIdx)
                    
                ' 1.ƒf[ƒ^—Ìˆæƒ`ƒFƒbƒN
                If .Cells(ColPosLevel) = "" Then
                    Exit For
                End If
            
                ' 2.ƒ`ƒFƒbƒNˆ—
                If checkInputData(rowIdx, columnIdx) = False Then
                    Exit Function
                End If
            
                ' 2.5 ‹óƒf[ƒ^ƒ`ƒFƒbƒNˆ—
                If manyFlg = True And .Cells(columnIdx) = "#noItems" Then
                    GoTo Continue
                End If
            
            
                ' 3.•ÒWˆ—
                ' ‡@.ƒf[ƒ^‹æØ•ÒWˆ—
                If rowIdx = 0 Then
                    preLevel = getLevelNum(.Cells(ColPosLevel))
                Else
                    If preLevel = getLevelNum(.Cells(ColPosLevel)) Then
                        ' “¯ŠK‘w‚Ìê‡
                    
                        If manyFlg Then
                            ' ƒŠƒXƒg\‘¢‚Ìê‡
                            If breakFlg Then
                                ' ‘O‰ñƒuƒŒƒCƒN‚Ìê‡
                                itemData = itemData & "}" & "," & "{"
                                breakFlg = False
                            Else
                                ' ‘O‰ñƒuƒŒƒCƒN‚Å‚Í‚È‚¢ê‡
                                itemData = itemData & ","
                            End If
                        Else
                            ' ã‹LˆÈŠO‚Ìê‡
                            If (.Cells(ColPosDataType) = "") And (myRange.Rows(rowIdx - 1).Cells(ColPosDataType) = "") Then
                                ' ‘O‰ñ ¡‰ñ‹¤‚ÉŒ©o€–Ú‚Ìê‡i¨@‘O‰ñŒ©o€–Ú‚É‚Â‚¢‚ÄÚ×‚ª‹ó‚Ìê‡j
                                ' Š‡ŒÊ"}"•t‰Áˆ—
                                enCloseCnt = addStringEnClose(True, enCloseCnt, 1, itemData)
                            Else
                                itemData = itemData & ","
                            End If
                        End If
                    ElseIf preLevel > getLevelNum(.Cells(ColPosLevel)) Then
                        ' ‘O‰ñŠK‘w > ¡‰ñŠK‘w‚Ìê‡
                                                            
                        ' ƒŠƒXƒg\‘¢ ––’[”»’ 
                        If manyFlg Then
                            ' Š‡ŒÊ"}"•t‰Áˆ—
                            enCloseCnt = addStringEnClose(False, enCloseCnt, preLevel - getLevelNum(.Cells(ColPosLevel)), itemData)
                        
                            ' ƒŠƒXƒg‹æØ"]"•t‰ 
                            itemData = itemData & "]" & ","
                        
                            manyFlg = False
                        Else
                            ' Š‡ŒÊ"}"•t‰Áˆ—
                            enCloseCnt = addStringEnClose(True, enCloseCnt, preLevel - getLevelNum(.Cells(ColPosLevel)), itemData)
                        End If

                    End If
                End If
    
                ' ‡A.Key‹y‚ÑValueo—Íˆ—
                If .Cells(ColPosDataType) = "" Then
                    '--------------------
                    ' A.Œ©o€–Ú‚Ìê‡
                    '--------------------
                    itemData = itemData & """" & .Cells(ColPosItem) & """" & ":"
                
                    ' ”»’è‰Šú‰ 
                    manyFlg = False
                    breakFlg = False
                
                    ' ƒŠƒXƒg\‘¢”»’ 
                    If .Cells(ColPosArray) = "Many" Then
                        itemData = itemData & "["
                        manyFlg = True
                    End If
                                
                    itemData = itemData & "{"
                
                    ' ŠK‘w”ƒCƒ“ƒNƒŠƒƒ“ƒg
                    enCloseCnt = enCloseCnt + 1
                
                Else
                    '--------------------
                    ' B.Œ©o€–ÚˆÈŠO‚Ìê‡
                    '--------------------
                    Select Case .Cells(ColPosDataType)
                        Case "String"
                            ' •¶š—ñ‚Ìê‡
                            itemData = itemData & """" & .Cells(ColPosItem) & """" & ":" & """" & .Cells(columnIdx) & """"
                        Case Else
                            ' ã‹LˆÈŠO‚Ìê‡
                            itemData = itemData & """" & .Cells(ColPosItem) & """" & ":" & .Cells(columnIdx)
                    End Select
                
                    ' ƒŠƒXƒg\‘¢ Break‚Ìê‡
                    If manyFlg And .Cells(ColPosArray) = "Break" Then
                        breakFlg = True
                    End If
                End If
            
                ' ŠK‘wi¡‰ñj‘Ò” 
                preLevel = getLevelNum(.Cells(ColPosLevel))
        
            End With

Continue:
        
        Next rowIdx
        
        ' ƒŠƒXƒg\‘¢ ––’[”»’ 
        If breakFlg Then
            itemData = itemData & "}"
            itemData = itemData & "]"

            ' ŠK‘w”ƒfƒNƒŠƒƒ“ƒg
            enCloseCnt = enCloseCnt - 1
        End If

        ' Š‡ŒÊ"}"•t‰Áˆ—
        enCloseCnt = addStringEnClose(False, enCloseCnt, enCloseCnt, itemData)
        
        ' ƒf[ƒ^‚­‚­‚è‚Ì•Â‚ßŠ‡ŒÊ"}"
        itemData = itemData & "}"
        columnIdx = columnIdx + 1
    
    Loop

    
    
'    fileName = Application.GetSaveAsFilename( _
'                                FileFilter:="Text Files (*.txt), *.txt", _
'                                FilterIndex:=1, _
'                                InitialFileName:=sheetName & ".txt", _
'                                Title:="JSONƒtƒ@ƒCƒ‹‚Ì•Û‘¶")
    
    Dim fullFolderName As String
    fullFolderName = parentFolderName & "\" & folderName
    If Dir(fullFolderName, vbDirectory) = "" Then
        MkDir fullFolderName
    End If
 
 
    '--------------------
    ' ƒtƒ@ƒCƒ‹‘
    '--------------------
    ' FreeNo‚ÌŠ„‚è“–‚ 
    fileNo = FreeFile
    
    ' ƒtƒ@ƒCƒ‹o— 
    Open fullFolderName & "\" & fileName For Output As #fileNo
    Print #fileNo, itemData
    Close fileNo
    

End Function

'------------------------------
' “ü—Íƒ`ƒFƒbƒN
'------------------------------
Function checkInputData(ByVal rowIdx As Integer, ByVal columnIdx As Integer) As Boolean
    
    With myRange.Rows(rowIdx)
    
        ' Œ©o€–Ú”»’èiMandatory€–Ú‚É‚æ‚é”»’èj
        If .Cells(ColPosMandatory) = "" Then
            checkInputData = True
            Exit Function
        End If
        
        ' –¢“ü—Íƒ`ƒFƒbƒN
'        If .Cells(ColPosMandatory) = "Yes" And .Cells(columnIdx) = "" Then
'            MsgBox (rowIdx & "s–Úi" & .Cells(ColPosItem) & ")‚Ì’l‚ª–¢“ü—Í‚Å‚·B")
'            checkInputData = False
'            Exit Function
'        End If
    
        ' —LŒøŒ…”ƒ`ƒFƒbƒN
        If .Cells(columnIdx) <> "" Then
            If (.Cells(ColPosMinLength) <> "NA") And (.Cells(ColPosMaxLength) <> "NA") Then
                ' ’è‹`’l‚ª"NA"‚Å‚È‚¢ê‡
                If (Len(.Cells(columnIdx)) < .Cells(ColPosMinLength)) Or (.Cells(ColPosMaxLength) < Len(.Cells(columnIdx))) Then
                    MsgBox (rowIdx & "s–Úi" & .Cells(ColPosItem) & ")‚Ì’l‚É‚Â‚¢‚ÄA—LŒøŒ…”" & .Cells(ColPosMinLength) & "Œ…`" & .Cells(ColPosMaxLength) & "Œ…‚Ì”ÍˆÍŠO‚Æ‚È‚Á‚Ä‚¢‚Ü‚·B")
                    checkInputData = False
                    Exit Function
                End If
            End If
        End If

    End With
               
    checkInputData = True

End Function

'------------------------------
' ŠK‘w”æ“ 
'------------------------------
Function getLevelNum(ByVal level As String) As Integer
    
    getLevelNum = CInt(Mid(level, InStr(level, "L") + 1))
    
End Function

'------------------------------
' "}"•t‰Áˆ—
'------------------------------
Function addStringEnClose(ByVal addComma As Boolean, ByVal enAllCloseCnt As Integer, ByVal enCloseCnt As Integer, ByRef itemData As String) As Integer

    Dim wkEnCloseCnt As Integer
    
    For wkEnCloseCnt = 1 To enCloseCnt
        If addComma Then
            itemData = itemData & "}" + ","
        Else
            itemData = itemData & "}"
        End If
    Next wkEnCloseCnt
    
    addStringEnClose = enAllCloseCnt - enCloseCnt

End Function
