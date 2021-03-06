Attribute VB_Name = "Module1"
Option Explicit

'------------------------------
' 定数定義
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
' 変数定義
'------------------------------
Dim myRange As Range


Function createJsonParent(param As String)

    Dim parentFolderName As String


    ' 保存先選択ダイアログ表� 
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
' [JSON生成]ボタン
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
    
    
    
    
    
    ' 対象シート選� 
    Worksheets(sheetName).Activate

    ' 対象データ領域選� 
    With ActiveSheet
        Set myRange = .Range("B5", .Range("A1").SpecialCells(xlLastCell))
        fileName = .Range("C3") & ".txt"
        folderName = .Range("B3")
    End With
    
    '--------------------
    ' 編集初期� 
    '--------------------
    itemData = ""
        
    '--------------------
    ' JSON形式編集処理
    '--------------------
    columnIdx = ColPosValue
    Do While Not myRange.Cells(2, columnIdx) = ""
    
        enCloseCnt = 0
        preLevel = 0
        manyFlg = False
        breakFlg = False
        
        ' 始まりのくくり括弧"{"
        If itemData = "" Then
            itemData = "{"
        Else
            itemData = itemData & vbCrLf & "{"
        End If
    
        For rowIdx = 1 To myRange.Rows.Count
        
            With myRange.Rows(rowIdx)
                    
                ' 1.データ領域チェック
                If .Cells(ColPosLevel) = "" Then
                    Exit For
                End If
            
                ' 2.チェック処理
                If checkInputData(rowIdx, columnIdx) = False Then
                    Exit Function
                End If
            
                ' 2.5 空データチェック処理
                If manyFlg = True And .Cells(columnIdx) = "#noItems" Then
                    GoTo Continue
                End If
            
            
                ' 3.編集処理
                ' �@.データ区切編集処理
                If rowIdx = 0 Then
                    preLevel = getLevelNum(.Cells(ColPosLevel))
                Else
                    If preLevel = getLevelNum(.Cells(ColPosLevel)) Then
                        ' 同階層の場合
                    
                        If manyFlg Then
                            ' リスト構造の場合
                            If breakFlg Then
                                ' 前回ブレイクの場合
                                itemData = itemData & "}" & "," & "{"
                                breakFlg = False
                            Else
                                ' 前回ブレイクではない場合
                                itemData = itemData & ","
                            End If
                        Else
                            ' 上記以外の場合
                            If (.Cells(ColPosDataType) = "") And (myRange.Rows(rowIdx - 1).Cells(ColPosDataType) = "") Then
                                ' 前回 今回共に見出項目の場合（→　前回見出項目について詳細が空の場合）
                                ' 括弧"}"付加処理
                                enCloseCnt = addStringEnClose(True, enCloseCnt, 1, itemData)
                            Else
                                itemData = itemData & ","
                            End If
                        End If
                    ElseIf preLevel > getLevelNum(.Cells(ColPosLevel)) Then
                        ' 前回階層 > 今回階層の場合
                                                            
                        ' リスト構造 末端判� 
                        If manyFlg Then
                            ' 括弧"}"付加処理
                            enCloseCnt = addStringEnClose(False, enCloseCnt, preLevel - getLevelNum(.Cells(ColPosLevel)), itemData)
                        
                            ' リスト区切"]"付� 
                            itemData = itemData & "]" & ","
                        
                            manyFlg = False
                        Else
                            ' 括弧"}"付加処理
                            enCloseCnt = addStringEnClose(True, enCloseCnt, preLevel - getLevelNum(.Cells(ColPosLevel)), itemData)
                        End If

                    End If
                End If
    
                ' �A.Key及びValue出力処理
                If .Cells(ColPosDataType) = "" Then
                    '--------------------
                    ' A.見出項目の場合
                    '--------------------
                    itemData = itemData & """" & .Cells(ColPosItem) & """" & ":"
                
                    ' 判定初期� 
                    manyFlg = False
                    breakFlg = False
                
                    ' リスト構造判� 
                    If .Cells(ColPosArray) = "Many" Then
                        itemData = itemData & "["
                        manyFlg = True
                    End If
                                
                    itemData = itemData & "{"
                
                    ' 階層数インクリメント
                    enCloseCnt = enCloseCnt + 1
                
                Else
                    '--------------------
                    ' B.見出項目以外の場合
                    '--------------------
                    Select Case .Cells(ColPosDataType)
                        Case "String"
                            ' 文字列の場合
                            itemData = itemData & """" & .Cells(ColPosItem) & """" & ":" & """" & .Cells(columnIdx) & """"
                        Case Else
                            ' 上記以外の場合
                            itemData = itemData & """" & .Cells(ColPosItem) & """" & ":" & .Cells(columnIdx)
                    End Select
                
                    ' リスト構造 Breakの場合
                    If manyFlg And .Cells(ColPosArray) = "Break" Then
                        breakFlg = True
                    End If
                End If
            
                ' 階層（今回）待� 
                preLevel = getLevelNum(.Cells(ColPosLevel))
        
            End With

Continue:
        
        Next rowIdx
        
        ' リスト構造 末端判� 
        If breakFlg Then
            itemData = itemData & "}"
            itemData = itemData & "]"

            ' 階層数デクリメント
            enCloseCnt = enCloseCnt - 1
        End If

        ' 括弧"}"付加処理
        enCloseCnt = addStringEnClose(False, enCloseCnt, enCloseCnt, itemData)
        
        ' データくくりの閉め括弧"}"
        itemData = itemData & "}"
        columnIdx = columnIdx + 1
    
    Loop

    
    
'    fileName = Application.GetSaveAsFilename( _
'                                FileFilter:="Text Files (*.txt), *.txt", _
'                                FilterIndex:=1, _
'                                InitialFileName:=sheetName & ".txt", _
'                                Title:="JSONファイルの保存")
    
    Dim fullFolderName As String
    fullFolderName = parentFolderName & "\" & folderName
    If Dir(fullFolderName, vbDirectory) = "" Then
        MkDir fullFolderName
    End If
 
 
    '--------------------
    ' ファイル書込
    '--------------------
    ' FreeNoの割り当� 
    fileNo = FreeFile
    
    ' ファイル出� 
    Open fullFolderName & "\" & fileName For Output As #fileNo
    Print #fileNo, itemData
    Close fileNo
    

End Function

'------------------------------
' 入力チェック
'------------------------------
Function checkInputData(ByVal rowIdx As Integer, ByVal columnIdx As Integer) As Boolean
    
    With myRange.Rows(rowIdx)
    
        ' 見出項目判定（Mandatory項目による判定）
        If .Cells(ColPosMandatory) = "" Then
            checkInputData = True
            Exit Function
        End If
        
        ' 未入力チェック
'        If .Cells(ColPosMandatory) = "Yes" And .Cells(columnIdx) = "" Then
'            MsgBox (rowIdx & "行目（" & .Cells(ColPosItem) & ")の値が未入力です。")
'            checkInputData = False
'            Exit Function
'        End If
    
        ' 有効桁数チェック
        If .Cells(columnIdx) <> "" Then
            If (.Cells(ColPosMinLength) <> "NA") And (.Cells(ColPosMaxLength) <> "NA") Then
                ' 定義値が"NA"でない場合
                If (Len(.Cells(columnIdx)) < .Cells(ColPosMinLength)) Or (.Cells(ColPosMaxLength) < Len(.Cells(columnIdx))) Then
                    MsgBox (rowIdx & "行目（" & .Cells(ColPosItem) & ")の値について、有効桁数" & .Cells(ColPosMinLength) & "桁〜" & .Cells(ColPosMaxLength) & "桁の範囲外となっています。")
                    checkInputData = False
                    Exit Function
                End If
            End If
        End If

    End With
               
    checkInputData = True

End Function

'------------------------------
' 階層数取� 
'------------------------------
Function getLevelNum(ByVal level As String) As Integer
    
    getLevelNum = CInt(Mid(level, InStr(level, "L") + 1))
    
End Function

'------------------------------
' "}"付加処理
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
