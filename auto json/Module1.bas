Attribute VB_Name = "Module1"
Option Explicit

'------------------------------
' �萔��`
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
' �ϐ���`
'------------------------------
Dim myRange As Range


Function createJsonParent(param As String)

    Dim parentFolderName As String


    ' �ۑ���I���_�C�A���O�\� 
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
' [JSON����]�{�^��
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
    
    
    
    
    
    ' �ΏۃV�[�g�I� 
    Worksheets(sheetName).Activate

    ' �Ώۃf�[�^�̈�I� 
    With ActiveSheet
        Set myRange = .Range("B5", .Range("A1").SpecialCells(xlLastCell))
        fileName = .Range("C3") & ".txt"
        folderName = .Range("B3")
    End With
    
    '--------------------
    ' �ҏW����� 
    '--------------------
    itemData = ""
        
    '--------------------
    ' JSON�`���ҏW����
    '--------------------
    columnIdx = ColPosValue
    Do While Not myRange.Cells(2, columnIdx) = ""
    
        enCloseCnt = 0
        preLevel = 0
        manyFlg = False
        breakFlg = False
        
        ' �n�܂�̂����芇��"{"
        If itemData = "" Then
            itemData = "{"
        Else
            itemData = itemData & vbCrLf & "{"
        End If
    
        For rowIdx = 1 To myRange.Rows.Count
        
            With myRange.Rows(rowIdx)
                    
                ' 1.�f�[�^�̈�`�F�b�N
                If .Cells(ColPosLevel) = "" Then
                    Exit For
                End If
            
                ' 2.�`�F�b�N����
                If checkInputData(rowIdx, columnIdx) = False Then
                    Exit Function
                End If
            
                ' 2.5 ��f�[�^�`�F�b�N����
                If manyFlg = True And .Cells(columnIdx) = "#noItems" Then
                    GoTo Continue
                End If
            
            
                ' 3.�ҏW����
                ' �@.�f�[�^��ؕҏW����
                If rowIdx = 0 Then
                    preLevel = getLevelNum(.Cells(ColPosLevel))
                Else
                    If preLevel = getLevelNum(.Cells(ColPosLevel)) Then
                        ' ���K�w�̏ꍇ
                    
                        If manyFlg Then
                            ' ���X�g�\���̏ꍇ
                            If breakFlg Then
                                ' �O��u���C�N�̏ꍇ
                                itemData = itemData & "}" & "," & "{"
                                breakFlg = False
                            Else
                                ' �O��u���C�N�ł͂Ȃ��ꍇ
                                itemData = itemData & ","
                            End If
                        Else
                            ' ��L�ȊO�̏ꍇ
                            If (.Cells(ColPosDataType) = "") And (myRange.Rows(rowIdx - 1).Cells(ColPosDataType) = "") Then
                                ' �O�� ���񋤂Ɍ��o���ڂ̏ꍇ�i���@�O�񌩏o���ڂɂ��ďڍׂ���̏ꍇ�j
                                ' ����"}"�t������
                                enCloseCnt = addStringEnClose(True, enCloseCnt, 1, itemData)
                            Else
                                itemData = itemData & ","
                            End If
                        End If
                    ElseIf preLevel > getLevelNum(.Cells(ColPosLevel)) Then
                        ' �O��K�w > ����K�w�̏ꍇ
                                                            
                        ' ���X�g�\�� ���[��� 
                        If manyFlg Then
                            ' ����"}"�t������
                            enCloseCnt = addStringEnClose(False, enCloseCnt, preLevel - getLevelNum(.Cells(ColPosLevel)), itemData)
                        
                            ' ���X�g���"]"�t� 
                            itemData = itemData & "]" & ","
                        
                            manyFlg = False
                        Else
                            ' ����"}"�t������
                            enCloseCnt = addStringEnClose(True, enCloseCnt, preLevel - getLevelNum(.Cells(ColPosLevel)), itemData)
                        End If

                    End If
                End If
    
                ' �A.Key�y��Value�o�͏���
                If .Cells(ColPosDataType) = "" Then
                    '--------------------
                    ' A.���o���ڂ̏ꍇ
                    '--------------------
                    itemData = itemData & """" & .Cells(ColPosItem) & """" & ":"
                
                    ' ���菉��� 
                    manyFlg = False
                    breakFlg = False
                
                    ' ���X�g�\����� 
                    If .Cells(ColPosArray) = "Many" Then
                        itemData = itemData & "["
                        manyFlg = True
                    End If
                                
                    itemData = itemData & "{"
                
                    ' �K�w���C���N�������g
                    enCloseCnt = enCloseCnt + 1
                
                Else
                    '--------------------
                    ' B.���o���ڈȊO�̏ꍇ
                    '--------------------
                    Select Case .Cells(ColPosDataType)
                        Case "String"
                            ' ������̏ꍇ
                            itemData = itemData & """" & .Cells(ColPosItem) & """" & ":" & """" & .Cells(columnIdx) & """"
                        Case Else
                            ' ��L�ȊO�̏ꍇ
                            itemData = itemData & """" & .Cells(ColPosItem) & """" & ":" & .Cells(columnIdx)
                    End Select
                
                    ' ���X�g�\�� Break�̏ꍇ
                    If manyFlg And .Cells(ColPosArray) = "Break" Then
                        breakFlg = True
                    End If
                End If
            
                ' �K�w�i����j�Ҕ 
                preLevel = getLevelNum(.Cells(ColPosLevel))
        
            End With

Continue:
        
        Next rowIdx
        
        ' ���X�g�\�� ���[��� 
        If breakFlg Then
            itemData = itemData & "}"
            itemData = itemData & "]"

            ' �K�w���f�N�������g
            enCloseCnt = enCloseCnt - 1
        End If

        ' ����"}"�t������
        enCloseCnt = addStringEnClose(False, enCloseCnt, enCloseCnt, itemData)
        
        ' �f�[�^������̕ߊ���"}"
        itemData = itemData & "}"
        columnIdx = columnIdx + 1
    
    Loop

    
    
'    fileName = Application.GetSaveAsFilename( _
'                                FileFilter:="Text Files (*.txt), *.txt", _
'                                FilterIndex:=1, _
'                                InitialFileName:=sheetName & ".txt", _
'                                Title:="JSON�t�@�C���̕ۑ�")
    
    Dim fullFolderName As String
    fullFolderName = parentFolderName & "\" & folderName
    If Dir(fullFolderName, vbDirectory) = "" Then
        MkDir fullFolderName
    End If
 
 
    '--------------------
    ' �t�@�C������
    '--------------------
    ' FreeNo�̊��蓖� 
    fileNo = FreeFile
    
    ' �t�@�C���o� 
    Open fullFolderName & "\" & fileName For Output As #fileNo
    Print #fileNo, itemData
    Close fileNo
    

End Function

'------------------------------
' ���̓`�F�b�N
'------------------------------
Function checkInputData(ByVal rowIdx As Integer, ByVal columnIdx As Integer) As Boolean
    
    With myRange.Rows(rowIdx)
    
        ' ���o���ڔ���iMandatory���ڂɂ�锻��j
        If .Cells(ColPosMandatory) = "" Then
            checkInputData = True
            Exit Function
        End If
        
        ' �����̓`�F�b�N
'        If .Cells(ColPosMandatory) = "Yes" And .Cells(columnIdx) = "" Then
'            MsgBox (rowIdx & "�s�ځi" & .Cells(ColPosItem) & ")�̒l�������͂ł��B")
'            checkInputData = False
'            Exit Function
'        End If
    
        ' �L�������`�F�b�N
        If .Cells(columnIdx) <> "" Then
            If (.Cells(ColPosMinLength) <> "NA") And (.Cells(ColPosMaxLength) <> "NA") Then
                ' ��`�l��"NA"�łȂ��ꍇ
                If (Len(.Cells(columnIdx)) < .Cells(ColPosMinLength)) Or (.Cells(ColPosMaxLength) < Len(.Cells(columnIdx))) Then
                    MsgBox (rowIdx & "�s�ځi" & .Cells(ColPosItem) & ")�̒l�ɂ��āA�L������" & .Cells(ColPosMinLength) & "���`" & .Cells(ColPosMaxLength) & "���͈̔͊O�ƂȂ��Ă��܂��B")
                    checkInputData = False
                    Exit Function
                End If
            End If
        End If

    End With
               
    checkInputData = True

End Function

'------------------------------
' �K�w���� 
'------------------------------
Function getLevelNum(ByVal level As String) As Integer
    
    getLevelNum = CInt(Mid(level, InStr(level, "L") + 1))
    
End Function

'------------------------------
' "}"�t������
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
