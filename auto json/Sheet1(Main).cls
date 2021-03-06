VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'------------------------------
' 定数定義
'------------------------------
Const pSheetName = "Main"       'シート名 [メイン]

'------------------------------
' [シート読込]ボタン
'------------------------------
Private Sub cmdSheetList_Click()

    Dim st As Object
    
    cmbSheetName.Clear
    
    ' 存在するシート名の取� 
    For Each st In Sheets
    
       If st.Name <> pSheetName Then
       
            cmbSheetName.AddItem (st.Name)
            
            
       End If
        
    Next
    cmbSheetName.AddItem ("ALL")

End Sub

'------------------------------
' [JSON生成]ボタン
'------------------------------
Private Sub cmdCreateJson_Click()

    ' 入力チェック
    If Me.cmbSheetName.Text = "" Then
        MsgBox ("処理対象のシートが選択されていません。")
        Exit Sub
    End If
    
    ' JSONファイル作成処理
    Call createJsonParent(Me.cmbSheetName.Text)
    
End Sub


'------------------------------
' [データ追加]ボタン
'------------------------------
Private Sub btnDataAdd_Click()
    
    Call addData(Me.txtPID.Text)

End Sub

