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
' �萔��`
'------------------------------
Const pSheetName = "Main"       '�V�[�g�� [���C��]

'------------------------------
' [�V�[�g�Ǎ�]�{�^��
'------------------------------
Private Sub cmdSheetList_Click()

    Dim st As Object
    
    cmbSheetName.Clear
    
    ' ���݂���V�[�g���̎� 
    For Each st In Sheets
    
       If st.Name <> pSheetName Then
       
            cmbSheetName.AddItem (st.Name)
            
            
       End If
        
    Next
    cmbSheetName.AddItem ("ALL")

End Sub

'------------------------------
' [JSON����]�{�^��
'------------------------------
Private Sub cmdCreateJson_Click()

    ' ���̓`�F�b�N
    If Me.cmbSheetName.Text = "" Then
        MsgBox ("�����Ώۂ̃V�[�g���I������Ă��܂���B")
        Exit Sub
    End If
    
    ' JSON�t�@�C���쐬����
    Call createJsonParent(Me.cmbSheetName.Text)
    
End Sub


'------------------------------
' [�f�[�^�ǉ�]�{�^��
'------------------------------
Private Sub btnDataAdd_Click()
    
    Call addData(Me.txtPID.Text)

End Sub
