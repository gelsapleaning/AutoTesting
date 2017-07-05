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
' ’è”’è‹`
'------------------------------
Const pSheetName = "Main"       'ƒV[ƒg–¼ [ƒƒCƒ“]

'------------------------------
' [ƒV[ƒg“Ç]ƒ{ƒ^ƒ“
'------------------------------
Private Sub cmdSheetList_Click()

    Dim st As Object
    
    cmbSheetName.Clear
    
    ' ‘¶İ‚·‚éƒV[ƒg–¼‚Ìæ“ 
    For Each st In Sheets
    
       If st.Name <> pSheetName Then
       
            cmbSheetName.AddItem (st.Name)
            
            
       End If
        
    Next
    cmbSheetName.AddItem ("ALL")

End Sub

'------------------------------
' [JSON¶¬]ƒ{ƒ^ƒ“
'------------------------------
Private Sub cmdCreateJson_Click()

    ' “ü—Íƒ`ƒFƒbƒN
    If Me.cmbSheetName.Text = "" Then
        MsgBox ("ˆ—‘ÎÛ‚ÌƒV[ƒg‚ª‘I‘ğ‚³‚ê‚Ä‚¢‚Ü‚¹‚ñB")
        Exit Sub
    End If
    
    ' JSONƒtƒ@ƒCƒ‹ì¬ˆ—
    Call createJsonParent(Me.cmbSheetName.Text)
    
End Sub


'------------------------------
' [ƒf[ƒ^’Ç‰Á]ƒ{ƒ^ƒ“
'------------------------------
Private Sub btnDataAdd_Click()
    
    Call addData(Me.txtPID.Text)

End Sub

