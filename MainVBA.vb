
Private Sub HelpBtn_Click()
    
    ' หา path ของ ไฟล์ WI
    Sheets("Main").Select                                ' ถ้าเลือกให้เปลี่ยน path
    Range("E45").Select
    PathWI = ActiveCell.Value
    PathWI = PathWI & "\LPDM manual.pdf"

    ThisWorkbook.FollowHyperlink PathWI
End Sub



Private Sub LaunchBtn_Click()
    MainForm.Show 'vbModeless   'showmodal' property ให้เป็น false หรือใช้ vbModeless
End Sub



Private Sub DataMartPathBtn_Click()
' ส่วนของการเปิด file
With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "Select a Database Folder"
    .AllowMultiSelect = False
    .InitialFileName = "" '<~~ The start folder path for the file picker.
    If .Show = -1 Then 'GoTo NextCode
        If Right(.SelectedItems(1), 1) = "\" Then
            '
            MsgBox ("Please debug it")
            'SaveFolder = .SelectedItems(1)    'เพิ่มขึ้นมา
            'DBpath = .SelectedItems(1)
        Else
            ' กรณีที่เลือก foloder
            SaveFolder = .SelectedItems(1)    'เพิ่มขึ้นมา เพื่อ assign ค่าให้ global variable
            'MEApath = .SelectedItems(1) & "\" & "MEA" & "\" & Year & "\" & "MEA" & Year & Month & "stat.xlsx"
        End If
    End If
    
    ' ทำกรณีเลือก และ ไม่เลือก folder
    If .SelectedItems.Count = 0 Then
        MsgBox "Canceled by user"                        ' ถ้าไม่เลือก ให้คง folder เดิมไว้
    Else
        Sheets("Main").Select                                ' ถ้าเลือกให้เปลี่ยน path
        Range("E46").Select
        ActiveCell.FormulaR1C1 = SaveFolder
    End If
End With

End Sub

Private Sub DataWarehousePathBtn_Click()
' ส่วนของการเปิด file
With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "Select a Database Folder"
    .AllowMultiSelect = False
    .InitialFileName = "" '<~~ The start folder path for the file picker.
    If .Show = -1 Then 'GoTo NextCode
        If Right(.SelectedItems(1), 1) = "\" Then
            '
            DBFolder = .SelectedItems(1)    'เพิ่มขึ้นมา
            DBpath = .SelectedItems(1)
        Else
            ' กรณีที่เลือก foloder
            DBFolder = .SelectedItems(1)    'เพิ่มขึ้นมา เพื่อ assign ค่าให้ global variable
            'MEApath = .SelectedItems(1) & "\" & "MEA" & "\" & Year & "\" & "MEA" & Year & Month & "stat.xlsx"
        End If
    End If
    
    ' ทำกรณีเลือก และ ไม่เลือก folder
    If .SelectedItems.Count = 0 Then
        MsgBox "Canceled by user"                        ' ถ้าไม่เลือก ให้คง folder เดิมไว้
    Else
        Sheets("Main").Select                                ' ถ้าเลือกให้เปลี่ยน path
        Range("E45").Select
        ActiveCell.FormulaR1C1 = DBFolder
    End If
End With
End Sub

Private Sub VisualizeBtn_Click()
    UserForm1.Show
End Sub

