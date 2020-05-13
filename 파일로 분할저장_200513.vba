sub Sep_Data()

    Dim MyWB, OPWB as Workbook
    Dim MyST, OPST as Workbook
    Dim W_Rng as Range
    Dim MyRow, OPRow as Interger
    Dim F_Name AS Sting, L_Name as String, Name as String
    Dim i as Interger
    Dim F_dir as FileDialog

    '분할 파일을 저장할 폴더 지정
    Set F_dir = Application.FileDialog(msoFileDialogFolderPicker)
    F_dir.AllowMutiSelect = False

    
end sub