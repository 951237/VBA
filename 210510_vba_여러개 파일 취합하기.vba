Sub 엑셀데이터취합()

Dim fileNo As Variant
Dim i As Integer
Dim ingFile As Workbook
Dim SumSheet As Worksheet
Dim iRow As Integer

Set SumSheet = ThisWorkbook.Worksheets(1)   ' 취합파일명

On Error GoTo 에러처리

fileNo = Application.GetOpenFilename(Filefilter:="엑셀파일(*.xlsx*),*.xlsx*", MultiSelect:=True)

Application.ScreenUpdating = False
Application.DisplayAlerts = False

        For i = 1 To UBound(fileNo) ' 순서대로 불러오기
            Set ingFile = Workbooks.Open(Filename:=fileNo(i), ReadOnly:=True)
            
            iRow = 7 + i    '붙여넣기 위치 행
    
    
            With ingFile.Sheets(1)  '데이터 복사하기 
                .Range("B8:BA8").Copy   ' 한행만 복사
                
            End With
            SumSheet.Cells(iRow, 2).PasteSpecial    ' 취합파일에 붙여넣기
  
            ingFile.Close   '소스파일 닫기
        Next i
       
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        
        MsgBox "파일 취합이 완료되었습니다"
        Exit Sub
    
에러처리:
    MsgBox "파일을 선택하지 않았습니다"
End Sub