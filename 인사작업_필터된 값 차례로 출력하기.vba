Sub 매크로3()
'
' 인사구역 및 가산점에서 가산점 부분 출력
'
    arr = Array("가평", "고양", "광명", "광주하남", "구리남양주", "군포의왕", "김포", "동두천양주", "부천", "성남", "수원", "시흥", "안산", "안성", "안양과천", "양평", "여주", "연천", "용인", "의정부", "이천", "파주", "평택", "포천", "화성오산")
    Columns("S:AF").Select   '출력부분 선택하기
    For Each arr_i In arr  ' 배열의 지역 하나씩 가져오기
        ActiveSheet.Range("$s$2:$N$1346").AutoFilter Field:=1, Criteria1:=arr_i   ' field - 반복하고자하는 필터 부분 / 출력전체 부분 설정
        Application.PrintCommunication = False
        With ActiveSheet.PageSetup
            .PrintTitleRows = "$1:$2"   ' 반복행 설정
            .PrintTitleColumns = ""
        End With
        Application.PrintCommunication = True
        ActiveSheet.PageSetup.PrintArea = ""
        Application.PrintCommunication = False
        With ActiveSheet.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""     '페이지 가운데 상단 출력
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = "&P / &N"   '페이지 쪽수 출력
            .RightFooter = ""
            .LeftMargin = Application.InchesToPoints(0.25)
            .RightMargin = Application.InchesToPoints(0.25)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            .PrintHeadings = False
            .PrintGridlines = False
            .PrintComments = xlPrintNoComments
            .PrintQuality = 600
            .CenterHorizontally = True
            .CenterVertically = False
            .Orientation = xlPortrait
            .Draft = False
            .PaperSize = xlPaperA4
            .FirstPageNumber = xlAutomatic
            .Order = xlDownThenOver
            .BlackAndWhite = False
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 0
            .PrintErrors = xlPrintErrorsDisplayed
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
            .EvenPage.LeftHeader.Text = ""
            .EvenPage.CenterHeader.Text = ""
            .EvenPage.RightHeader.Text = ""
            .EvenPage.LeftFooter.Text = ""
            .EvenPage.CenterFooter.Text = ""
            .EvenPage.RightFooter.Text = ""
            .FirstPage.LeftHeader.Text = ""
            .FirstPage.CenterHeader.Text = ""
            .FirstPage.RightHeader.Text = ""
            .FirstPage.LeftFooter.Text = ""
            .FirstPage.CenterFooter.Text = ""
            .FirstPage.RightFooter.Text = ""
        End With
        Application.PrintCommunication = True
        Selection.PrintOut Copies:=1, Collate:=True  ' 출력물 수량 설정
    Next
End Sub
