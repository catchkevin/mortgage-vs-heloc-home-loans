Sub BuildMortgageVsHeloc_Final()
    Dim wb As Workbook
    Dim wsInputs As Worksheet, wsBills As Worksheet, wsSide As Worksheet
    Dim lastRow As Long
    
    Set wb = ActiveWorkbook
    
    ' Create Sheets (Delete old ones if they exist)
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Sheets("Loan Info Inputs").Delete
    wb.Sheets("Bills").Delete
    wb.Sheets("Side-by-Side").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsInputs = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsInputs.Name = "Loan Info Inputs"
    Set wsBills = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsBills.Name = "Bills"
    Set wsSide = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsSide.Name = "Side-by-Side"
    
    ' ==========================================
    ' SHEET 1: LOAN INFO INPUTS
    ' ==========================================
    With wsInputs
        .Range("A1").Value = "Global Inputs": .Range("D1").Value = "Notes"
        .Range("A1:D1").Font.Bold = True
        .Range("A2").Value = "Loan Balance": .Range("B2").Value = 490000
        .Range("A3").Value = "Mortgage Interest Rate": .Range("B3").Value = 0.04
        .Range("A4").Value = "Mortgage Term (Years)": .Range("B4").Value = 30
        .Range("A5").Value = "Start Date of Loan": .Range("B5").Value = DateSerial(2024, 1, 1)
        .Range("A6").Value = "Mortgage Payment Day": .Range("B6").Value = 1
        .Range("A8").Value = "HELOC Strategy Inputs": .Range("A8").Font.Bold = True
        .Range("A9").Value = "HELOC Interest Rate": .Range("B9").Value = 0.07
        .Range("A10").Value = "Weekly HELOC Payment": .Range("B10").Value = 2000
        .Range("A11").Value = "Bi-Weekly HELOC Payment": .Range("B11").Value = 4000
        .Range("A12").Value = "HELOC First Pmt Date (Weekly)": .Range("B12").Value = DateSerial(2024, 1, 5)
        .Range("A13").Value = "HELOC First Pmt Date (Bi-Weekly)": .Range("B13").Value = DateSerial(2024, 1, 12)
        .Range("A15").Value = "Calculated Fields": .Range("A15").Font.Bold = True
        .Range("A16").Value = "Mortgage Monthly Payment"
        .Range("B16").Formula = "=PMT(B3/12, B4*12, -B2)"
        .Range("A19").Value = "Other Details": .Range("A19").Font.Bold = True
        .Range("A20").Value = "Credit Card Charge for Draw Payments"
        .Range("B20").Value = "y"
        .Range("D20").Value = "Enter a y if you want all possible bills to be paid with Credit card..."
        .Range("A1:D18").Borders.LineStyle = xlContinuous
        .Columns("A:D").AutoFit
        .Range("B2, B10, B11, B16").Style = "Currency"
        .Range("B3, B9").NumberFormat = "0.00%"
    End With

    ' ==========================================
    ' SHEET 2: BILLS
    ' ==========================================
    With wsBills
        ' Header Text
        .Range("A1").Value = "Bill Description"
        .Range("B1").Value = "Bill Amount"
        .Range("C1").Value = "Week/Month/Annual" & vbLf & "(Enter: w,m,a)"
        .Range("D1").Value = "Due Date Month" & vbLf & "(Annual Bills Only)"
        .Range("E1").Value = "Due Date Day" & vbLf & "(Annual/Monthly Bills Only)"
        .Range("F1").Value = "Can be paid on CC" & vbLf & "(Enter: y,n)"
        .Range("H1").Value = "Current Gallon Cost"
        .Range("I1").Value = "Fuel Tank Gallons"
        .Range("J1").Value = "Tanks Per Week"
        .Range("K1").Value = "Notes/Comments"
        .Range("A1:K1").Font.Bold = True
        .Range("A1:K1").Borders.LineStyle = xlContinuous
        
        ' Background Update: Clear F1:J1, Black for F2:J2
        .Range("F1:J1").Interior.ColorIndex = xlNone
        .Range("F2:J2").Interior.Color = vbBlack

        ' Column Alignment
        .Columns("C:J").HorizontalAlignment = xlCenter
        
        ' Row 2 logic
        .Range("A2").Value = "Credit Card Payment Draw"
        .Range("B2").Formula = "=SUMIFS(B10:B100, F10:F100, ""y"", C10:C100, ""<>a"")"
        .Range("C2").Value = "m"
        .Range("E2").Interior.Color = RGB(220, 255, 220)
        .Range("A2:B2").Interior.Color = vbRed
        .Range("A2:B2").Font.Color = vbWhite
        .Range("A2:E2").Borders.LineStyle = xlContinuous
        
        ' Row 3 instruction area
        .Range("A3").Value = "For Bill amount, it will calculate based on your entries of data into Columns H, I, J ONLY"
        .Range("A3:K3").Interior.Color = RGB(200, 200, 200)
        
        ' Vehicles section
        .Range("A4:A7").Value = Application.Transpose(Array("vehicle 1", "vehicle 2", "vehicle 3", "vehicle 4"))
        .Range("B4:B7").Formula = "=H4*I4*J4"
        .Range("C4:C7").Value = "w"
        
        ' Borders for row 4
        .Range("A4:J4").Borders.LineStyle = xlContinuous
        
        ' Blackout and Light Green highlights
        .Range("D4:G7").Interior.Color = vbBlack
        .Range("F4:F7").Interior.ColorIndex = xlNone
        .Range("H4:J7").Interior.Color = RGB(220, 255, 220)
        
        .Range("H4:J7").Borders.LineStyle = xlContinuous
        .Range("A4:C7").Borders.LineStyle = xlContinuous
        
        ' Row 8 Yellow background
        .Range("A8").Value = "If more vehicles, right click row, ""Insert"""
        .Range("A8").Font.Italic = True
        .Range("A8:K8").Interior.Color = vbYellow
        
        ' Row 9 Separator
        .Range("A9").Value = "Start Entering Bills below"
        .Range("A9:K9").Interior.Color = RGB(200, 200, 200)
        
        ' Bill Entry Area (Light Green)
        .Range("C10:F60").Interior.Color = RGB(220, 255, 220)
        .Range("A10:K60").Borders.LineStyle = xlContinuous
        
        ' Patterns
        With .Range("H10:J60").Interior
            .Pattern = xlPatternUpward
            .PatternColorIndex = xlAutomatic
        End With
        
        ' Cell K2 Note
        .Range("K2").Value = "Charge all your bills on your credit card. Earn points/cash back for your all charges. Key, pay off Credit Card each month prior to being charged interest on your statement."
        
        ' Column Formatting
        .Columns("I:J").NumberFormat = "General"
        .Columns("B").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        .Range("H4:H7").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        
        ' AutoSize
        .Range("A1:F1").WrapText = True
        .Columns("A:K").AutoFit
        .Columns("K").ColumnWidth = 50
    End With

    ' ==========================================
    ' SHEET 3: SIDE-BY-SIDE (RESTORED BASELINE)
    ' ==========================================
    With wsSide
        lastRow = 11000
        .Range("B1").Value = "Date Loan Paid Off"
        .Range("B2").Value = "Total Months to Payoff"
        .Range("B3").Formula = "=IFERROR(INDEX(A8:A" & lastRow & ", MATCH(0, F8:F" & lastRow & ", 0)), ""Not Paid"")"
        .Range("B2").Formula = "=IF(ISNUMBER(B3), DATEDIF('Loan Info Inputs'!$B$5, B3, ""m""), 0)"
        .Range("K1").Value = "Date Loan Paid Off"
        .Range("K2").Value = "Total Months to Payoff"
        .Range("K3").Formula = "=IFERROR(INDEX(A8:A" & lastRow & ", MATCH(0, K8:K" & lastRow & ", 0)), ""Not Paid"")"
        .Range("K2").Formula = "=IF(ISNUMBER(K3), DATEDIF('Loan Info Inputs'!$B$5, K3, ""m""), 0)"
        .Range("L1").Value = "Net HELOC Principal Paid"
        .Range("L2").Formula = "=(SUM(L8:L" & lastRow & ") - P2)"
        
        Dim headers As Variant
        headers = Array("Date", "Mortgage Starting Balance", "Mortgage Principle Paid", "Mortgage Int. Paid", _
                        "Mortgage Total Payment", "Mortgage End Balance", "Cum Mortgage Total Paid", _
                        "Cum Mortgage Principle Paid", "Cum Mortgage Interest Paid", " ", _
                        "HELOC Start Balance", "HELOC Principle Paid", "HELOC Daily Interest Amount", _
                        "HELOC Interest Accrued", "HELOC Payment", "HELOC Draw", "HELOC Draw Description", _
                        "Cum HELOC Payment Total Paid", "Cum HELOC Principle Total Paid", _
                        "Cum HELOC Interest Total Paid", "Cum Draw Total")
        
        .Range("A6").Resize(1, UBound(headers) + 1).Value = headers
        .Range("A6:U6").Font.Bold = True

        Dim colArr As Variant, col As Variant
        colArr = Array("C", "D", "E", "M", "O", "P", "Q")
        For Each col In colArr
            .Range(col & "1").Formula = "=" & col & "6"
            .Range(col & "2").Formula = "=SUM(" & col & "8:" & col & lastRow & ")"
            .Range(col & "1:" & col & "2").Font.Bold = True
            .Range(col & "2").Style = "Currency"
        Next col

        .Range("A7").Formula = "='Loan Info Inputs'!B5-1"
        .Range("F7").Formula = "='Loan Info Inputs'!B2"
        .Range("K7").Formula = "='Loan Info Inputs'!B2"
        .Range("G7, H7, I7, N7, R7, S7, T7, U7").Value = 0
        
        .Range("A8").Formula = "=A7+1"
        .Range("B8").Formula = "=F7"
        .Range("D8").Formula = "=IF(DAY(A8)='Loan Info Inputs'!$B$6, B8*('Loan Info Inputs'!$B$3/12), 0)"
        .Range("E8").Formula = "=IF(B8<=0, 0, IF(DAY(A8)='Loan Info Inputs'!$B$6, MIN('Loan Info Inputs'!$B$16, B8 + D8), 0))"
        .Range("C8").Formula = "=IF(E8>0, E8-D8, 0)"
        .Range("F8").Formula = "=MAX(0, B8 - C8)"
        .Range("G8").Formula = "=G7 + E8": .Range("H8").Formula = "=H7 + C8": .Range("I8").Formula = "=I7 + D8"

        .Range("P8").Formula = _
            "=IF(K7 <= 0, 0, IF(UPPER('Loan Info Inputs'!$B$20)=""Y""," & _
                "(SUMIFS(Bills!$B$4:$B$100, Bills!$C$4:$C$100, ""a"", Bills!$D$4:$D$100, MONTH(A8), Bills!$E$4:$E$100, DAY(A8), Bills!$F$4:$F$100, ""n"") + " & _
                "SUMIFS(Bills!$B$4:$B$100, Bills!$C$4:$C$100, ""m"", Bills!$E$4:$E$100, DAY(A8), Bills!$F$4:$F$100, ""n"") + " & _
                "IF(WEEKDAY(A8)=7, SUMIFS(Bills!$B$4:$B$100, Bills!$C$4:$C$100, ""w"", Bills!$F$4:$F$100, ""n""), 0) + " & _
                "IF(AND(DAY(A8)=Bills!$E$2, MONTH(A8)=Bills!$D$2, Bills!$C$2=""a""), Bills!$B$2, " & _
                "IF(AND(DAY(A8)=Bills!$E$2, Bills!$C$2=""m""), Bills!$B$2, 0))), " & _
                "(SUMIFS(Bills!$B$4:$B$100, Bills!$C$4:$C$100, ""a"", Bills!$D$4:$D$100, MONTH(A8), Bills!$E$4:$E$100, DAY(A8)) + " & _
                "SUMIFS(Bills!$B$4:$B$100, Bills!$C$4:$C$100, ""m"", Bills!$E$4:$E$100, DAY(A8)) + " & _
                "IF(WEEKDAY(A8)=7, SUMIFS(Bills!$B$4:$B$100, Bills!$C$4:$C$100, ""w""), 0))))"

        .Range("Q8").Formula2 = _
            "=IF(P8=0, """", TEXTJOIN(CHAR(10), TRUE, " & _
            "IFERROR(FILTER(Bills!$A$4:$A$100, ( (Bills!$E$4:$E$100=DAY(A8)) * ( (Bills!$C$4:$C$100=""m"") + ((Bills!$C$4:$C$100=""a"")*(Bills!$D$4:$D$100=MONTH(A8))) ) * IF('Loan Info Inputs'!$B$20=""y"", Bills!$F$4:$F$100=""n"", 1) ) + ( (Bills!$C$4:$C$100=""w"")*(WEEKDAY(A8)=7)*IF('Loan Info Inputs'!$B$20=""y"", Bills!$F$4:$F$100=""n"", 1) ) ), """"), " & _
            "IF(AND('Loan Info Inputs'!$B$20=""y"", DAY(A8)=Bills!$E$2, OR(Bills!$C$2=""m"", AND(Bills!$C$2=""a"", Bills!$D$2=MONTH(A8)))), Bills!$A$2, """") ))"

        .Range("M8").Formula = "=K7 * ('Loan Info Inputs'!$B$9 / 365)"
        .Range("O8").Formula = _
            "=IF(K7+N7+M8<=0, 0, " & _
            "IF(OR(MOD(A8-'Loan Info Inputs'!$B$12, 7)=0, MOD(A8-'Loan Info Inputs'!$B$13, 14)=0), " & _
            "MIN(K7 + N7 + M8, " & _
            "IF(MOD(A8-'Loan Info Inputs'!$B$12, 7)=0, 'Loan Info Inputs'!$B$10, 0) + " & _
            "IF(MOD(A8-'Loan Info Inputs'!$B$13, 14)=0, 'Loan Info Inputs'!$B$11, 0)), 0))"

        .Range("N8").Formula = "=MAX(0, N7 + M8 - O8)"
        .Range("L8").Formula = "=MAX(0, O8 - (N7 + M8))"
        .Range("K8").Formula = "=MAX(0, K7 + P8 - L8)"
        .Range("R8").Formula = "=R7 + O8": .Range("S8").Formula = "=S7 + L8"
        .Range("T8").Formula = "=T7 + MIN(O8, N7 + M8)": .Range("U8").Formula = "=U7 + P8"
        
        .Range("A8:U8").AutoFill Destination:=.Range("A8:U" & lastRow)
        .Range("A:A").NumberFormat = "mm/dd/yyyy"
        .Range("B:F, G:I, K:P, R:U").Style = "Currency"
        .Range("Q:Q").WrapText = True: .Range("Q:Q").VerticalAlignment = xlVAlignCenter
        .Columns("A:U").AutoFit: .Columns("Q").ColumnWidth = 35
        .Activate: ActiveWindow.SplitColumn = 1: ActiveWindow.SplitRow = 6: ActiveWindow.FreezePanes = True
    End With
    
    MsgBox "Bills update: F1:J1 clear, F2:J2 black!", vbInformation
End Sub

