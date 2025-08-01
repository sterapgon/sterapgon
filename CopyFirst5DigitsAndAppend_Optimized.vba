Sub CopyFirst5DigitsAndAppend000_Optimized()
    ' --- ตั้งค่าเริ่มต้น ---
    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = InputBox("กรุณาใส่ชื่อชีทที่ต้องการประมวลผล", "เลือกชีท", "Sheet1")
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo Cleanup
    If ws Is Nothing Then
        MsgBox "ไม่พบชีทชื่อ '" & sheetName & "'", vbCritical
        Exit Sub
    End If

    ' ปิดการทำงานเบื้องหลังของ Excel เพื่อเพิ่มความเร็วสูงสุด
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With

    ' --- เตรียมข้อมูล ---
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "ไม่พบข้อมูลในคอลัมน์ F", vbInformation
        GoTo Cleanup
    End If

    Debug.Print "เริ่มประมวลผลข้อมูลจำนวน " & (lastRow - 1) & " แถว"

    ' --- ประมวลผลบน Array ---
    Dim dataF As Variant, dataT As Variant
    Dim i As Long
    Dim valueInF As String

    dataF = ws.Range("F2:F" & lastRow).Value
    ReDim dataT(1 To UBound(dataF, 1), 1 To 1)

    For i = 1 To UBound(dataF, 1)
        valueInF = Trim(CStr(dataF(i, 1)))

        If Len(valueInF) = 0 Then
            dataT(i, 1) = "ว่าง"
        ElseIf Len(valueInF) >= 5 Then
            dataT(i, 1) = Left(valueInF, 5) & "000"
        Else
            dataT(i, 1) = "ข้อมูลไม่ถูกต้อง"
        End If

        ' แสดงสถานะทุก ๆ 50,000 แถว
        If i Mod 50000 = 0 Then
            Debug.Print Format(Now, "hh:mm:ss") & " - ประมวลผลถึงแถว: " & i
        End If
    Next i

    ' --- เขียนข้อมูลกลับลงชีต ---
    ws.Range("T2").Resize(UBound(dataT, 1), 1).Value = dataT

    Debug.Print "------ เสร็จสิ้นการประมวลผล ------"
    MsgBox "ประมวลผลข้อมูลเรียบร้อยแล้ว จำนวน " & (lastRow - 1) & " แถว", vbInformation

Cleanup:
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With

    If Err.Number <> 0 Then
        MsgBox "เกิดข้อผิดพลาด: " & Err.Description, vbCritical
    End If
End Sub
