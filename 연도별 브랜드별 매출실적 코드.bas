Sub 자료입력()
    Sheets("자료입력").Activate
    
    Dim Year as Integer, MediumName as String, DataLastRow as Integer, DataLastColumn as Integer, InputData as ListObjects, InputDataRange as Range
    MediumName = Range("C2").Value
    Year = Range("E2").Value

    LastRow = Cells.Find(What:="*", After:=Range("B5"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LastColumn = Cells.Find(What:="*", After:=Range("B5"), SearchOrder:xlByColumn, SearchDirection:=xlPrevious).Column
    Range("B5").Resize(LastRow, LastColumn).Select

    Set InputData = ActiveSheets.ListObjects.Add
        (
            SourceType:= xlSrcRange, _ ,            
        )

    Sheets("BackData").Activate
    // Here Goes "Sub 표작성()"

End Sub

Sub 표작성()
    Dim TableLastRow as Integer, TableLastColumn as Integer

    Sheets("BackData").Activate

Sub 브랜드입력() // 전면 재수정 필요
    dim TargetBrandsArray as Variant, TargetBrandCell as Range("G5"), TargetBrandsAmount as Integer

    if
        TargetBrandCell.Value = null && TargetBrandsAmount <= 0
        then
            MsgBox "대상 브랜드가 없습니다."
            End Sub
        Else
            MsgBox "대상 브랜드를 입력합니다."
            Set TargetBrandsAmount 0

            Do While (TargetBrandCell.Value =! null)
                TargetBrandsArray(TargetBrandsAmount) = TargetBrandCell.Value
                TargetBrandsAmount =+ 1
            Loop

            Dim TargetBrandsCompleteString as String
            Set TargetBrandsCompleteString = "대상 브랜드 " & CStr(TargetBrandsAmount) & "종의 입력이 완료되었습니다."

            MsgBox TargetBrandsCompleteString
End Sub