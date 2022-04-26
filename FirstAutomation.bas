Dim StartingYear As Integer, EndingYear As Integer, YearRange As Integer
Dim MediaNameArray As Variant, MediaNameCell As Range, MediaAmount As Integer
Dim TargetBrandsArray As Variant, TargetBrandCell As Range, TargetBrandsAmount As Integer

Sub 기본세팅()
    Set StartingYear = , EndingYear = , TargetBrandCell = , TargetBrandsAmount = 0

    If
        IsNull(StartingYear) || IsNull(EndingYear) || IsNull(TargetBrandCell)
        then
            MsgBox "자료가 충분히 입력되지 않았습니다."
            End Sub
        Else
            Do While TargetBrandCell.Value = !Null
                TargetBrandsArray(TargetBrandsAmount) = TargetBrandCell.Value
                TargetBrandsAmount = 1
            Loop
                Dim BasicSetupCompletionMsg As String
                Set BasicSetupCompletionMsg = "시작년도 " & CStr(StartingYear) & ", 종료년도 " & CStr(EndingYear) & ", 브랜드 " & CStr(TargetBrandsAmount) & "종 입력 완료되었습니다."
                MsgBox BasicSetupCompletionMsg
                End Sub

Sub 자료입력()
    
    If
        IsNull(StartingYear) || IsNull(EndingYear) || IsNull(TargetBrandCell)
        then
            MsgBox "기본자료를 먼저 입력하세요."
            Sheets("기본자료").Activate
            End Sub
        Else
            Dim 03tYear as Integer, MediumName as String, DataLastRow as Integer, DataLastColumn as Integer, InputData as ListObjects, InputDataRange as Range
            MediumName = Range("C2").Value
            InputYear = Range("E2").Value

            LastRow = Cells.Find(What:="*", After:=Range("B5"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
            LastColumn = Cells.Find(What:="*", After:=Range("B5"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
            Range("B5").Resize(LastRow, LastColumn).Select

            'Set InputData = ActiveSheets.ListObjects.Add
            '    (
            '        SourceType:= xlSrcRange, _ ,
            '    )
            'Here Goes "Sub 표작성()"

End Sub

Sub 표작성()
    Dim TableLastRow As Integer, TableLastColumn As Integer

    Sheets("BackData").Activate

End Sub
