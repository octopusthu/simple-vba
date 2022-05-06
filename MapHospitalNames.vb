
' 医院数据
Private Const hospitalSheetName As String = "医院列表"
Private Const scanFromRow As Integer = 2
Private Const scanToRow As Integer = 1215
Private Const scanFromColumn As Integer = 4
Private Const scanToColumn As Integer = 8

' 统计数据
Private Const dataSheetName As String = "统计"
Private Const applicantColumn As Integer = 1
Private Const resultColumn As Integer = 3
Private Const applicantFirstRow As Integer = 2
Private Const applicantLastRow As Long = 141

Sub Main()
 MapApplicantsToStandardHospitalNames
End Sub

' 读取“统计”页“申请人”列，比对“医院列表”页的字典数据，将其转换为医院标准名称，填入“申请人（医院）”列。
Private Sub MapApplicantsToStandardHospitalNames()

Dim standardHospitalName As String
For scanningRow = scanFromRow To scanToRow
 standardHospitalName = ThisWorkbook.Worksheets(hospitalSheetName).Cells(scanningRow, scanFromColumn).Value

 If IsEmptyString(standardHospitalName) = False Then

  Dim keyword As String
  For scanningColumn = scanFromColumn To scanToColumn
   keyword = ThisWorkbook.Worksheets(hospitalSheetName).Cells(scanningRow, scanningColumn).Value
   If IsEmptyString(keyword) = False Then
    MapHospital Trim(keyword), Trim(standardHospitalName)
   End If
  Next scanningColumn

 End If

Next scanningRow

End Sub

' 读取医院标准名称及其别名
Private Sub MapHospital(keyword As String, standardHospitalName As String)
 Debug.Print "Mapping " & keyword & " to " & standardHospitalName

 Dim applicant As String
 For processingRow = applicantFirstRow To applicantLastRow
  applicant = ThisWorkbook.Worksheets(dataSheetName).Cells(processingRow, applicantColumn).Value
  If ContainsKeyword(applicant, keyword) = True Then
   ThisWorkbook.Worksheets(dataSheetName).Cells(processingRow, resultColumn).Value = standardHospitalName
  End If
 Next processingRow

End Sub

Private Function IsEmptyString(str As String) As Boolean
 If Trim(str) = "" Then
  IsEmptyString = True
 Else
  IsEmptyString = False
 End If
End Function

Private Function ContainsKeyword(str As String, keyword As String) As Boolean
 Dim pos As Integer
 pos = InStr(str, keyword)
 If pos = 0 Then
  ContainsKeyword = False
 Else
  ContainsKeyword = True
 End If
End Function
