Sub ListAllCombinations()
'Updateby Extendoffice
Dim xDRg1, xDRg2, xDRg3 As Range
Dim xRg As Range
Dim xStr As String
Dim xFN1, xFN2, xFN3 As Integer
Dim xSV1, xSV2, xSV3 As String
Set xDRg1 = Range("A2:A5") 'First column data
Set xDRg2 = Range("B2:B4") 'Second column data
Set xDRg3 = Range("C2:C4") 'Third column data
xStr = "-" 'Separator
Set xRg = Range("E2") 'Output cell
For xFN1 = 1 To xDRg1.Count
 xSV1 = xDRg1.Item(xFN1).Text
 For xFN2 = 1 To xDRg2.Count
  xSV2 = xDRg2.Item(xFN2).Text
  For xFN3 = 1 To xDRg3.Count
   xSV3 = xDRg3.Item(xFN3).Text
   xRg.Value = xSV1 & xStr & xSV2 & xStr & xSV3
   Set xRg = xRg.Offset(1, 0)
  Next
 Next
Next
End Sub
