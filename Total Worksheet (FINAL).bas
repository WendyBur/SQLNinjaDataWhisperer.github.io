Attribute VB_Name = "Module1"
Sub Subtotal_Worksheet()
'
' Subtotal_Worksheet Macro
'
' Keyboard Shortcut: Ctrl+Shift+W
'
' Use this macro to check data entry by subtotaling each page of check


Worksheets("EXAMPLE check").Activate

'Calculate Sum of Range 1
Static Sum1 As Long
Dim SumBox1 As Long

Sum1 = ActiveSheet.Range("D3", Range("D3").End(xlDown)).Select

If ActiveWindow.RangeSelection.count > 1 Then
      SumTxt = ActiveWindow.RangeSelection.AddressLocal
    Else
      SumTxt = ActiveSheet.UsedRange.AddressLocal
    End If

SumBox1 = ActiveSheet.Range("D3").End(xlDown).Offset(0, 2).Select

ActiveCell.Formula = "=SUM(" & SumTxt & ")"


'Calculate Sum of Range 2
Dim R2Start As Integer
Static Sum2 As Long
Dim mycell As Range


R2Start = ActiveSheet.Range("D3").End(xlDown).Offset(2, 0).Select
Set mycell = ActiveCell
Sum2 = ActiveSheet.Range("D3", Range(ActiveCell.Address).End(xlDown)).Select
If ActiveWindow.RangeSelection.count > 1 Then
      SumTxt2 = ActiveWindow.RangeSelection.AddressLocal
    Else
      SumTxt2 = ActiveSheet.UsedRange.AddressLocal
    End If

SumBox2 = mycell.End(xlDown).Offset(0, 2).Select
ActiveCell.Formula = "=SUM(" & SumTxt2 & ")"


'Calculate Sum of Range 3
Dim R3Start As Range
Static Sum3 As Long

ActiveCell.Offset(2, -2).Select
Set R3Start = ActiveCell

Sum3 = ActiveSheet.Range("D3", Range(ActiveCell.Address).End(xlDown)).Select

If ActiveWindow.RangeSelection.count > 1 Then
      SumTxt3 = ActiveWindow.RangeSelection.AddressLocal
    Else
      SumTxt3 = ActiveSheet.UsedRange.AddressLocal
    End If

SumBox3 = R3Start.End(xlDown).Offset(0, 2).Select
ActiveCell.Formula = "=SUM(" & SumTxt3 & ")"
   
   
End Sub

