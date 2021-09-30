Attribute VB_Name = "Module3"
Sub Format_Allocation_By_Store()
Attribute Format_Allocation_By_Store.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' Format_Allocation_By_Store Macro
'
' Keyboard Shortcut: Ctrl+Shift+F
'
' Use this macro to break out allocations by store numbers.

Set all_wks = ThisWorkbook.Sheets("Allocations")
Set sk_wks = ThisWorkbook.Sheets("EXAMPLE check")


'Delete Old Data from C3:E43
all_wks.Activate
Range("C3", "F43").Clear

'Find & sum the store numbers in sk_wks
'Dim Alloc_Rng As Range
Dim Search_Rng As Range
Dim Alloc_Rng As Range

Dim rng As Range
Dim Store_Found As Range
Dim Alloc_Found As Range

Dim Store_Sum As Integer

Set Alloc_Rng = all_wks.Range("B3:B43")
Set Search_Rng = sk_wks.Range("B2:B100")


For Each Alloc_Num In Alloc_Rng
    Alloc_Num = Alloc_Num.Value
    sk_wks.Activate
    
    Set Store_Found = Search_Rng.Find(What:=Alloc_Num, LookAt:=xlWhole)
    If Not Store_Found Is Nothing Then
        First_Found = Store_Found.Address 'First Occurance
    Else
    End If
        
    Set rng = Store_Found
        
    Do Until Store_Found Is Nothing
        Set Store_Found = Search_Rng.FindNext(After:=Store_Found)
        Set rng = Union(rng, Store_Found)
        If Store_Found.Address = First_Found Then Exit Do
    Loop
    
    If Not rng Is Nothing Then
        rng.Offset(0, 2).Select
        Store_Sum = Application.WorksheetFunction.Sum(Selection)
    Else
        Store_Sum = 0
    End If
    
    all_wks.Activate
    
    If Store_Sum = 0 Then
        Store_Sum = 0
        ElseIf Store_Sum > 0 Then
            Set Alloc_Found = Alloc_Rng.Find(What:=Alloc_Num, LookAt:=xlWhole)
            all_wks.Range(Alloc_Found.Address).Offset(0, 1).Value = Store_Sum
        Else
            Set Alloc_Found = Alloc_Rng.Find(What:=Alloc_Num, LookAt:=xlWhole)
            all_wks.Range(Alloc_Found.Address).Offset(0, 2).Value = Store_Sum
    End If

    
Next Alloc_Num

End Sub

Sub Allocate_Negative_Numbers()
Attribute Allocate_Negative_Numbers.VB_ProcData.VB_Invoke_Func = "N\n14"
'
' Format_Allocate_Negative_Numbers Macro
'
' Keyboard Shortcut: Ctrl+Shift+N
'
' Use this macro to break allocate the negative amounts to stores with positive amounts.
Dim Alloc_Rng As Range
Dim neg_alloc_rng As Range
Dim slct As Range

Set all_wks = ThisWorkbook.Sheets("Allocations")
Set sk_wks = ThisWorkbook.Sheets("EXAMPLE check")

Set Alloc_Rng = all_wks.Range("B3:B43")

Set neg_alloc_rng = Alloc_Rng.Offset(0, 2)

For Each neg_num In neg_alloc_rng
    If neg_num < 0 Then
        'find positive number to adjust
         pos_add = Alloc_Rng.Offset(0, 1).Select
         pos_add = Range(Selection(1).Address).Select
         pos_num = Selection(1).Value
         Do Until pos_num > neg_num * -1 And IsEmpty(Range(Selection.Address).Offset(0, 2).Value)
             pos_add = Range(Selection.Address).Offset(1, 0).Select
             pos_num = Selection.Value
            If IsEmpty(pos_num) Then
                pos_add = Range(Selection.Address).Offset(1, 0).Select
                pos_num = Selection.Value
            ElseIf pos_num < neg_num * -1 Then
                pos_add = Range(Selection.Address).Offset(1, 0).Select
                pos_num = Selection.Value
            End If
         Loop
         
         pos_add = Selection.Address
         neg_add = neg_num.Address
         Range(pos_add).Offset(0, 2).Value = Application.WorksheetFunction.Sum(pos_num + neg_num)
         Range(pos_add).Offset(0, 2).Interior.ColorIndex = 35
    End If
    
Next neg_num

End Sub


