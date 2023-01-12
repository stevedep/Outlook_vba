Attribute VB_Name = "Mod_CategoryManagement"


Function SortDict(Dict As Scripting.Dictionary) As Variant
    'Set a reference to Microsoft Scripting Runtime by using
    'Tools > References in the Visual Basic Editor (Alt+F11)

    'Declare the variables
    Dim Arr() As Variant
    Dim Temp1 As Variant
    Dim Temp2 As Variant
    Dim Txt As String
    Dim i As Long
    Dim j As Long
    
    'Allocate storage space for the dynamic array
    ReDim Arr(0 To Dict.Count - 1, 0 To 1)
    
    'Fill the array with the keys and items from the Dictionary
    For i = 0 To Dict.Count - 1
        Arr(i, 0) = Dict.Keys(i)
        Arr(i, 1) = Dict.Items(i)
    Next i
    
    'Sort the array using the bubble sort method
    For i = LBound(Arr, 1) To UBound(Arr, 1) - 1
        For j = i + 1 To UBound(Arr, 1)
            If Arr(i, 1) < Arr(j, 1) Then
                Temp1 = Arr(j, 0)
                Temp2 = Arr(j, 1)
                Arr(j, 0) = Arr(i, 0)
                Arr(j, 1) = Arr(i, 1)
                Arr(i, 0) = Temp1
                Arr(i, 1) = Temp2
            End If
        Next j
    Next i
    
    'Clear the Dictionary
    Dict.RemoveAll
    
SortDict = Arr
End Function

