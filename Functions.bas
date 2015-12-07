Option Explicit

Private Function RangeContent() As Range
Dim x As Double
Dim y As Double

'   identify Range
With ThisWorkbook.Worksheets("Tabelle1")
    x = .Range("A2").End(xlToRight).Column
    y = .Range("A2").End(xlDown).Row
End With

'   span and output Range
With ThisWorkbook.Sheets("Tabelle1")
    Set RangeContent = Range(.Cells(2, 1), .Cells(y, x))
End With

End Function

Private Function UniqueItems(Ar As Variant) As Variant
Dim UnAr() As Variant   'UniqueArray
Dim UnC As Double       'UniqueCounter
Dim FMatch As Boolean   'FoundMatch
Dim FMItem As Variant   'FoundMatchItem

UnC = -1

'   every item
For Each FMItem In Ar
    FMatch = False
    
    If UnC > -1 Then 'the first item with UnC -1 has to be added!
        If IsInArray(FMItem, UnAr) = True Then
            FMatch = True
        End If
    End If
    
    '   if the item is not in UnAr
    If FMatch = False Then
        UnC = UnC + 1
        '   ReDim UnAr and ad item
        ReDim Preserve UnAr(UnC)
        UnAr(UnC) = FMItem
    End If
Next FMItem

'   output unique array
UniqueItems = UnAr
End Function

Private Function IsInArray(FMItem As Variant, Ar As Variant) As Boolean
    IsInArray = (UBound(Filter(Ar, FMItem)) > -1)
End Function

Function CompArray(Ar As Variant, CompAr As Variant) As Variant
Dim UnAr() As Variant
Dim UnC As Double
Dim FMatch As Boolean
Dim FMElement As Variant

UnC = -1

For Each FMElement In Ar

    FMatch = bIsInArray(FMElement, CompAr)
    
    If FMatch = False Then
        UnC = UnC + 1
        ReDim Preserve UnAr(UnC)
        UnAr(UnC) = FMElement
    End If
Next FMElement

If UnC > -1 Then
    vCompArray = UnAr
End If
End Function
