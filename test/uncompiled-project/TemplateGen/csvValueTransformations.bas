Attribute VB_Name = "csvValueTransformations"
Option Explicit

Function applyCsvTransformations(tableArr As Variant) As Variant
    
    Dim i As Long, j As Long
    For i = LBound(tableArr, 1) To UBound(tableArr, 1)
        For j = LBound(tableArr, 2) To UBound(tableArr, 2)
            tableArr(i, j) = applyTransformation(tableArr(i, j))
        Next j
    Next i
    
    applyCsvTransformations = tableArr
    
End Function


Private Function applyTransformation(val As Variant) As Variant

    If IsError(val) Then ' set all error values to an empty string
        applyTransformation = vbNullString
    ElseIf IsNumeric(val) Then ' force numeric values to use . as decimal separator
        applyTransformation = decStr(val)
    ElseIf IsDate(val) Then ' format dates as strings to avoid some user's stupid default date settings
        applyTransformation = dateToString(CDate(val))
    Else ' do nothing
        applyTransformation = val
    End If

End Function

Private Function dateToString(d As Date) As String
    If d = Int(d) Then ' no hours, etc:
        dateToString = Format(d, "yyyy-mm-dd")
    Else ' add hours and seconds - VBA can't keep more details in any case...
        dateToString = Format(d, "yyyy-mm-dd hh:mm:ss")
    End If
End Function

Private Function decStr(x As Variant) As String
     decStr = CStr(x)

     'Frikin ridiculous loops for VBA
     If IsNumeric(x) Then
        decStr = Replace(decStr, Format(0, "."), ".")
        ' Format(0, ".") gives the system decimal separator
     End If

End Function
