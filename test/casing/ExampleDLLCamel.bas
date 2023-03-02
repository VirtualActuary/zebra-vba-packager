Attribute VB_Name = "miscCollection"
Option Explicit

#If VBA7 Then
    #If Win64 Then
        Declare PtrSafe Function GetKeyState Lib "User32" (ByVal vKey As Integer) As Integer
    #Else
        Declare Function GetKeyState Lib "User32" (ByVal vKey As Integer) As Integer
    #End If
#Else
    Declare Function GetKeyState Lib "User32" (ByVal vKey As Integer) As Integer
#End If


Public Function min(ByVal col As collection) As Variant
    ' Returns the minimum value from the input Collection.
    '
    ' Args:
    '   col: Collection with numerical values.
    
    ' Returns:
    '   The minimum value in the collection.
    
    If col Is Nothing Then
        err.raise number:=91, _
              description:="Collection input can't be empty"
    End If
    
    Dim entry As Variant
    min = col(1)
    
    For Each entry In col
        If entry < min Then
            min = entry
        End If
    Next entry  
End Function
