Attribute VB_Name = "ComputeMd5Hash"
Option Explicit

Public Function ComputeMd5(FilePath As String) As String
    ComputeMd5 = LCase(CreateMD5HashFile(FilePath))
End Function
