Attribute VB_Name = "Module5"
Option Explicit
Option Base 1

'��2�����z���1�����ڂ𑝂₷���W���[��
Public Function RedimPreserveArray(ByVal arr As Variant, ByVal sLen As Long)
    Dim temp() As Variant
     
    temp = WorksheetFunction.Transpose(arr)
    ReDim Preserve temp(UBound(temp, 1), sLen)
     
    RedimPreserveArray = WorksheetFunction.Transpose(temp)
End Function

