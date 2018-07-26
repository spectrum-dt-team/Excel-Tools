# Excel-Tools
  found from (https://www.extendoffice.com/documents/excel/3790-excel-find-position-of-first-letter-in-string.html) 

To look up the first text value starts in a cell- this will need to be used
  This function will need to be entered into Visual Basics

Function FirstNonDigit(xStr As String) As Long
'Updateby20160711
    Dim xChar As Integer
    Dim xPos As Integer
    Dim I As Integer
    Application.Volatile
    For I = 1 To Len(xStr)
        xChar = Asc(Mid(xStr, I, 1))
        If xChar <= 47 Or _
           xChar >= 58 Then
            xPos = I
            Exit For
        End If
    Next
    FirstNonDigit = xPos
End Function
