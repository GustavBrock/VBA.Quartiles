Attribute VB_Name = "QuartileDemo"
Option Compare Database
Option Explicit

Public Function ListFirstQuartile()

    Dim Method  As ApQuartileMethod
    Dim Count   As Long
    
    For Count = 40 To 70 Step 10
        Debug.Print , Count;
    Next
    Debug.Print
    For Method = ApQuartileMethod.[_First] To ApQuartileMethod.[_Last]
        Debug.Print Method, ;
        For Count = 40 To 70 Step 10
            Debug.Print Format(Quartile("Data", "Observation", "Step = 10 And Data <= " & Count & "", apFirst, Method), "0.00"), ;
        Next
        Debug.Print
    Next
    Debug.Print
    
    For Count = 100 To 95 Step -1
        Debug.Print , Count;
    Next
    Debug.Print
    For Method = ApQuartileMethod.[_First] To ApQuartileMethod.[_Last]
        Debug.Print Method, ;
        For Count = 100 To 95 Step -1
            Debug.Print Format(Quartile("Data", "Observation", "Step = 1 And Data <= " & Count & "", apFirst, Method), "0.00"), ;
        Next
        Debug.Print
    Next

End Function

Public Function ListSecondQuartile()

    Dim Count   As Long
    
    For Count = 40 To 70 Step 10
        Debug.Print , Count;
    Next
    Debug.Print
    Debug.Print 1, ;
    For Count = 40 To 70 Step 10
        Debug.Print Format(DMedian("Data", "Observation", "Step = 10 And Data <= " & Count & ""), "0.00"), ;
    Next
    Debug.Print
    Debug.Print
    
    For Count = 100 To 95 Step -1
        Debug.Print , Count;
    Next
    Debug.Print
    Debug.Print 1, ;
    For Count = 100 To 95 Step -1
        Debug.Print Format(DMedian("Data", "Observation", "Step = 1 And Data <= " & Count & ""), "0.00"), ;
    Next
    Debug.Print

End Function

Public Function ListThirdQuartile()

    Dim Method  As ApQuartileMethod
    Dim Count   As Long
    
    For Count = 40 To 70 Step 10
        Debug.Print , Count;
    Next
    Debug.Print
    For Method = ApQuartileMethod.[_First] To ApQuartileMethod.[_Last]
        Debug.Print Method, ;
        For Count = 40 To 70 Step 10
            Debug.Print Format(Quartile("Data", "Observation", "Step = 10 And Data <= " & Count & "", apThird, Method), "0.00"), ;
        Next
        Debug.Print
    Next
    Debug.Print
    
    For Count = 100 To 95 Step -1
        Debug.Print , Count;
    Next
    Debug.Print
    For Method = ApQuartileMethod.[_First] To ApQuartileMethod.[_Last]
        Debug.Print Method, ;
        For Count = 100 To 95 Step -1
            Debug.Print Format(Quartile("Data", "Observation", "Step = 1 And Data <= " & Count & "", apThird, Method), "0.00"), ;
        Next
        Debug.Print
    Next

End Function

' Lists the two examples found here:
'
'   https://en.wikipedia.org/wiki/Quartile
'
Public Function ListWikipediaSamples()

    Dim Part    As ApQuartilePart
    Dim Example As Integer
    
    Debug.Print , "Method 1", "Method 2", "Method 3"
    For Example = 1 To 2
        For Part = ApQuartilePart.apFirst To ApQuartilePart.apThird
            Debug.Print "Q" & Part, _
                Quartile("Data", "Wikipedia", "Example = " & Example & "", Part, apTukeyMooreMcCabe), _
                Quartile("Data", "Wikipedia", "Example = " & Example & "", Part, apTukey), _
                Quartile("Data", "Wikipedia", "Example = " & Example & "", Part, apHazen)
        Next
        Debug.Print
    Next

End Function

' Lists the example found in Quartile.xlsx:
'
Public Function ListExcelQuartile()

    Dim Method  As ApQuartileMethod
    Dim Count   As Long
    Dim Part    As Long
    
    For Count = 100 To 95 Step -1
        Debug.Print , Count;
    Next
    Debug.Print
    
    For Method = ApQuartileMethod.[_Last] To ApQuartileMethod.[_First] Step -1
        Select Case Method
            Case ApQuartileMethod.apFreundPerlesGumbell, ApQuartileMethod.apWeibull
                If Method = ApQuartileMethod.apFreundPerlesGumbell Then
                    Debug.Print "INCLUDE (LEGACY)"
                Else
                    Debug.Print "EXCLUDE"
                End If
                For Part = apFirst To apThird
                    Debug.Print Method, ;
                    For Count = 100 To 95 Step -1
                        Debug.Print Format(Quartile("Data", "Observation", "Step = 1 And Data <= " & Count & "", Part, Method), "0.00"), ;
                    Next
                    Debug.Print
                Next
                Debug.Print
        End Select
    Next

End Function

Public Function Log10( _
    ByVal Value As Double) _
    As Double

' Returns Log 10 of Value.
' 2015-08-17. Gustav Brock, Cactus Data ApS, CPH.

    Const Base10    As Double = 10

    ' No error handling is included, as that should be handled
    ' where this function is called.
    '
    ' Example:
    '
    '     If MyValue > 0 then
    '         LogMyValue = Log10(MyValue)
    '     Else
    '         ' Do something else ...
    '     End If
    
    Log10 = Log(Value) / Log(Base10)

End Function


