Attribute VB_Name = "QuartileCode"
Option Compare Database
Option Explicit
'
' Cactus Data ApS, CPH
' 2019-08-16. (c) Gustav Brock
'

' Enums.

' Quartile parts.
' Values equal those of Excel's Quartile() function.
'
Public Enum ApQuartilePart
    [_First] = 0
    
    ' Minimum value.
    apMinimum = 0
    ' Lower quartile (first quartile, 25th percentile).
    apFirst = 1
    apLower = 1
    ' Median (second quartile, 50th percentile).
    apSecond = 2
    apMedian = 2
    ' Upper quartile (third quartile, 75th percentile).
    apThird = 3
    apUpper = 3
    ' Maximum value.
    apMaximum = 4
    
    [_Last] = 4
End Enum

' Quartile calculation methods.
' Values equal those listed in the source. See function Quartile.
'
' Common names of variables used in calculation formulas.
'
' L: Q1, Lower quartile.
' H: Q3, Higher quartile.
' M: Q2, Median (not used here).
' n: Count of elements.
' p: Calculated position of quartile.
' j: Element of dataset.
' g: Decimal part of p to be used for interpolation between j and j+1.
'
Public Enum ApQuartileMethod
    [_First] = 1
    
    ' Basic calculation methods.
    
    ' Step. Mendenhall and Sincich method.
    '   SAS #3.
    '   Round up to actual element of dataset.
    '   L:  -Int(-n/4)
    '   H: n-Int(-n/4)
    apMendenhallSincich = 1
    
    ' Average step.
    '   SAS #5, Minitab (%DESCRIBE), GLIM (percentile).    '
    '   Add bias of one on basis of n/4.
    '   L:   CLng((n+2)/2)/2
    '   H: n-Clng((n+2)/2)/2
    '   Note:
    '       Replaces these original formulas that don't return the expected values.
    '   L:   (Int((n+1)/4)+Int(n/4))/2+1
    '   H: n-(Int((n+1)/4)+Int(n/4))/2+1
    apAverage = 2
    
    ' Nearest integer to np.
    '   SAS #2.
    '   Round to nearest integer on basis of n/4.
    '   L:   CLng(n/4)
    '   H: n-CLng(n/4)
    '   Note:
    '       Replaces these original formulas that don't return the expected values.
    '   L:   Int((n+2)/4)
    '   H: n-Int((n+2)/4)
    apNearestInteger = 3
    
    ' Parzen method.
    '   Method 1 with interpolation.
    '   SAS #1.
    '   L: n/4
    '   H: 3n/4
    apParzen = 4
    
    ' Hazen method.
    '   Values midway between method 1 steps.
    '   GLIM (interpolate).
    '   Wikipedia method 3.
    '   Add bias of 2, don't round to actual element of dataset.
    '   L: (n+2)/4
    '   H: 3(n+2)/4-1
    apHazen = 5
    
    ' Weibull method.
    '   SAS #4. Minitab (DECRIBE), SPSS, BMDP, Excel exclusive.
    '   Add bias of 1, don't round to actual element of dataset.
    '   L: (n+1)/4
    '   H: 3(n+1)/4
    apWeibull = 6
    
    ' Freund, J. and Perles, B., Gumbell method.
    '   S-PLUS, R, Excel legacy, Excel inclusive, Star Office Calc.
    '   Add bias of 3, don't round to actual element of dataset.
    '   L: (n+3)/4
    '   H: (3n+1)/4
    apFreundPerlesGumbell = 7
    
    ' Median Position.
    '   Median unbiased.
    '   L: (3n+5)/12
    '   H: (9n+7)/12
    apMedianPosition = 8
    
    ' Bernard and Bos-Levenbach.
    '   L: (n/4)+0.4
    '   H: (3n/4)/+0.6
    '   Note:
    '       Reference claims L to be (n/4)+0.31.
    apBernardBosLevenbach = 9
    
    ' Blom's Plotting Position.
    '   Better approximation when the distribution is normal.
    '   L: (4n+7)/16
    '   H: (12n+9)/16
    apBlom = 10
    
    ' Moore's first method.
    '   Add bias of one half step.
    '   L: (n+0.5)/4
    '   H: n-(n+0.5)/4
    apMooreFirst = 11
    
    ' Moore's second method.
    '   Add bias of one or two steps on basis of (n+1)/4.
    '   L:   (Int((n+1)/4)+Int(n/4))/2+1
    '   H: n-(Int((n+1)/4)+Int(n/4))/2+1
    apMooreSecond = 12
    
    ' John Tukey's method.
    '   Include median from odd dataset in dataset for quartile.
    '   Wikipedia method 2.
    '   L:   (1-Int(-n/2))/2
    '   H: n-(-1-Int(-n/2))/2
    apTukey = 13
    
    ' Moore and McCabe (M & M), variation of John Tukey's method.
    '   TI-83.
    '   Wikipedia method 1.
    '   Exclude median from odd dataset in dataset for quartile.
    '   L:   (Int(n/2)+1)/2
    '   H: n-(Int(n/2)-1)/2
    apTukeyMooreMcCabe = 14
    
    ' Additional variations between Weibull's and Hazen's methods, from
    '   (i-0.000)/(n+1.00)
    ' to
    '   (i-0.500)/(n+0.00)
    
    ' Variation of Weibull.
    '   L: n(n/4-0)/(n+1)
    '   H: n(3n/4-0)/(n+1)
    apWeibullVariation = 15
    
    ' Variation of Blom.
    '   L: n(n/4-3/8)/(n+1/4)
    '   H: n(3n/4-3/8)/(n+1/4)
    apBlomVariation = 16
    
    ' Variation of Tukey.
    '   L: n(n/4-1/3)/(n+1/3)
    '   H: n(3n/4-1/3)/(n+1/3)
    apTukeyVariation = 17
    
    ' Variation of Cunnane.
    '   L: n(n/4-2/5)/(n+1/5)
    '   H: n(3n/4-2/5)/(n+1/5)
    apCunnaneVariation = 18
    
    ' Variation of Gringorten.
    '   L: n(n/4-0.44)/(n+0.12)
    '   H: n(3n/4-0.44)/(n+0.12)
    apGringortenVariation = 19
    
    ' Variation of Hazen.
    '   L: n(n/4-1/2)/n
    '   H: n(3n/4-1/2)/n
    apHazenVariation = 20
    
    [_Last] = 20
End Enum

' Returns the median of a field of a table/query.
'
' Parameters:
'   Expression: Name of the field or an expression to analyse.
'   Domain    : Name of the source/query, or an SQL select query, to analyse.
'   Criteria  : Optional. A filter expression for Domain.
'
' Reference and examples: See function Quartile.
'
' Data must be in ascending order by Field.
'
' 2019-08-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DMedian( _
    ByVal Expression As String, _
    ByVal Domain As String, _
    Optional ByVal Criteria As String) _
    As Double
    
    Dim Value       As Double
    
    Value = Quartile(Expression, Domain, Criteria)
    
    DMedian = Value

End Function

' Returns the upper or lower quartile or the median or the
' minimum or maximum value of a field of a table/query
' using the method by Freund, Perles, and Gumbell (Excel).
'
' Parameters:
'   Expression: Name of the field or an expression to analyse.
'   Domain    : Name of the source/query, or an SQL select query, to analyse.
'   Criteria  : Optional. A filter expression for Domain.
'   Part      : Optional. Which median/quartile or min/max value to return.
'               Default is the median value.
'
' Reference and examples: See function Quartile.
'
' 2019-08-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DQuartile( _
    ByVal Expression As String, _
    ByVal Domain As String, _
    Optional ByVal Criteria As String, _
    Optional ByVal Part As ApQuartilePart = ApQuartilePart.apMedian) _
    As Double
    
    Dim Value       As Double
    
    Value = Quartile(Expression, Domain, Criteria, Part)
    
    DQuartile = Value

End Function

' Returns the upper or lower quartile or the median or the
' minimum or maximum value of a field of a table/query
' using one of twenty calculation methods.
'
' Parameters:
'   Expression: Name of the field or an expression to analyse.
'   Domain    : Name of the Source/query, or an SQL select query, to analyse.
'   Criteria  : Optional. A filter expression for Domain.
'   Part      : Optional. Which median/quartile or min/max value to return.
'               Default is the median value.
'   Method    : Optional. Method for calculation of lower/higher quartile.
'               Default is the method by Freund, Perles, and Gumbell (Excel).
'
' Reference for the methods for calculation:
'
'   Original source (now off-line) by David A. Heiser:
'       http://www.daheiser.info/excel/notes/NOTE%20N.pdf
'   Archived source:
'       https://web.archive.org/web/20110721195325/http://www.daheiser.info/excel/notes/NOTE%20N.pdf
'
' Note: Source H-4, p. 4, has correct data for the dataset for 1-96 while the
'   datasets for 1-100 to 1-97 actually are the datasets for 1-99 to 1-96
'   shifted one column left.
'   Thus, the dataset for 1-100 is missing, and that for 1-96 is listed twice.
'
'   Method 3b is not implemented as no one seems to use it.
'   Neither are no example data given.
'   Thus method 3a has here been labelled method 3.
'
' Further notes on methods here:
'   https://en.wikipedia.org/wiki/Quartile
'   http://mathforum.org/library/drmath/view/60969.html
'   http://www.haiweb.org/medicineprices/manual/quartiles_iTSS.pdf
'   https://web.archive.org/web/20020707022505/http://wwwmaths.murdoch.edu.au/units/c503a/unitnotes/boxhisto/quartilesmore.html
'
' Example calls and the internally generated SQL:
'
'   With fieldname as expression, table (or query) as domain, no filter, and default sorting:
'       Q1 = Quartile("Data", "Observation", , apFirst, apFreundPerlesGumbell)
'       Select Data From Observation Order By Data Asc
'
'   With two fieldnames as expression, table (or query) as domain, no filter, and sorting on two fields:
'       Q1 = Quartile("Data, Step", "Observation", , apFirst, apFreundPerlesGumbell)
'       Select Data, Step From Observation Order By Data, Step Asc
'
'   With fieldname as expression, SQL as domain, no filter, and default sorting:
'       Q1 = Quartile("Data", "Select Data From Observation", , apFirst, apFreundPerlesGumbell)
'       Select Data From (Select Data From Observation) As T Order By Data Asc
'
'   With fieldname as expression, SQL as domain, simple filter, and sorting on one field:
'       Q1 = Quartile("Data", "Select Data, Step From Observation", "Step = 10", apFirst, apFreundPerlesGumbell)
'       Select Data From (Select Data, Step From Observation) As T Where Step = 10 Order By Data Asc
'
'   With calculated expression, SQL as domain, extended filter, and sorting on one field:
'       Q1 = Quartile("Data * 10", "Select Data, Step From Observation", "Step = 10 And Data <= 40", apFirst, apFreundPerlesGumbell)
'       Select Data * 10 From (Select Data, Step From Observation) As T Where Step = 10 And Data <= 40 Order By Data * 10 Asc
'
'   With filtered SQL domain, additional filter, and sorting on one field:
'       Q1 = Quartile("Data", "Select Data, Step From Observation Where Step = 10", "Data <= 40", apFirst, apFreundPerlesGumbell)
'       Select Data From (Select Data, Step From Observation Where Step = 10) As T Where Data <= 40 Order By Data Asc
'
'   With filtered SQL domain, additional filter, and sorting on two fields:
'       Q1 = Quartile("Step, Data", "Select Data, Step From Observation Where Step = 10", "Data <= 40", apFirst, apFreundPerlesGumbell)
'       Select Step, Data From (Select Data, Step From Observation Where Step = 10) As T Where Data <= 40 Order By Step, Data Asc
'
' 2019-08-16. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Quartile( _
    ByVal Expression As String, _
    ByVal Domain As String, _
    Optional ByVal Criteria As String, _
    Optional ByVal Part As ApQuartilePart = ApQuartilePart.apMedian, _
    Optional ByVal Method As ApQuartileMethod = ApQuartileMethod.apFreundPerlesGumbell) _
    As Double
  
    ' SQL.
    Const SqlMask           As String = "Select {0} From {1} {2}"
    Const SqlLead           As String = "Select "
    Const SubMask           As String = "({0}) As T"
    Const FilterMask        As String = "Where {0} "
    Const OrderByMask       As String = "Order By {0} Asc"
    
    Dim Records     As DAO.Recordset
    
    Dim Sql         As String
    Dim SqlSub      As String
    Dim Filter      As String
    Dim Count       As Long     ' n.
    Dim Position    As Double   ' p.
    Dim Element     As Long     ' j.
    Dim Interpolate As Double   ' g.
    Dim ValueOne    As Double
    Dim ValueTwo    As Double
    Dim Value       As Double
    
    ' Return default quartile part if choice of part is
    ' outside the range of ApQuartilePart.
    If Not IsQuartilePart(Part) Then
        Part = ApQuartilePart.apMedian
    End If
    
    ' Use a default calculation method if choice of method is
    ' outside the range of ApQuartileMethod.
    If Not IsQuartileMethod(Method) Then
        Method = ApQuartileMethod.apFreundPerlesGumbell
    End If
    
    If Domain <> "" And Expression <> "" Then
        ' Build SQL to lookup values.
        If InStr(1, LTrim(Domain), SqlLead, vbTextCompare) = 1 Then
            ' Domain is an SQL expression.
            SqlSub = Replace(SubMask, "{0}", Domain)
        Else
            ' Domain is a table or query name.
            SqlSub = Domain
        End If
        If Trim(Criteria) <> "" Then
            ' Build Where clause.
            Filter = Replace(FilterMask, "{0}", Criteria)
        End If
        ' Build final SQL.
        Sql = Replace(Replace(Replace(SqlMask, "{0}", Expression), "{1}", SqlSub), "{2}", Filter) & _
            Replace(OrderByMask, "{0}", Expression)
        Set Records = CurrentDb.OpenRecordset(Sql, dbOpenSnapshot)
      
        With Records
            If Not .EOF = True Then
                If Part = ApQuartilePart.apMinimum Then
                    ' No need to count records.
                    Count = 1
                Else
                    ' Count records.
                    .MoveLast
                    Count = .RecordCount
                End If
                Select Case Part
                    Case ApQuartilePart.apMinimum
                        ' Current record is first record.
                        ' Read value of this record.
                    Case ApQuartilePart.apMaximum
                        ' Current record is last record.
                        ' Read value of this record.
                    Case ApQuartilePart.apMedian
                        ' Locate position of median.
                        Position = (Count + 1) / 2
                    Case ApQuartilePart.apLower
                        Select Case Method
                            Case ApQuartileMethod.apMendenhallSincich
                                Position = -Int(-Count / 4)
                            Case ApQuartileMethod.apAverage
                                Position = CLng((Count + 2) / 2) / 2
                            Case ApQuartileMethod.apNearestInteger
                                Position = CLng(Count / 4)
                            Case ApQuartileMethod.apParzen
                                Position = Count / 4
                            Case ApQuartileMethod.apHazen
                                Position = (Count + 2) / 4
                            Case ApQuartileMethod.apWeibull
                                Position = (Count + 1) / 4
                            Case ApQuartileMethod.apFreundPerlesGumbell
                                Position = (Count + 3) / 4
                            Case ApQuartileMethod.apMedianPosition
                                Position = (3 * Count + 5) / 12
                            Case ApQuartileMethod.apBernardBosLevenbach
                                Position = (Count / 4) + 0.4
                            Case ApQuartileMethod.apBlom
                                Position = (4 * Count + 7) / 16
                            Case ApQuartileMethod.apMooreFirst
                                Position = (Count + 0.5) / 4
                            Case ApQuartileMethod.apMooreSecond
                                Position = (Int((Count + 1) / 4) + Int(Count / 4)) / 2 + 1
                            Case ApQuartileMethod.apTukey
                                Position = (1 - Int(-Count / 2)) / 2
                            Case ApQuartileMethod.apTukeyMooreMcCabe
                                Position = (Int(Count / 2) + 1) / 2
                            Case ApQuartileMethod.apWeibullVariation
                                Position = Count * (Count / 4) / (Count + 1)
                            Case ApQuartileMethod.apBlomVariation
                                Position = Count * (Count / 4 - 3 / 8) / (Count + 1 / 4)
                            Case ApQuartileMethod.apTukeyVariation
                                Position = Count * (Count / 4 - 1 / 3) / (Count + 1 / 3)
                            Case ApQuartileMethod.apCunnaneVariation
                                Position = Count * (Count / 4 - 2 / 5) / (Count + 1 / 5)
                            Case ApQuartileMethod.apGringortenVariation
                                Position = Count * (Count / 4 - 0.44) / (Count + 0.12)
                            Case ApQuartileMethod.apHazenVariation
                                Position = Count * (Count / 4 - 1 / 2) / Count
                        End Select
                    Case ApQuartilePart.apUpper
                        ' Default position for very low counts for several methods
                        Position = Count
                        Select Case Method
                            Case ApQuartileMethod.apMendenhallSincich
                                If Count > 2 Then
                                    Position = Count - (-Int(-Count / 4))
                                End If
                            Case ApQuartileMethod.apAverage
                                If Count > 2 Then
                                    Position = Count - CLng((Count + 2) / 2) / 2
                                End If
                            Case ApQuartileMethod.apNearestInteger
                                Position = Count - CLng(Count / 4)
                            Case ApQuartileMethod.apParzen
                                Position = 3 * Count / 4
                            Case ApQuartileMethod.apHazen
                                If Count > 1 Then
                                    Position = 3 * (Count + 2) / 4 - 1
                                End If
                            Case ApQuartileMethod.apWeibull
                                If Count > 2 Then
                                    Position = 3 * (Count + 1) / 4
                                End If
                            Case ApQuartileMethod.apFreundPerlesGumbell
                                Position = (3 * Count + 1) / 4
                            Case ApQuartileMethod.apMedianPosition
                                If Count > 2 Then
                                    Position = (9 * Count + 7) / 12
                                End If
                            Case ApQuartileMethod.apBernardBosLevenbach
                                If Count > 2 Then
                                    Position = (3 * Count / 4) + 0.6
                                End If
                            Case ApQuartileMethod.apBlom
                                If Count > 2 Then
                                    Position = (12 * Count + 9) / 16
                                End If
                            Case ApQuartileMethod.apMooreFirst
                                Position = Count - (Count + 0.5) / 4
                            Case ApQuartileMethod.apMooreSecond
                                ' Basic calculation method. Will fail for 2 or 3 elements.
                                '   Position = Count - (Int((Count + 1) / 4) + Int(Count / 4)) / 2 + 1
                                ' Calculation method adjusted to accept 2 or 3 elements.
                                Position = Count - (Int((Count + Int((Count * 2) / (Count + 4))) / 4) + Int(Count / 4)) / 2 + 1
                            Case ApQuartileMethod.apTukey
                                Position = Count - (-1 - Int(-Count / 2)) / 2
                            Case ApQuartileMethod.apTukeyMooreMcCabe
                                If Count > 1 Then
                                    Position = Count - (Int(Count / 2) - 1) / 2
                                End If
                            Case ApQuartileMethod.apWeibullVariation
                                Position = Count * (3 * Count / 4) / (Count + 1)
                            Case ApQuartileMethod.apBlomVariation
                                Position = Count * (3 * Count / 4 - 3 / 8) / (Count + 1 / 4)
                            Case ApQuartileMethod.apTukeyVariation
                                Position = Count * (3 * Count / 4 - 1 / 3) / (Count + 1 / 3)
                            Case ApQuartileMethod.apCunnaneVariation
                                Position = Count * (3 * Count / 4 - 2 / 5) / (Count + 1 / 5)
                            Case ApQuartileMethod.apGringortenVariation
                                Position = Count * (3 * Count / 4 - 0.44) / (Count + 0.12)
                            Case ApQuartileMethod.apHazenVariation
                                Position = Count * (3 * Count / 4 - 1 / 2) / Count
                        End Select
                End Select
                Select Case Part
                    Case ApQuartilePart.apMinimum, ApQuartilePart.apMaximum
                        ' Read current row.
                    Case Else
                        .MoveFirst
                        ' Find position of first observation to retrieve.
                        ' If Element is 0, then upper position is first record.
                        ' If Element is not 0 and position is not an integer, then
                        ' read the next observation too.
                        Element = Fix(Position)
                        Interpolate = Position - Element
                        If Count = 1 Then
                            ' Nowhere else to move.
                            If Interpolate < 0 Then
                                ' Prevent values to be created by extrapolation beyond zero from observation one
                                ' for these methods:
                                '   ApQuartileMethod.apBlomVariation
                                '   ApQuartileMethod.apTukeyVariation
                                '   ApQuartileMethod.apCunnaneVariation
                                '   ApQuartileMethod.apGringortenVariation
                                '   ApQuartileMethod.apHazenVariation
                                '
                                ' Comment this line out, if reading by extrapolation *is* requested.
                                Interpolate = 0
                            End If
                        ElseIf Element > 1 Then
                            ' Move to the record to read.
                            .Move Element - 1
                            ' Special case for apMooreSecond and upper quartile for 2 and 3 elements.
                            If .EOF Then
                                .MoveLast
                            End If
                        End If
                End Select
                ' Retrieve value from first observation.
                ValueOne = .Fields(0).Value
          
                Select Case Part
                    Case ApQuartilePart.apMinimum, ApQuartilePart.apMaximum
                        Value = ValueOne
                    Case Else
                        If Interpolate = 0 Then
                            ' Only one observation to read.
                            If Element = 0 Then
                                ' Return 0.
                            Else
                                Value = ValueOne
                            End If
                        Else
                            If Element = 0 Or Element = Count Then
                                ' No first/last observation to retrieve.
                                ValueTwo = ValueOne
                                If ValueOne > 0 Then
                                    ' Use 0 as other observation.
                                    ValueOne = 0
                                Else
                                    ValueOne = 2 * ValueOne
                                End If
                            Else
                                ' Move to next observation.
                                .MoveNext
                                ' Retrieve value from second observation.
                                ValueTwo = .Fields(0).Value
                            End If
                            ' For positive values interpolate between 0 and ValueOne.
                            ' For negative values interpolate between 2 * ValueOne and ValueOne.
                            ' Calculate quartile using linear interpolation.
                            Value = ValueOne + Interpolate * CDec(ValueTwo - ValueOne)
                        End If
                End Select
            End If
            .Close
        End With
    End If
      
    Quartile = Value

End Function

' Returns True if the passed Method is member of enum ApQuartileMethod,
' False if not.
'
' 2019-08-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsQuartileMethod( _
    ByVal Method As ApQuartileMethod) _
    As Boolean

    Dim Found   As Boolean
    
    If Method >= ApQuartileMethod.[_First] And Method <= ApQuartileMethod.[_Last] Then
        Found = True
    End If
    
    IsQuartileMethod = Found
    
End Function

' Returns True if the passed Part is member of enum ApQuartilePart,
' False if not.
'
' 2019-08-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsQuartilePart( _
    ByVal Part As ApQuartilePart) _
    As Boolean

    Dim Found   As Boolean
    
    If Part >= ApQuartilePart.[_First] And Part <= ApQuartilePart.[_Last] Then
        Found = True
    End If
    
    IsQuartilePart = Found
    
End Function

' Returns the literal name of the passed quartile calculation method.
'
' 2019-08-15. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function QuartileMethodName( _
    ByVal Method As ApQuartileMethod) _
    As String
    
    Dim Name    As String
    
    Select Case Method
        Case apMendenhallSincich
            Name = "Step. Mendenhall and Sincich"
        Case apAverage
            Name = "Average step"
        Case apNearestInteger
            Name = "Nearest integer to np"
        Case apParzen
            Name = "Parzen"
        Case apHazen
            Name = "Hazen"
        Case apWeibull
            Name = "Weibull"
        Case apFreundPerlesGumbell
            Name = "Freund, J. and Perles, B., Gumbell"
        Case apMedianPosition
            Name = "Median Position"
        Case apBernardBosLevenbach
            Name = "Bernard and Bos-Levenbach"
        Case apBlom
            Name = "Blom's Plotting Position"
        Case apMooreFirst
            Name = "Moore's first method"
        Case apMooreSecond
            Name = "Moore's second method"
        Case apTukey
            Name = "John Tukey's method"
        Case apTukeyMooreMcCabe
            Name = "Moore and McCabe (M & M)"
        Case apWeibullVariation
            Name = "Variation of Weibull"
        Case apBlomVariation
            Name = "Variation of Blom"
        Case apTukeyVariation
            Name = "Variation of Tukey"
        Case apCunnaneVariation
            Name = "Variation of Cunnane"
        Case apGringortenVariation
            Name = "Variation of Gringorten"
        Case apHazenVariation
            Name = "Variation of Hazen"
    End Select
    
    QuartileMethodName = Name
    
End Function

