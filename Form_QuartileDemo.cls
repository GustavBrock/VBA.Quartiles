VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_QuartileDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()

    Const HeaderIndex   As Integer = 0
    Const FirstIndex    As Integer = ApQuartileMethod.[_First]
    Const LastIndex     As Integer = ApQuartileMethod.[_Last]
    
    Dim Methods(HeaderIndex To LastIndex)   As String
    Dim Method          As ApQuartileMethod
    
    Methods(HeaderIndex) = """    #"";""Method"""
    For Method = ApQuartileMethod.[_First] To ApQuartileMethod.[_Last]
        Methods(Method) = """" & Space(4 - 2 * Int(Log10(Method))) & Format(Method, "0") & """;""" & QuartileMethodName(Method) & """"
    Next
    Me!Methods.RowSource = Join(Methods, ";")
    
End Sub

Private Sub Methods_AfterUpdate()

    Const HeaderIndex   As Integer = 0
    Const FirstIndex    As Integer = 1
    Const LastIndex     As Integer = 100
    
    Dim Values(HeaderIndex To LastIndex)    As String
    Dim Value       As String
    Dim Index       As Integer
    
    Dim Method      As ApQuartileMethod
    Dim Quartile1   As Double
    Dim Quartile2   As Double
    Dim Quartile3   As Double
    
    Method = Me!Methods.Value
    
    Values(HeaderIndex) = """    #"";""    Q1"";""    Q2"";""    Q3"""
    For Index = FirstIndex To LastIndex
        Quartile1 = Quartile("Data", "Select * From Observation Where Step = 1", "Data <= " & Index & "", apFirst, Method)
        Quartile2 = Quartile("Data", "Select * From Observation Where Step = 1", "Data <= " & Index & "", apSecond, Method)
        Quartile3 = Quartile("Data", "Select * From Observation Where Step = 1", "Data <= " & Index & "", apThird, Method)
        Values(Index) = _
            """" & Space(4 - 2 * Int(Log10(Index))) & Format(Index, "0") & """;""" & _
            Format(Quartile1, " 00.000") & """;""" & _
            Format(Quartile2, " 00.000") & """;""" & _
            Format(Quartile3, " 00.000") & """"
    Next
    
    Me!Values.RowSource = Join(Values, ";")
        
End Sub


