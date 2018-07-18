Attribute VB_Name = "Variables"
Option Explicit
   Dim slsTax As Single
   
   Sub CalcCost()
      ' declaration of variables
      Dim slsPrice As Currency
      
      Dim cost As Currency
      Dim strMsg As String

      slsPrice = 35
      slsTax = 0.085
      Range("A1").Formula = "The cost of calculator"
      Range("A4").Formula = "Price"
      Range("B4").Formula = slsPrice
      Range("A5").Formula = "Sales Tax"
      Range("A6").Formula = "Cost"
      Range("B5").Formula = Format((slsPrice * slsTax), "0.00")
      cost = Format(slsPrice + (slsPrice * slsTax), "0.00")

      Range("B6").Formula = cost
      strMsg = "The calculator total is $" & cost & "."
      Range("A8").Formula = strMsg
    End Sub

    Sub ExpenseRep()
      Dim slsPrice As Currency
      Dim cost As Currency

      slsPrice = 55.99

      cost = slsPrice + (slsPrice * slsTax)
      MsgBox slsTax
      MsgBox cost
    End Sub

Sub CostOfPurchase()
  ' declare variables
  Static allPurchase
  Dim newPurchase As String
  Dim purchCost As Single

  newPurchase = InputBox("Enter the cost of a purchase:")
  purchCost = CSng(newPurchase)
  allPurchase = allPurchase + purchCost

  ' display results
  MsgBox "The cost of a new purchase is: " & newPurchase
  MsgBox "The running cost is: " & allPurchase
End Sub

Sub UseObjVariable()
  Dim myRange As Object
  Set myRange = Worksheets("Sheet1").Range(Cells(1, 1), Cells(10, 5))
  myRange.BorderAround Weight:=xlMedium

  With myRange.Interior
    .ColorIndex = 6
    .Pattern = xlSolid
  End With

  Set myRange = Worksheets("Sheet1").Range(Cells(12, 5), Cells(12, 10))
  myRange.Value = 54

  Debug.Print IsObject(myRange)
End Sub


