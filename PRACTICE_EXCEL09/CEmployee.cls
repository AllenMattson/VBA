VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' declarations
Private m_LastName As String
Private m_FirstName As String
Private m_Salary As Currency
Private m_ID As String

Property Get ID() As String
    ID = m_ID
End Property

Property Get LastName() As String
    LastName = m_LastName
End Property

Property Get FirstName() As String
    FirstName = m_FirstName
End Property

Property Get Salary() As Currency
    Salary = m_Salary
End Property

Property Let LastName(L As String)
    m_LastName = L
End Property

Property Let FirstName(F As String)
    m_FirstName = F
End Property

Property Let Salary(ByVal dollar As Currency)
    m_Salary = dollar
End Property


Public Function CalcNewSalary(choice As Integer, _
        curSalary As Currency, amount As Long) As Currency
    Select Case choice
        Case 1 ' by percent
            CalcNewSalary = curSalary + ((curSalary + amount) / 100)
        Case 2 ' by amount
            CalcNewSalary = curSalary + amount
    End Select
End Function

Private Sub Class_Initialize()
    Randomize
    m_ID = Int((99999 - 10000) * Rnd + 10000)
End Sub


