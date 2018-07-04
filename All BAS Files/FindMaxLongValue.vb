Function FindMaxLongValue( _
    ByVal SearchRange As Range) As Long

    'Finds the maximum integer value within the given range
    '  and returns is as a long variable. Will error out
    '  if max value is outside the range -2,147,483,648 to
    '  2,147,483,647

    ' Uses the Application.Max function to find the max value

    FindMaxLongValue = Application.Max(SearchRange)

End Function
