Private Function stripLeadingSlashes(ByVal strText As String) As String

    Dim lngCounter As Long
    Dim lngStringLength As Long
    
    lngStringLength = Len(strText)              'Define string length
    
    For lngCounter = 1 To Len(strText)
    
        Select Case Left(strText, 1)            'Loop through string
            Case "\", "/"                       'If char is slash, strip it
                strText = Right(strText, lngStringLength - 1)
                lngStringLength = lngStringLength - 1
        
            Case Else                           'If char is not slash, exit
                stripLeadingSlashes = strText
                GoTo exitMe

        End Select
    Next lngCounter
    
exitMe:
                                                'Nothing to clean up
End Function
