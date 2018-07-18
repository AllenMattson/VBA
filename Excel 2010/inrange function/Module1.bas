Attribute VB_Name = "Module1"
Option Explicit

Function InRange(rng1, rng2) As Boolean
'   Returns True if rng1 is a subset of rng2
    InRange = False
    If rng1.Parent.Parent.Name = rng2.Parent.Parent.Name Then
        If rng1.Parent.Name = rng2.Parent.Name Then
            If Union(rng1, rng2).Address = rng2.Address Then
                InRange = True
            End If
        End If
    End If
End Function

