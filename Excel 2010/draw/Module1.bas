Attribute VB_Name = "Module1"
Function Draw(Rng As Variant, Optional Recalc As Boolean = False)
'    Chooses one cell at random from a range

'    Make function volatile if Recalc is True
     Application.Volatile Recalc

'    Determine a random cell
     Draw = Rng(Int((Rng.Count) * Rnd + 1))
End Function



