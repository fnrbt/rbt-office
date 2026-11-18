Sub Main()
    Rnd -1
    Randomize 42
    P "rnd 1", Rnd
    P "rnd 2", Rnd
    P "rnd int", Int(Rnd * 100)
    P "rnd 0", Rnd(0)
End Sub
