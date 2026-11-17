#load "function-table-kit.fsx"
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open ``Function-table-kit``

// Database, engineering (base conversion, bits) and statistical functions.

// ----- database: Microsoft's tree example (A1:F3 criteria, A5:E11 database)
let trees =
    let header = [ "Tree"; "Height"; "Age"; "Yield"; "Profit"; "Height" ]
    let rows =
        [ [ S "Apple"; N 18.0; N 20.0; N 14.0; N 105.0 ]
          [ S "Pear"; N 12.0; N 12.0; N 10.0; N 96.0 ]
          [ S "Cherry"; N 13.0; N 14.0; N 9.0; N 105.0 ]
          [ S "Apple"; N 14.0; N 15.0; N 10.0; N 75.0 ]
          [ S "Pear"; N 9.0; N 8.0; N 8.0; N 76.8 ]
          [ S "Apple"; N 8.0; N 9.0; N 6.0; N 45.0 ] ]
    modelOf (
        [ for c, h in List.indexed header -> (0, c), S h ]
        @ [ (1, 0), S "=Apple"; (1, 1), S ">10"; (1, 5), S "<16"; (2, 0), S "=Pear" ]
        @ [ for c, h in List.indexed (List.take 5 header) -> (4, c), S h ]
        @ [ for r, row in List.indexed rows do for c, v in List.indexed row -> (5 + r, c), v ])
approx trees "=DSTDEV(A5:E11,\"Yield\",A1:A3)" 5 2.96648
approx trees "=DSTDEVP(A5:E11,\"Yield\",A1:A3)" 5 2.65330
approx trees "=DVAR(A5:E11,\"Yield\",A1:A3)" 6 8.8
approx trees "=DVARP(A5:E11,\"Yield\",A1:A3)" 6 7.04
// Apples taller than 10 and shorter than 16, or pears: yields 10, 10, 8.
approx trees "=DVAR(A5:E11,4,A1:F3)" 6 (4.0 / 3.0)
error trees "=DSTDEV(A5:E11,\"Nope\",A1:A3)" Err.value
error trees "=DVAR(A5:E11,\"Yield\",A3:A3)" Err.value

// ----- base conversion and bits
let m = empty
text m "=BIN2HEX(11111011,4)" "00FB"
text m "=BIN2HEX(1110)" "E"
text m "=BIN2HEX(1111111111)" "FFFFFFFFFF"
text m "=BIN2OCT(1001,3)" "011"
text m "=BIN2OCT(1100100)" "144"
text m "=BIN2OCT(1111111111)" "7777777777"
text m "=HEX2BIN(\"F\",8)" "00001111"
text m "=HEX2BIN(\"B7\")" "10110111"
text m "=HEX2BIN(\"FFFFFFFFFF\")" "1111111111"
text m "=HEX2OCT(\"F\",3)" "017"
text m "=HEX2OCT(\"3B4E\")" "35516"
text m "=HEX2OCT(\"FFFFFFFF00\")" "7777777400"
text m "=OCT2BIN(3,3)" "011"
text m "=OCT2BIN(7777777000)" "1000000000"
text m "=OCT2HEX(100,4)" "0040"
text m "=OCT2HEX(7777777533)" "FFFFFFFF5B"
error m "=HEX2BIN(\"200\")" Err.num
error m "=BIN2HEX(12)" Err.num
error m "=BIN2HEX(11111111111)" Err.num
error m "=BIN2HEX(11111011,1)" Err.num
error m "=BIN2HEX(1,-1)" Err.num
error m "=BIN2HEX(1,\"a\")" Err.value
error m "=BIN2HEX(TRUE)" Err.value
number m "=BITLSHIFT(4,2)" 16.0
number m "=BITRSHIFT(13,2)" 3.0
number m "=BITLSHIFT(13,-2)" 3.0
error m "=BITLSHIFT(1.5,1)" Err.num
error m "=BITLSHIFT(2^47,2)" Err.num
number m "=DELTA(5,4)" 0.0
number m "=DELTA(5,5)" 1.0
number m "=DELTA(0.5)" 0.0
number m "=GESTEP(5,4)" 1.0
number m "=GESTEP(5,5)" 1.0
number m "=GESTEP(-4,-5)" 1.0
number m "=GESTEP(-1)" 0.0
error m "=DELTA(\"a\",1)" Err.value

// ----- statistics
let data = "{3,4,5,2,3,4,5,6,4,7}"
approx m ("=KURT(" + data + ")") 9 -0.151799637
approx m ("=SKEW(" + data + ")") 6 0.359543
approx m ("=SKEW.P(" + data + ")") 6 0.303193
error m "=KURT(1,2,3)" Err.div0
error m "=SKEW(1,1,1)" Err.div0
// STDEVA and kin over cells: text counts 0, TRUE 1, blank cells are skipped.
let tensile = modelOf [ for r, v in List.indexed [ 1345.0; 1301.0; 1368.0; 1322.0; 1310.0; 1370.0; 1318.0; 1350.0; 1303.0; 1299.0 ] -> (r, 0), N v ]
approx tensile "=STDEVA(A1:A10)" 5 27.46392
approx tensile "=STDEVPA(A1:A10)" 5 26.05456
approx tensile "=VARA(A1:A10)" 4 754.2667
approx tensile "=VARPA(A1:A10)" 4 678.84
let mixed = modelOf [ (0, 0), N 2.0; (1, 0), S "x"; (2, 0), B true; (4, 0), N 4.0 ]
// 2, 0, 1, 4: mean 1.75, squares 0.0625+3.0625+0.5625+5.0625 = 8.75.
approx mixed "=VARPA(A1:A5)" 9 (8.75 / 4.0)
approx mixed "=VARA(A1:A5)" 9 (8.75 / 3.0)
approx mixed "=VARA(2,TRUE)" 9 0.5
error mixed "=VARA(2,\"x\")" Err.value
error mixed "=STDEVA(A1)" Err.div0
approx m "=TRIMMEAN({4,5,6,7,2,3,4,5,1,2,3},0.2)" 9 3.777777778
error m "=TRIMMEAN({1,2},1)" Err.num
let ranked = "{1,2,3,6,6,6,7,8,9}"
approx m ("=PERCENTILE.EXC(" + ranked + ",0.25)") 9 2.5
error m ("=PERCENTILE.EXC(" + ranked + ",0)") Err.num
error m ("=PERCENTILE.EXC(" + ranked + ",0.01)") Err.num
approx m ("=PERCENTILE.EXC(" + ranked + ",0.5)") 9 6.0
let quartiles = "{6,7,15,36,39,40,41,42,43,47,49}"
approx m ("=QUARTILE.EXC(" + quartiles + ",1)") 9 15.0
approx m ("=QUARTILE.EXC(" + quartiles + ",3)") 9 43.0
error m ("=QUARTILE.EXC(" + quartiles + ",4)") Err.num
approx m ("=PERCENTRANK.EXC(" + ranked + ",7)") 9 0.7
approx m ("=PERCENTRANK.EXC(" + ranked + ",5.43)") 9 0.381
approx m ("=PERCENTRANK.EXC(" + ranked + ",5.43,1)") 9 0.3
error m ("=PERCENTRANK.EXC(" + ranked + ",10)") Err.na
eq "MODE.MULT" (arrayOf m "=MODE.MULT({1,2,3,4,3,2,1,2,3,5,6,1})") [| [| N 1.0 |]; [| N 2.0 |]; [| N 3.0 |] |]
error m "=MODE.MULT({1,2,3})" Err.na
number m "=PERMUT(100,3)" 970200.0
number m "=PERMUT(3,2)" 6.0
error m "=PERMUT(2,3)" Err.num
number m "=PERMUTATIONA(3,2)" 9.0
number m "=PERMUTATIONA(2,2)" 4.0
approx m "=PROB({0,1,2,3},{0.2,0.3,0.1,0.4},2)" 9 0.1
approx m "=PROB({0,1,2,3},{0.2,0.3,0.1,0.4},1,3)" 9 0.8
error m "=PROB({0,1,2,3},{0.2,0.3,0.1,0.3},1,3)" Err.num
error m "=PROB({0,1,2},{0.2,0.8},1)" Err.na
approx m "=FISHER(0.75)" 6 0.972955
approx m "=FISHERINV(0.972955)" 6 0.75
error m "=FISHER(1)" Err.num
approx m "=STANDARDIZE(42,40,1.5)" 6 1.333333
error m "=STANDARDIZE(42,40,0)" Err.num
approx m "=PHI(0.75)" 9 0.301137432
approx m "=GAUSS(2)" 5 0.47725
approx m "=CONFIDENCE.NORM(0.05,2.5,50)" 6 0.692952
approx m "=CONFIDENCE(0.05,2.5,50)" 6 0.692952
error m "=CONFIDENCE.NORM(0,2.5,50)" Err.num
let zdata = "{3,6,7,8,6,5,4,2,1,9}"
approx m ("=Z.TEST(" + zdata + ",4)") 6 0.090574
approx m ("=Z.TEST(" + zdata + ",6)") 6 0.863043
approx m ("=ZTEST(" + zdata + ",4)") 6 0.090574
error m "=Z.TEST({1,1},1)" Err.div0

printfn "stage60 database, engineering and statistics functions: %d passed, %d failed" pass fail
if fail > 0 then exit 1
