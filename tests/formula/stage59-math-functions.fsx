#load "function-table-kit.fsx"
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open ``Function-table-kit``

// Math and trigonometry functions of the Math module.
let m = empty
approx m "=ACOT(2)" 4 0.4636
approx m "=ACOTH(6)" 4 0.1682
error m "=ACOTH(0.5)" Err.num
approx m "=COTH(2)" 4 1.0373
error m "=COTH(0)" Err.div0
approx m "=CSC(15)" 4 1.5378
error m "=CSC(0)" Err.div0
approx m "=CSCH(1.5)" 4 0.4696
approx m "=SEC(45)" 4 1.9036
error m "=SEC(134217728)" Err.num
approx m "=SECH(45)" 22 5.73e-20
approx m "=SQRTPI(1)" 6 1.772454
approx m "=SQRTPI(2)" 6 2.506628
error m "=SQRTPI(-1)" Err.num
number m "=FACTDOUBLE(6)" 48.0
number m "=FACTDOUBLE(7)" 105.0
number m "=FACTDOUBLE(-1)" 1.0
error m "=FACTDOUBLE(-2)" Err.num
number m "=ARABIC(\"LVII\")" 57.0
number m "=ARABIC(\"mcmxii\")" 1912.0
number m "=ARABIC(\"\")" 0.0
number m "=ARABIC(\"-MMXI\")" -2011.0
number m "=ARABIC(\" CD \")" 400.0
error m "=ARABIC(\"ABC\")" Err.value
text m "=BASE(7,2)" "111"
text m "=BASE(100,16)" "64"
text m "=BASE(15,2,10)" "0000001111"
text m "=BASE(0,2)" "0"
error m "=BASE(-1,2)" Err.num
error m "=BASE(7,37)" Err.num
number m "=DECIMAL(\"FF\",16)" 255.0
number m "=DECIMAL(111,2)" 7.0
number m "=DECIMAL(\"zap\",36)" 45745.0
error m "=DECIMAL(\"12\",2)" Err.num
number m "=CEILING.MATH(24.3,5)" 25.0
number m "=CEILING.MATH(6.7)" 7.0
number m "=CEILING.MATH(-8.1,2)" -8.0
number m "=CEILING.MATH(-5.5,2,-1)" -6.0
number m "=FLOOR.MATH(24.3,5)" 20.0
number m "=FLOOR.MATH(6.7)" 6.0
number m "=FLOOR.MATH(-8.1,2)" -10.0
number m "=FLOOR.MATH(-5.5,2,-1)" -4.0
number m "=CEILING.PRECISE(4.3)" 5.0
number m "=CEILING.PRECISE(-4.3)" -4.0
number m "=CEILING.PRECISE(4.3,2)" 6.0
number m "=CEILING.PRECISE(4.3,-2)" 6.0
number m "=CEILING.PRECISE(-4.3,2)" -4.0
number m "=CEILING.PRECISE(-4.3,-2)" -4.0
number m "=ISO.CEILING(4.3)" 5.0
number m "=ISO.CEILING(-4.3,-2)" -4.0
number m "=FLOOR.PRECISE(-3.2,-1)" -4.0
number m "=FLOOR.PRECISE(3.2,1)" 3.0
number m "=FLOOR.PRECISE(-3.2,1)" -4.0
number m "=FLOOR.PRECISE(3.2,-1)" 3.0
number m "=FLOOR.PRECISE(3.2)" 3.0
number m "=CEILING.MATH(0.3,0.1)" 0.3
number m "=FLOOR.PRECISE(5,0)" 0.0
number m "=COMBINA(4,3)" 20.0
number m "=COMBINA(10,3)" 220.0
number m "=COMBINA(0,0)" 1.0
error m "=COMBINA(-1,2)" Err.num
number m "=MULTINOMIAL(2,3,4)" 1260.0
error m "=MULTINOMIAL(2,-3)" Err.num
eq "MUNIT(3)" (arrayOf m "=MUNIT(3)") [| [| N 1.0; N 0.0; N 0.0 |]; [| N 0.0; N 1.0; N 0.0 |]; [| N 0.0; N 0.0; N 1.0 |] |]
error m "=MUNIT(0)" Err.value
text m "=ROMAN(499,0)" "CDXCIX"
text m "=ROMAN(499,1)" "LDVLIV"
text m "=ROMAN(499,2)" "XDIX"
text m "=ROMAN(499,3)" "VDIV"
text m "=ROMAN(499,4)" "ID"
text m "=ROMAN(2013,0)" "MMXIII"
text m "=ROMAN(499,TRUE)" "CDXCIX"
text m "=ROMAN(499,FALSE)" "ID"
text m "=ROMAN(0)" ""
text m "=ROMAN(1999)" "MCMXCIX"
error m "=ROMAN(4000)" Err.value
error m "=ROMAN(-1)" Err.value
approx m "=SERIESSUM(PI()/4,0,2,{1,-0.5,0.041666667,-0.001388889})" 6 0.707103
error m "=SERIESSUM(1,0,1,{1,\"a\"})" Err.value
number m "=SUMX2MY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})" -55.0
number m "=SUMX2PY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})" 521.0
number m "=SUMXMY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})" 79.0
error m "=SUMXMY2({1,2},{1,2,3})" Err.na
// Lifting: a scalar function over an array.
eq "SQRTPI over an array" (arrayOf m "=SQRTPI({0,1})") [| [| N 0.0; N (sqrt System.Math.PI) |] |]
// From cells: text in a pair is skipped; PERCENTOF sums like SUM.
let cells =
    modelOf [ (0, 0), N 2.0; (1, 0), N 3.0; (2, 0), S "x"; (0, 1), N 6.0; (1, 1), N 5.0; (2, 1), N 1.0
              (0, 2), N 10.0; (1, 2), N 30.0; (2, 2), N 60.0 ]
number cells "=SUMXMY2(A1:A3,B1:B3)" 20.0
number cells "=PERCENTOF(C1,C1:C3)" 0.1
number cells "=PERCENTOF(C1:C2,C1:C3)" 0.4
error cells "=PERCENTOF(C1,D1:D3)" Err.div0

printfn "stage59 math functions: %d passed, %d failed" pass fail
if fail > 0 then exit 1
