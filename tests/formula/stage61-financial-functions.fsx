#load "function-table-kit.fsx"
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open ``Function-table-kit``

// Financial functions of the Financial module (Microsoft's examples; dates by DATE()).
let m = empty
let db period = sprintf "=DB(1000000,100000,6,%d,7)" period
approx m (db 1) 2 186083.33
approx m (db 2) 2 259639.42
approx m (db 3) 2 176814.44
approx m (db 4) 2 120410.64
approx m (db 5) 2 81999.64
approx m (db 6) 2 55841.76
approx m (db 7) 2 15845.10
error m "=DB(1000000,100000,6,8,7)" Err.num
error m "=DB(1000000,100000,6,7)" Err.num
approx m "=DDB(2400,300,10*365,1)" 2 1.32
approx m "=DDB(2400,300,10*12,1,2)" 2 40.0
approx m "=DDB(2400,300,10,1,2)" 2 480.0
approx m "=DDB(2400,300,10,2,1.5)" 2 306.0
approx m "=DDB(2400,300,10,10)" 2 22.12
error m "=DDB(2400,300,10,11)" Err.num
approx m "=SYD(30000,7500,10,1)" 2 4090.91
approx m "=SYD(30000,7500,10,10)" 2 409.09
error m "=SYD(30000,7500,10,11)" Err.num
approx m "=VDB(2400,300,10*365,0,1)" 2 1.32
approx m "=VDB(2400,300,10*12,0,1)" 2 40.0
approx m "=VDB(2400,300,10,0,1)" 2 480.0
approx m "=VDB(2400,300,10*12,6,18)" 2 396.31
approx m "=VDB(2400,300,10*12,6,18,1.5)" 2 311.81
approx m "=VDB(2400,300,10,0,0.875,1.5)" 2 315.0
approx m "=VDB(2400,300,10,0,10)" 6 2100.0
error m "=VDB(2400,300,10,5,4)" Err.num
approx m "=DOLLARDE(1.02,16)" 9 1.125
approx m "=DOLLARDE(1.1,32)" 9 1.3125
approx m "=DOLLARFR(1.125,16)" 9 1.02
approx m "=DOLLARFR(1.125,32)" 9 1.04
error m "=DOLLARDE(1.02,0)" Err.div0
error m "=DOLLARDE(1.02,-1)" Err.num
approx m "=EFFECT(0.0525,4)" 6 0.053543
error m "=EFFECT(0.0525,0.5)" Err.num
approx m "=NOMINAL(0.053543,4)" 8 0.05250032
approx m "=FVSCHEDULE(1,{0.09,0.11,0.1})" 5 1.33089
error m "=FVSCHEDULE(1,{0.09,\"a\"})" Err.value
approx m "=ISPMT(0.1/12,1,3*12,8000000)" 1 -64814.8
approx m "=ISPMT(0.1,1,3,8000000)" 0 -533333.0
error m "=ISPMT(0.1,1,0,8000000)" Err.div0
approx m "=PDURATION(2.5%,2000,2200)" 2 3.86
approx m "=PDURATION(0.025/12,1000,1200)" 1 87.6
error m "=PDURATION(0,1000,1200)" Err.num
approx m "=RRI(96,10000,11000)" 7 0.0009933
error m "=RRI(0,10000,11000)" Err.num
let s, mat = "DATE(2008,3,31)", "DATE(2008,6,1)"
approx m (sprintf "=TBILLEQ(%s,%s,9.14%%)" s mat) 6 0.094151
approx m (sprintf "=TBILLPRICE(%s,%s,9%%)" s mat) 2 98.45
approx m (sprintf "=TBILLYIELD(%s,%s,98.45)" s mat) 6 0.091417
error m (sprintf "=TBILLPRICE(%s,%s,9%%)" mat s) Err.num
error m "=TBILLPRICE(DATE(2008,1,1),DATE(2009,6,1),9%)" Err.num
// A bill of more than half a year (274 days): the bond-equivalent yield y of price P, from
// P(1 + y/2)(1 + (t - 1/2)y) = 100, t = 274/365, worked out by hand (Microsoft gives no example).
approx m "=TBILLEQ(DATE(2008,1,1),DATE(2008,10,1),9%)" 6 0.096408
approx m "=DISC(DATE(2018,1,25),DATE(2018,6,15),97.975,100,1)" 9 0.052420213
approx m "=INTRATE(DATE(2008,2,15),DATE(2008,5,15),1000000,1014420,2)" 5 0.05768
approx m "=RECEIVED(DATE(2008,2,15),DATE(2008,5,15),1000000,5.75%,2)" 2 1014584.65
approx m "=PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),5.25%,100,2)" 5 99.79583
approx m "=YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,100,2)" 6 0.052823
approx m "=PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),6.1%,6.1%,0)" 8 99.98449888
approx m "=YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),6.25%,100.0123,0)" 6 0.060954
approx m "=ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),10%,1000,3)" 8 20.54794521
error m "=DISC(DATE(2018,6,15),DATE(2018,1,25),97.975,100,1)" Err.num
error m "=DISC(DATE(2018,1,25),DATE(2018,6,15),97.975,100,5)" Err.num
error m "=ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0,1000,3)" Err.num

printfn "stage61 financial functions: %d passed, %d failed" pass fail
if fail > 0 then exit 1
