#load "function-table-kit.fsx"
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open ``Function-table-kit``

// The normal distribution (NormalDistribution.fs and its callers), BAHTTEXT, EUROCONVERT, FILTERXML
// and the argument rules of FORECAST.ETS and its kin, against Excel 2019: every expected value below
// is Excel's own, taken from tools/excel-parity/{normal,bahttext,euro,filterxml,ets}-excel.tsv (the
// probes of make-*-probe.py; the full comparison is tools/excel-parity/function-parity-engine.fsx).
// The normal distribution to 4e-15 relative (Excel's own digits where they are reproduced; its far
// tail is), the rest exactly. FORECAST.ETS's fitted values are not Excel's (its choice of the
// smoothing parameters is not reproduced), so only its rules and the cases that do not depend on
// that choice are here.
let m = empty

let rel (m: SpreadsheetModel) (formula: string) (expected: float) (tolerance: float) =
    match value m formula with
    | CellValue.Number x when (if expected = 0.0 then abs x <= tolerance else abs (x - expected) <= tolerance * abs expected) -> pass <- pass + 1
    | other -> (fail <- fail + 1; printfn "FAIL rel %s expected %.17g  got %A" formula expected other)

// --- FORECAST.ETS: the rules ---
error m "=FORECAST.ETS(0,{1,2,3,4,5,6},{1,2,3,4,5,6})" "NUM!"
error m "=FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,3,4,5})" "N/A"
error m "=FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,3,4,5,6},-1)" "NUM!"
error m "=FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,3,4,5,6},2.7)" "NUM!"
error m "=FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,3,4,5,6},0,2)" "NUM!"
error m "=FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,3,4,5,6},0,1,0)" "NUM!"
error m "=FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,3,4,5,6},0,1,8)" "NUM!"
error m "=FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,3,4,5,6},\"x\")" "VALUE!"
error m "=FORECAST.ETS(\"x\",{1,2,3,4,5,6},{1,2,3,4,5,6})" "VALUE!"
error m "=FORECAST.ETS.STAT({1,2,3,4,5,6},{1,2,3,4,5,6},9)" "NUM!"
error m "=FORECAST.ETS.STAT({1,2,3,4,5,6},{1,2,3,4,5,6},0)" "NUM!"
error m "=FORECAST.ETS.CONFINT(7,{1,2,3,4,5,6},{1,2,3,4,5,6},1)" "NUM!"
rel m "=FORECAST.ETS(3,{1,2,3,4,5,6},{1,2,3,4,5,6})" 3.0 0.0
rel m "=FORECAST.ETS(6,{1,2,3,4,5,6},{1,2,3,4,5,6})" 6.0 0.0
rel m "=FORECAST.ETS(7,{1,2,3},{1,2,3})" 7.0 1e-12
rel m "=FORECAST.ETS(7,{1,2},{1,2})" 7.0 1e-12
rel m "=FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,3,4,5,6},8761)" 7.0 1e-12
rel m "=FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,3,4,5,6},4)" 7.0 1e-12
rel m "=FORECAST.ETS(7,{1,2,\"a\",4,5,6},{1,2,3,4,5,6})" 7.0 1e-12
rel m "=FORECAST.ETS(7,{1,2,3,4,5,6},{1,2,\"a\",4,5,6})" 7.0 1e-12
rel m "=FORECAST.ETS(7,{5,5,5,5,5,5},{1,2,3,4,5,6})" 5.0 1e-12
rel m "=FORECAST.ETS.STAT({5,5,5,5,5,5},{1,2,3,4,5,6},4)" 0.0 0.0
rel m "=FORECAST.ETS.CONFINT(7,{1,2,3,4,5,6},{1,2,3,4,5,6},0)" 0.0 0.0
rel m "=FORECAST.ETS.CONFINT(7,{1,2,3,4,5,6},{1,2,3,4,5,6})" 0.0 1e-12
rel m "=FORECAST.ETS.SEASONALITY({1,2,3,4,5,6},{1,2,3,4,5,6})" 0.0 0.0
rel m "=FORECAST.ETS.STAT({12,15,14,18,21,19,25,27,26,30,33,31,36,38,37,41},{0,10,20,30,40,50,60,70,80,90,100,110,120,130,140,150},8,0)" 10.0 0.0
rel m "=FORECAST.ETS.STAT({12,15,14,18,19,25,27,33,31,36,38,37,41},{1,2,3,4,6,7,8,11,12,13,14,15,16},3,0,1)" 2.220446049250313e-16 0.0
rel m "=FORECAST.ETS.SEASONALITY({12,15,14,18,21,19,25,27,26,30,33,31,36,38,37,41},{0,0.5,1,1.5,2,2.5,3,3.5,4,4.5,5,5.5,6,6.5,7,7.5},0)" 3.0 0.0

// --- normal distribution ---
rel m "=NORM.S.DIST(-40.0,TRUE)" 0.0 4e-15
rel m "=NORM.S.DIST(-34.25,TRUE)" 2.180693926866097e-257 4e-15
rel m "=NORM.S.DIST(-28.5,TRUE)" 5.857141253805672e-179 4e-15
rel m "=NORM.S.DIST(-22.75,TRUE)" 7.17504567166862e-115 4e-15
rel m "=NORM.S.DIST(-17.0,TRUE)" 4.10599620209879e-65 4e-15
rel m "=NORM.S.DIST(-11.25,TRUE)" 1.1579603185686336e-29 4e-15
rel m "=NORM.S.DIST(-5.5,TRUE)" 1.898956246588769e-08 4e-15
rel m "=NORM.S.DIST(0.25,TRUE)" 0.5987063256829237 4e-15
rel m "=NORM.S.DIST(6.0,TRUE)" 0.9999999990134123 4e-15
rel m "=NORM.S.DIST(11.75,TRUE)" 1.0 4e-15
rel m "=NORM.S.DIST(17.5,TRUE)" 1.0 4e-15
rel m "=NORM.S.DIST(23.25,TRUE)" 1.0 4e-15
rel m "=NORM.S.DIST(29.0,TRUE)" 1.0 4e-15
rel m "=NORM.S.DIST(34.75,TRUE)" 1.0 4e-15
rel m "=NORM.S.DIST(-37.025,TRUE)" 2.2681292127220604e-300 4e-15
rel m "=NORM.S.DIST(-37.6,TRUE)" 0.0 4e-15
rel m "=NORM.S.DIST(-38.175,TRUE)" 0.0 4e-15
rel m "=NORM.S.DIST(-7.6,TRUE)" 1.4806537490048066e-14 4e-15
rel m "=NORM.S.DIST(-3.3699105822640405,TRUE)" 0.00037596289588385243 4e-15
rel m "=NORM.S.DIST(-29.474833275981492,TRUE)" 3.025968939371348e-191 4e-15
rel m "=NORM.S.DIST(-23.972709376257466,TRUE)" 2.678674876039552e-127 4e-15
rel m "=NORM.S.DIST(-7.066850454027897,TRUE)" 7.924479407503459e-13 4e-15
rel m "=NORM.S.DIST(-29.69293839954532,TRUE)" 4.736193919627217e-194 4e-15
rel m "=NORMSDIST(-35.5)" 2.4576915406619364e-276 4e-15
rel m "=NORMSDIST(-18.25)" 1.0345463677569806e-74 4e-15
rel m "=NORMSDIST(-1.0)" 0.158655253931457 4e-15
rel m "=NORMSDIST(16.25)" 1.0 4e-15
rel m "=NORMSDIST(33.5)" 1.0 4e-15
rel m "=NORMSDIST(-38.05)" 0.0 4e-15
rel m "=NORMSDIST(-17.844563344142976)" 1.5929180241065146e-71 4e-15
rel m "=NORMSDIST(6.773198007146874)" 0.9999999999937017 4e-15
rel m "=NORM.S.DIST(-32.5,FALSE)" 1.7343502636064355e-230 4e-15
rel m "=NORM.S.DIST(-21.0,FALSE)" 6.902029420127219e-97 4e-15
rel m "=NORM.S.DIST(-9.5,FALSE)" 1.007793539430001e-20 4e-15
rel m "=NORM.S.DIST(2.0,FALSE)" 0.05399096651318806 4e-15
rel m "=NORM.S.DIST(13.5,FALSE)" 1.061268813915216e-40 4e-15
rel m "=NORM.S.DIST(25.0,FALSE)" 7.653929736419393e-137 4e-15
rel m "=NORM.S.DIST(36.5,FALSE)" 2.0253417162117463e-290 4e-15
rel m "=NORM.S.DIST(37.875,FALSE)" 0.0 4e-15
rel m "=NORM.S.DIST(38.45,FALSE)" 0.0 4e-15
rel m "=NORM.S.DIST(-2.4627855474555105,FALSE)" 0.01922401770464755 4e-15
rel m "=NORM.S.DIST(27.42319242005476,FALSE)" 1.992271538032958e-164 4e-15
rel m "=NORM.S.INV(1e-17)" -8.4937932241096 4e-15
rel m "=NORM.S.INV(1e-40)" -13.310921371425172 4e-15
rel m "=NORM.S.INV(1e-63)" -16.811801442450445 4e-15
rel m "=NORM.S.INV(1e-86)" -19.703807157156703 4e-15
rel m "=NORM.S.INV(1e-109)" -22.224294638257714 4e-15
rel m "=NORM.S.INV(1e-132)" -24.487648483987773 4e-15
rel m "=NORM.S.INV(1e-155)" -26.559403396043333 4e-15
rel m "=NORM.S.INV(1e-178)" -28.481247624839572 4e-15
rel m "=NORM.S.INV(1e-201)" -30.281645787437245 4e-15
rel m "=NORM.S.INV(1e-224)" -31.98105699308846 4e-15
rel m "=NORM.S.INV(1e-247)" -33.59477116781466 4e-15
rel m "=NORM.S.INV(1e-270)" -35.13457108169203 4e-15
rel m "=NORM.S.INV(1e-293)" -36.60976465333901 4e-15
rel m "=NORM.S.INV(0.739531753863005)" 0.6419024979030127 4e-15
rel m "=NORM.S.INV(0.6056333637960464)" 0.2679558896467113 4e-15
rel m "=NORM.S.INV(0.09076571618896534)" -1.336054585641915 4e-15
rel m "=NORM.S.INV(0.5302009300704787)" 0.07577495748876566 4e-15
rel m "=NORM.S.INV(0.03049807258352343)" -1.8735233955615849 4e-15
rel m "=NORM.S.INV(0.99999999999)" 6.706023143414748 4e-15
rel m "=NORMSINV(1e-41)" -13.481860193554807 4e-15
rel m "=NORMSINV(1e-133)" -24.58134426228077 4e-15
rel m "=NORMSINV(1e-225)" -32.05290466082274 4e-15
rel m "=NORMSINV(0.37954885458836685)" -0.30666587186093575 4e-15
rel m "=NORMSINV(0.002063585606060855)" -2.868274626116897 4e-15
rel m "=NORMINV(0.9999999999999,14.154360190945098,0.5)" 17.82873746109502 4e-15
rel m "=NORMDIST(0.4165904948787684,-15.266589810528286,2.5,TRUE)" 0.9999999998232308 4e-15
rel m "=NORM.DIST(-16.106669073318805,7.029814151349811,1,TRUE)" 9.944198555891915e-119 4e-15
rel m "=NORM.INV(0.37047693110563606,-9.877374564652325,2.5)" -10.703850684942989 4e-15
rel m "=NORM.DIST(23.98207480922939,7.495922773790738,2.5,FALSE)" 5.7531644525854774e-11 4e-15
rel m "=NORMINV(0.4084573169925092,-14.456181447547166,2.5)" -15.034969268255898 4e-15
rel m "=NORMDIST(18.056988910675102,9.439218754650529,10,TRUE)" 0.8055948823593084 4e-15
rel m "=NORM.DIST(21.423566067437264,2.2443398417716587,10,TRUE)" 0.9724395883772222 4e-15
rel m "=NORM.INV(0.4766335535791324,13.305333075829964,10)" 12.719287834707202 4e-15
rel m "=NORM.DIST(49.472493497914556,-7.776251569738539,10,FALSE)" 3.0484953823683596e-09 4e-15
rel m "=NORMINV(1e-298,13.161335637315425,10)" -356.06538636724133 4e-15
rel m "=NORMDIST(34.32893022669592,9.842103087134678,0.5,FALSE)" 0.0 4e-15
rel m "=NORM.DIST(-24.503764243755587,-0.7951746010054634,0.5,TRUE)" 0.0 4e-15
rel m "=PHI(-37.0)" 2.1200065515246056e-298 4e-15
rel m "=GAUSS(-31.0)" -0.5 4e-15
rel m "=PHI(-25.5)" 2.5171937051934048e-142 4e-15
rel m "=GAUSS(-19.5)" -0.5 4e-15
rel m "=PHI(-14.0)" 1.0966065593889713e-43 4e-15
rel m "=GAUSS(-8.0)" -0.4999999999999994 4e-15
rel m "=PHI(-2.5)" 0.01752830049356854 4e-15
rel m "=GAUSS(3.5)" 0.49976737092096446 4e-15
rel m "=PHI(9.0)" 1.0279773571668917e-18 4e-15
rel m "=GAUSS(15.0)" 0.5 4e-15
rel m "=PHI(20.5)" 2.2119843802105703e-92 4e-15
rel m "=GAUSS(26.5)" 0.5 4e-15
rel m "=PHI(32.0)" 1.746366256758777e-223 4e-15
rel m "=GAUSS(38.0)" 0.5 4e-15
rel m "=PHI(18.315195149691675)" 5.7498378977220495e-74 4e-15
rel m "=GAUSS(-3.8472460318315456)" -0.49994027350061226 4e-15
rel m "=Z.TEST({-1.677,0.165},0.9139510279870393,2.4517367164462116)" 0.8322922348964825 4e-15
rel m "=Z.TEST({-2.935,3.48,2.361,-0.186,-4.117},-1.8479176029481366)" 0.1420724220421192 4e-15
rel m "=ZTEST({-3.664,-2.496,4.469,-3.497,3.519,-4.767,3.645,-2.896},0.5883034524451802)" 0.8291031579363639 4e-15
error m "=Z.TEST({2.255},30.320273251449464,0.4169505245926378)" "DIV/0!"
rel m "=Z.TEST({-2.572,1.461,-1.292,0.421,4.168,-2.473,-1.704},-56.13397813015408)" 0.0 4e-15
error m "=CONFIDENCE.NORM(1e-217,4.124744438879403,7.9)" "NUM!"
error m "=CONFIDENCE(1e-267,3.4927500204888453,7.9)" "NUM!"
rel m "=CONFIDENCE.NORM(0.99999999999,3.8271968892559105,7.9)" 1.8129747661519034e-11 4e-15
rel m "=LOGNORM.INV(0.6118184045675533,2.9519884278358584,1.6444986936214256)" 30.542749397106824 4e-15
rel m "=LOGNORM.DIST(1e-42,-1.1363148754421428,0.20957458961741796,FALSE)" 0.0 4e-15
rel m "=LOGINV(1e-80,2.0098456268869436,1.1156923570001775)" 4.684637166483862e-09 4e-15
rel m "=LOGNORMDIST(264225.1842397521,1.9885288178552152,1.3856348192633403)" 0.999999999999982 4e-15
rel m "=LOGNORM.DIST(4.8823173653602945,2.0066492123732447,1.5146962233695793,TRUE)" 0.3905204494921468 4e-15
rel m "=LOGNORM.INV(1e-171,2.611985590300371,0.6453943781207911)" 2.0481172794126408e-07 4e-15
rel m "=LOGNORM.DIST(1e-208,1.2169794064091661,1.2064484159315814,FALSE)" 0.0 4e-15
rel m "=LOGINV(1e-273,0.8552630909460213,2.1917727774733695)" 5.511198042916273e-34 4e-15
rel m "=LOGNORMDIST(4.562852843047508,-2.7625123088901167,2.0173734087212876)" 0.9830726788194966 4e-15
rel m "=LOGNORM.DIST(1e-86,1.232920617518884,1.803201796367076,TRUE)" 0.0 4e-15
rel m "=LOGNORM.INV(0.9,1.3477503581833208,1.8127974339729251)" 39.28896724198412 4e-15
rel m "=LOGINV(1e-22,-1.7554382837306062,2.1179479071107115)" 1.8923302851799763e-10 4e-15
error m "=NORM.S.INV(-0.1)" "NUM!"
error m "=CONFIDENCE.NORM(1,1,1)" "NUM!"
// --- BAHTTEXT ---
text m "=BAHTTEXT(0)" "ศูนย์บาทถ้วน"
text m "=BAHTTEXT(9)" "เก้าบาทถ้วน"
text m "=BAHTTEXT(18)" "สิบแปดบาทถ้วน"
text m "=BAHTTEXT(27)" "ยี่สิบเจ็ดบาทถ้วน"
text m "=BAHTTEXT(36)" "สามสิบหกบาทถ้วน"
text m "=BAHTTEXT(45)" "สี่สิบห้าบาทถ้วน"
text m "=BAHTTEXT(54)" "ห้าสิบสี่บาทถ้วน"
text m "=BAHTTEXT(63)" "หกสิบสามบาทถ้วน"
text m "=BAHTTEXT(72)" "เจ็ดสิบสองบาทถ้วน"
text m "=BAHTTEXT(81)" "แปดสิบเอ็ดบาทถ้วน"
text m "=BAHTTEXT(90)" "เก้าสิบบาทถ้วน"
text m "=BAHTTEXT(99)" "เก้าสิบเก้าบาทถ้วน"
text m "=BAHTTEXT(108)" "หนึ่งร้อยแปดบาทถ้วน"
text m "=BAHTTEXT(117)" "หนึ่งร้อยสิบเจ็ดบาทถ้วน"
text m "=BAHTTEXT(126)" "หนึ่งร้อยยี่สิบหกบาทถ้วน"
text m "=BAHTTEXT(999)" "เก้าร้อยเก้าสิบเก้าบาทถ้วน"
text m "=BAHTTEXT(10001)" "หนึ่งหมื่นเอ็ดบาทถ้วน"
text m "=BAHTTEXT(1010101)" "หนึ่งล้านหนึ่งหมื่นหนึ่งร้อยเอ็ดบาทถ้วน"
text m "=BAHTTEXT(1000000000001)" "หนึ่งล้านล้านเอ็ดบาทถ้วน"
text m "=BAHTTEXT(1.5e+21)" "หนึ่งพันห้าร้อยล้านล้านล้านบาทถ้วน"
text m "=BAHTTEXT(263001714)" "สองร้อยหกสิบสามล้านหนึ่งพันเจ็ดร้อยสิบสี่บาทถ้วน"
text m "=BAHTTEXT(95028)" "เก้าหมื่นห้าพันยี่สิบแปดบาทถ้วน"
text m "=BAHTTEXT(5)" "ห้าบาทถ้วน"
text m "=BAHTTEXT(15)" "สิบห้าบาทถ้วน"
text m "=BAHTTEXT(75)" "เจ็ดสิบห้าบาทถ้วน"
text m "=BAHTTEXT(76990565010499)" "เจ็ดสิบหกล้านเก้าแสนเก้าหมื่นห้าร้อยหกสิบห้าล้านหนึ่งหมื่นสี่ร้อยเก้าสิบเก้าบาทถ้วน"
text m "=BAHTTEXT(591490)" "ห้าแสนเก้าหมื่นหนึ่งพันสี่ร้อยเก้าสิบบาทถ้วน"
text m "=BAHTTEXT(78149)" "เจ็ดหมื่นแปดพันหนึ่งร้อยสี่สิบเก้าบาทถ้วน"
text m "=BAHTTEXT(95772833740356)" "เก้าสิบห้าล้านเจ็ดแสนเจ็ดหมื่นสองพันแปดร้อยสามสิบสามล้านเจ็ดแสนสี่หมื่นสามร้อยห้าสิบหกบาทถ้วน"
text m "=BAHTTEXT(16137846381)" "หนึ่งหมื่นหกพันหนึ่งร้อยสามสิบเจ็ดล้านแปดแสนสี่หมื่นหกพันสามร้อยแปดสิบเอ็ดบาทถ้วน"
text m "=BAHTTEXT(49337423553334)" "สี่สิบเก้าล้านสามแสนสามหมื่นเจ็ดพันสี่ร้อยยี่สิบสามล้านห้าแสนห้าหมื่นสามพันสามร้อยสามสิบสี่บาทถ้วน"
text m "=BAHTTEXT(53419889893198)" "ห้าสิบสามล้านสี่แสนหนึ่งหมื่นเก้าพันแปดร้อยแปดสิบเก้าล้านแปดแสนเก้าหมื่นสามพันหนึ่งร้อยเก้าสิบแปดบาทถ้วน"
text m "=BAHTTEXT(62464)" "หกหมื่นสองพันสี่ร้อยหกสิบสี่บาทถ้วน"
text m "=BAHTTEXT(4680237901)" "สี่พันหกร้อยแปดสิบล้านสองแสนสามหมื่นเจ็ดพันเก้าร้อยเอ็ดบาทถ้วน"
text m "=BAHTTEXT(21.1)" "ยี่สิบเอ็ดบาทสิบสตางค์"
text m "=BAHTTEXT(1000000000.11)" "หนึ่งพันล้านบาทสิบเอ็ดสตางค์"
text m "=BAHTTEXT(21.25)" "ยี่สิบเอ็ดบาทยี่สิบห้าสตางค์"
text m "=BAHTTEXT(1000000000.5)" "หนึ่งพันล้านบาทห้าสิบสตางค์"
text m "=BAHTTEXT(21.999)" "ยี่สิบสองบาทถ้วน"
text m "=BAHTTEXT(1000000000.995)" "หนึ่งพันล้านเอ็ดบาทถ้วน"
text m "=BAHTTEXT(21.005)" "ยี่สิบเอ็ดบาทหนึ่งสตางค์"
text m "=BAHTTEXT(1000000000.004)" "หนึ่งพันล้านบาทถ้วน"
text m "=BAHTTEXT(21.015)" "ยี่สิบเอ็ดบาทสองสตางค์"
text m "=BAHTTEXT(1000000000.025)" "หนึ่งพันล้านบาทสามสตางค์"
text m "=BAHTTEXT(21.045)" "ยี่สิบเอ็ดบาทห้าสตางค์"
text m "=BAHTTEXT(1000000000.105)" "หนึ่งพันล้านบาทสิบเอ็ดสตางค์"
text m "=BAHTTEXT(21.135)" "ยี่สิบเอ็ดบาทสิบสี่สตางค์"
text m "=BAHTTEXT(1000000000.145)" "หนึ่งพันล้านบาทสิบห้าสตางค์"
text m "=BAHTTEXT(21.165)" "ยี่สิบเอ็ดบาทสิบเจ็ดสตางค์"
text m "=BAHTTEXT(1000000000.175)" "หนึ่งพันล้านบาทสิบแปดสตางค์"
text m "=BAHTTEXT(21.195)" "ยี่สิบเอ็ดบาทยี่สิบสตางค์"
text m "=BAHTTEXT(1000000000.205)" "หนึ่งพันล้านบาทยี่สิบเอ็ดสตางค์"
text m "=BAHTTEXT(21.285)" "ยี่สิบเอ็ดบาทยี่สิบเก้าสตางค์"
text m "=BAHTTEXT(1000000000.565)" "หนึ่งพันล้านบาทห้าสิบเจ็ดสตางค์"
text m "=BAHTTEXT(21.585)" "ยี่สิบเอ็ดบาทห้าสิบเก้าสตางค์"
text m "=BAHTTEXT(1000000000.675)" "หนึ่งพันล้านบาทหกสิบแปดสตางค์"
text m "=BAHTTEXT(21.333333333333332)" "ยี่สิบเอ็ดบาทสามสิบสามสตางค์"
text m "=BAHTTEXT(1000000000.6666666)" "หนึ่งพันล้านบาทหกสิบเจ็ดสตางค์"
text m "=BAHTTEXT(328.4)" "สามร้อยยี่สิบแปดบาทสี่สิบสตางค์"
text m "=BAHTTEXT(0.805)" "แปดสิบเอ็ดสตางค์"
text m "=BAHTTEXT(646.4487)" "หกร้อยสี่สิบหกบาทสี่สิบห้าสตางค์"
text m "=BAHTTEXT(244982366567.76)" "สองแสนสี่หมื่นสี่พันเก้าร้อยแปดสิบสองล้านสามแสนหกหมื่นหกพันห้าร้อยหกสิบเจ็ดบาทเจ็ดสิบหกสตางค์"
text m "=BAHTTEXT(39289.2192)" "สามหมื่นเก้าพันสองร้อยแปดสิบเก้าบาทยี่สิบสองสตางค์"
text m "=BAHTTEXT(87484.97)" "แปดหมื่นเจ็ดพันสี่ร้อยแปดสิบสี่บาทเก้าสิบเจ็ดสตางค์"
text m "=BAHTTEXT(6148.4)" "หกพันหนึ่งร้อยสี่สิบแปดบาทสี่สิบสตางค์"
text m "=BAHTTEXT(82449.12)" "แปดหมื่นสองพันสี่ร้อยสี่สิบเก้าบาทสิบสองสตางค์"
text m "=BAHTTEXT(-1)" "ลบหนึ่งบาทถ้วน"
text m "=BAHTTEXT(\"12\")" "สิบสองบาทถ้วน"
text m "=BAHTTEXT(\"1,000\")" "หนึ่งพันบาทถ้วน"
// --- EUROCONVERT ---
rel m "=EUROCONVERT(1,\"BEF\",\"EUR\")" 0.02 0.0
rel m "=EUROCONVERT(1,\"EUR\",\"BEF\",TRUE)" 40.3399 0.0
rel m "=EUROCONVERT(1,\"LUF\",\"EUR\",TRUE)" 0.0247893524773239 0.0
rel m "=EUROCONVERT(-55.555,\"EUR\",\"DEM\")" -108.66 0.0
rel m "=EUROCONVERT(-55.555,\"ESP\",\"EUR\")" -0.33 0.0
rel m "=EUROCONVERT(1000000000.0,\"EUR\",\"FRF\")" 6559570000.0 0.0
rel m "=EUROCONVERT(1000000000.0,\"IEP\",\"EUR\")" 1269738078.43 0.0
rel m "=EUROCONVERT(2.675,\"EUR\",\"ITL\")" 5180.0 0.0
rel m "=EUROCONVERT(2.675,\"NLG\",\"EUR\")" 1.21 0.0
rel m "=EUROCONVERT(0.015,\"EUR\",\"ATS\")" 0.21 0.0
rel m "=EUROCONVERT(0.015,\"PTE\",\"EUR\")" 0.0 0.0
rel m "=EUROCONVERT(0.005,\"EUR\",\"FIM\")" 0.03 0.0
rel m "=EUROCONVERT(0.005,\"GRD\",\"EUR\")" 0.0 0.0
rel m "=EUROCONVERT(1234.5678,\"EUR\",\"SIT\")" 295851.83 0.0
rel m "=EUROCONVERT(1234.5678,\"CYP\",\"EUR\")" 2109.38 0.0
rel m "=EUROCONVERT(100,\"EUR\",\"MTL\")" 42.93 0.0
rel m "=EUROCONVERT(100,\"SKK\",\"EUR\")" 3.32 0.0
rel m "=EUROCONVERT(1,\"EUR\",\"EEK\")" 15.65 0.0
error m "=EUROCONVERT(1,\"LVL\",\"EUR\")" "VALUE!"
error m "=EUROCONVERT(1,\"EUR\",\"LVL\",TRUE)" "VALUE!"
error m "=EUROCONVERT(1,\"LTL\",\"EUR\",TRUE)" "VALUE!"
error m "=EUROCONVERT(-55.555,\"EUR\",\"HRK\")" "VALUE!"
rel m "=EUROCONVERT(-55.555,\"EUR\",\"EUR\")" -55.555 0.0
rel m "=EUROCONVERT(1,\"ATS\",\"LUF\",TRUE,5)" 2.931500533 0.0
rel m "=EUROCONVERT(1000,\"SIT\",\"NLG\",TRUE,6)" 9.195918755460001 0.0
error m "=EUROCONVERT(1000,\"FRF\",\"LUF\",TRUE,1)" "VALUE!"
error m "=EUROCONVERT(100,\"NLG\",\"ITL\",FALSE,2)" "VALUE!"
rel m "=EUROCONVERT(1,\"DEM\",\"LUF\",TRUE,5)" 20.625387471 0.0
rel m "=EUROCONVERT(74355.94746341245,\"ESP\",\"DEM\",FALSE,3)" 874.04 0.0
rel m "=EUROCONVERT(787.38,\"BEF\",\"DEM\",TRUE,3)" 38.175845769999995 0.0
rel m "=EUROCONVERT(100,\"DEM\",\"IEP\",FALSE,4)" 40.27 0.0
rel m "=EUROCONVERT(617.23,\"BEF\",\"DEM\")" 29.93 0.0
rel m "=EUROCONVERT(4711.511073865756,\"FRF\",\"EUR\",FALSE)" 718.27 0.0
rel m "=EUROCONVERT(9341.214925567312,\"EUR\",\"FRF\",FALSE)" 61274.35 0.0
rel m "=EUROCONVERT(7775.472333124697,\"EUR\",\"PTE\",FALSE,3)" 1558842.0 0.0
rel m "=EUROCONVERT(1423.444374683077,\"GRD\",\"EUR\",FALSE,4)" 4.18 0.0
rel m "=EUROCONVERT(8237.374758045144,\"ESP\",\"EUR\",TRUE)" 49.5076193793056 0.0
rel m "=EUROCONVERT(2161.8071591629673,\"EUR\",\"IEP\",TRUE)" 1702.5614934990176 0.0
rel m "=EUROCONVERT(6209.887688221433,\"EUR\",\"GRD\",TRUE,3)" 2116019.2297614524 0.0
rel m "=EUROCONVERT(2082.7460182841705,\"GRD\",\"EUR\",TRUE,4)" 6.1122 0.0
rel m "=EUROCONVERT(3396.24478218815,\"BEF\",\"EUR\",FALSE)" 84.19 0.0
rel m "=EUROCONVERT(7814.630018454133,\"EUR\",\"FRF\",FALSE)" 51260.61 0.0
rel m "=EUROCONVERT(2.365684299814319,\"EUR\",\"IEP\",FALSE,3)" 1.86 0.0
rel m "=EUROCONVERT(3522.682912366081,\"NLG\",\"EUR\",FALSE,4)" 1598.52 0.0
rel m "=EUROCONVERT(5743.376146119081,\"ESP\",\"EUR\",TRUE)" 34.5183858384665 0.0
rel m "=EUROCONVERT(1092.423339705587,\"EUR\",\"NLG\",TRUE)" 2407.3842379425837 0.0
rel m "=EUROCONVERT(1874.9293397133104,\"EUR\",\"DEM\",TRUE,3)" 3667.043050491483 0.0
rel m "=EUROCONVERT(7582.918738421552,\"FRF\",\"EUR\",TRUE,4)" 1156.0085 0.0
rel m "=EUROCONVERT(7810.461255102347,\"BEF\",\"EUR\",FALSE)" 193.62 0.0
rel m "=EUROCONVERT(3707.9310520089934,\"EUR\",\"IEP\",FALSE)" 2920.23 0.0
rel m "=EUROCONVERT(2527.1866416163557,\"EUR\",\"DEM\",FALSE,3)" 4942.75 0.0
rel m "=EUROCONVERT(322.1684891538268,\"FIM\",\"EUR\",FALSE,4)" 54.18 0.0
rel m "=EUROCONVERT(2565.77425531727,\"DEM\",\"EUR\",TRUE)" 1311.85954572599 0.0
rel m "=EUROCONVERT(7539.682447235115,\"EUR\",\"DEM\",TRUE)" 14746.337120775845 0.0
rel m "=EUROCONVERT(2440.838560932936,\"EUR\",\"DEM\",TRUE,3)" 4773.865282629452 0.0
rel m "=EUROCONVERT(8124.1552753227315,\"BEF\",\"EUR\",TRUE,4)" 201.3925 0.0
rel m "=EUROCONVERT(31.350291916204533,\"FIM\",\"EUR\",FALSE)" 5.27 0.0
rel m "=EUROCONVERT(8241.75501086002,\"EUR\",\"FRF\",FALSE)" 54062.37 0.0
rel m "=EUROCONVERT(5461.044307204323,\"EUR\",\"NLG\",FALSE,3)" 12034.56 0.0
rel m "=EUROCONVERT(1372.4758192751085,\"DEM\",\"EUR\",FALSE,4)" 701.74 0.0
rel m "=EUROCONVERT(1,\"DEM\",\"FRF\",1)" 3.353854885138276 0.0
rel m "=EUROCONVERT(-0.004,\"DEM\",\"EUR\")" 0.0 0.0
rel m "=EUROCONVERT(0.5,\"DEM\",\"DEM\")" 0.5 0.0
rel m "=EUROCONVERT(1234.565,\"BEF\",\"BEF\")" 1234.565 0.0
rel m "=EUROCONVERT(1.95583*3/200,\"DEM\",\"EUR\")" 0.02 0.0
rel m "=EUROCONVERT(1936.27*5/2000,\"ITL\",\"NLG\",TRUE,3)" 0.006611130000000001 0.0
rel m "=EUROCONVERT(1.95583*125/2000,\"DEM\",\"EUR\",TRUE,3)" 0.063 0.0
rel m "=EUROCONVERT(2675/2000,\"EUR\",\"FRF\",TRUE,3)" 8.773424875 0.0
// --- FILTERXML ---
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a\")" 10.0 0.0
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a\"),2)" 2.5 0.0
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"/r/a\")" 10.0 0.0
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"/r/a\"),2)" 2.5 0.0
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a/@id\")" 1.0 0.0
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a/@id\"),2)" 2.0 0.0
text m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//@n\")" "x"
text m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//@n\"),2)" "y"
text m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id>2]\")" "abc"
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id>2]\"),2)" -7.0 0.0
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[2]\")" 2.5 0.0
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[2]\"),2)" "REF!"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"(//a)[2]\")" 2.5 0.0
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"(//a)[2]\"),2)" "REF!"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[last()]\")" -7.0 0.0
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[last()]\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"(//a)[last()]\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"(//a)[last()]\"),2)" "VALUE!"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[position()<3]\")" 10.0 0.0
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[position()<3]\"),2)" 2.5 0.0
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//b/a\")" -7.0 0.0
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//b/a\"),2)" "REF!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//b\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//b\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"/r\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"/r\"),2)" "VALUE!"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//c|//a[1]\")" 10.0 0.0
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//c|//a[1]\"),2)" -7.0 0.0
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@n]\")" 10.0 0.0
text m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@n]\"),2)" "abc"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[not(@n)]\")" 2.5 0.0
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[not(@n)]\"),2)" -7.0 0.0
text m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@n=\"\"y\"\"]\")" "abc"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@n='x']/@id\")" 1.0 0.0
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@n='x']/@id\"),2)" "REF!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a/text()\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a/text()\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//b//text()\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//b//text()\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//*\")" "VALUE!"
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//*\"),2)" 10.0 0.0
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"/*/*\")" 10.0 0.0
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"/*/*\"),2)" 2.5 0.0
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[.=10]\")" 10.0 0.0
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[.=10]\"),2)" "REF!"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[. > 0]\")" 10.0 0.0
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[. > 0]\"),2)" 2.5 0.0
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//c/..\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//c/..\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//c/parent::*\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//c/parent::*\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//c/ancestor::*\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//c/ancestor::*\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//c/ancestor-or-self::*\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//c/ancestor-or-self::*\"),2)" "VALUE!"
text m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id=2]/following-sibling::*\")" "abc"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id=2]/following-sibling::*\"),2)" "VALUE!"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id=3]/preceding-sibling::a\")" 10.0 0.0
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id=3]/preceding-sibling::a\"),2)" 2.5 0.0
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//b/following::a\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//b/following::a\"),2)" "VALUE!"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//c/preceding::a\")" 10.0 0.0
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//c/preceding::a\"),2)" 2.5 0.0
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//b/descendant::*\")" -7.0 0.0
text m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//b/descendant::*\"),2)" "t"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//b/descendant-or-self::*\")" "VALUE!"
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//b/descendant-or-self::*\"),2)" -7.0 0.0
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a/self::a\")" 10.0 0.0
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a/self::a\"),2)" 2.5 0.0
text m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[contains(.,\"\"b\"\")]\")" "abc"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[starts-with(.,\"\"1\"\")]\")" 10.0 0.0
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[starts-with(.,\"\"1\"\")]\"),2)" "REF!"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[string-length(.)=3]\")" 2.5 0.0
text m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[string-length(.)=3]\"),2)" "abc"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id mod 2=1]\")" 10.0 0.0
text m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id mod 2=1]\"),2)" "abc"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id div 2=1]\")" 2.5 0.0
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id div 2=1]\"),2)" "REF!"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id*2=4]\")" 2.5 0.0
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id*2=4]\"),2)" "REF!"
text m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id+1=4]\")" "abc"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id-1=3]\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id-1=3]\"),2)" "VALUE!"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id=1 or @id=5]\")" 10.0 0.0
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id=1 or @id=5]\"),2)" "VALUE!"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id>1 and @id<4]\")" 2.5 0.0
text m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id>1 and @id<4]\"),2)" "abc"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id!=1]\")" 2.5 0.0
text m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id!=1]\"),2)" "abc"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id<=2]\")" 10.0 0.0
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id<=2]\"),2)" 2.5 0.0
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id>=4]\")" -7.0 0.0
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id>=4]\"),2)" "VALUE!"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[-@id=-2]\")" 2.5 0.0
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[-@id=-2]\"),2)" "REF!"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[number(.)=2.5]\")" 2.5 0.0
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[number(.)=2.5]\"),2)" "REF!"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[name()=\"\"a\"\"][1]\")" 10.0 0.0
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[name()=\"\"a\"\"][1]\"),2)" -7.0 0.0
text m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//*[local-name()=\"\"c\"\"]\")" "t"
text m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[translate(.,\"\"abc\"\",\"\"ABC\"\")=\"\"ABC\"\"]\")" "abc"
text m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[substring(.,2,1)=\"\"b\"\"]\")" "abc"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[substring-before(@n,\"\"y\"\")=\"\"\"\"]\")" 10.0 0.0
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[substring-before(@n,\"\"y\"\")=\"\"\"\"]\"),2)" 2.5 0.0
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[concat(@id,@n)=\"\"1x\"\"]\")" 10.0 0.0
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[concat(@id,@n)=\"\"1x\"\"]\"),2)" "REF!"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[normalize-space(.)!=\"\"\"\"]\")" 10.0 0.0
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[normalize-space(.)!=\"\"\"\"]\"),2)" 2.5 0.0
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[boolean(@n)]\")" 10.0 0.0
text m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[boolean(@n)]\"),2)" "abc"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[true()]\")" 10.0 0.0
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[true()]\"),2)" 2.5 0.0
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[false()]\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[false()]\"),2)" "VALUE!"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[floor(.)=2]\")" 2.5 0.0
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[floor(.)=2]\"),2)" "REF!"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[ceiling(.)=3]\")" 2.5 0.0
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[ceiling(.)=3]\"),2)" "REF!"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[round(.)=3]\")" 2.5 0.0
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[round(.)=3]\"),2)" "REF!"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[count(../a)=4]\")" 10.0 0.0
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[count(../a)=4]\"),2)" 2.5 0.0
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//b[count(*)=2]\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//b[count(*)=2]\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[sum(../a)>0]\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[sum(../a)>0]\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//node()\")" "VALUE!"
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//node()\"),2)" 10.0 0.0
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[1]/@*\")" 1.0 0.0
text m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[1]/@*\"),2)" "x"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id=\"\"5\"\"]\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id=\"\"5\"\"]\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[.=\"\"\"\"]\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[.=\"\"\"\"]\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//x\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//x\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id=9]\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id=9]\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"count(//a)\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"count(//a)\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"sum(//a/@id)\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"sum(//a/@id)\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"string(//a)\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"string(//a)\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"1+1\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"1+1\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"true()\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"true()\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[1]=10\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[1]=10\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"concat(\"\"a\"\",\"\"b\"\")\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"concat(\"\"a\"\",\"\"b\"\")\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"a\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"a\"),2)" "VALUE!"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"r/a\")" 10.0 0.0
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"r/a\"),2)" 2.5 0.0
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"/\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"/\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\".\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\".\"),2)" "VALUE!"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id=1]/@id\")" 1.0 0.0
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id=1]/@id\"),2)" "REF!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//comment()\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//comment()\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//processing-instruction()\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//processing-instruction()\"),2)" "VALUE!"
text m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[position()=last()-1]\")" "abc"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id='1']/following::*[1]\")" 2.5 0.0
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[@id='1']/following::*[1]\"),2)" "REF!"
text m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//b/*[2]\")" "t"
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//b/a/@id|//a[1]\")" 10.0 0.0
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//b/a/@id|//a[1]\"),2)" -7.0 0.0
rel m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\" //a[1] \")" 10.0 0.0
rel m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\" //a[1] \"),2)" -7.0 0.0
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//A\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//A\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[id(1)]\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[id(1)]\"),2)" "VALUE!"
error m "=FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[lang(\"\"en\"\")]\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>\",\"//a[lang(\"\"en\"\")]\"),2)" "VALUE!"
text m "=FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"//title\")" "Dune"
text m "=INDEX(FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"//title\"),2)" "Faust"
text m "=FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"//book[price>10]/title\")" "Faust"
rel m "=FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"//book/@year\")" 2001.0 0.0
rel m "=INDEX(FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"//book/@year\"),2)" 1999.0 0.0
text m "=FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"//title[@lang=\"\"en\"\"]\")" "Dune"
text m "=INDEX(FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"//title[@lang=\"\"en\"\"]\"),2)" "Hyperion"
text m "=FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"//book[@cat=\"\"f\"\"][2]/title\")" "Hyperion"
rel m "=FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"//book[last()]/price\")" 7.5 0.0
error m "=INDEX(FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"//book[last()]/price\"),2)" "REF!"
text m "=FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"//*[price<8]/title\")" "Hyperion"
text m "=INDEX(FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"//*[price<8]/title\"),2)" "Wired"
text m "=FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"/lib/*/title\")" "Dune"
text m "=INDEX(FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"/lib/*/title\"),2)" "Faust"
rel m "=FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"//book[title=\"\"Faust\"\"]/@year\")" 1999.0 0.0
error m "=INDEX(FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"//book[title=\"\"Faust\"\"]/@year\"),2)" "REF!"
rel m "=FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"//price\")" 9.99 0.0
rel m "=INDEX(FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"//price\"),2)" 12.0 0.0
error m "=FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"//book[1]\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"//book[1]\"),2)" "VALUE!"
error m "=FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"//mag\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"//mag\"),2)" "VALUE!"
error m "=FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"sum(//price)\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"sum(//price)\"),2)" "VALUE!"
text m "=FILTERXML(\"<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book><book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book><book cat='f' year='2010'><title l\"&\"ang='en'>Hyperion</title><price>7.5</price></book><mag><title>Wired</title><price>5</price></mag></lib>\",\"//book[price=//book/price[.>10]]/title\")" "Faust"
rel m "=FILTERXML(\"<a>1<b>2</b>3</a>\",\"/a\")" 13.0 0.0
error m "=INDEX(FILTERXML(\"<a>1<b>2</b>3</a>\",\"/a\"),2)" "REF!"
error m "=FILTERXML(\"<a>1<b>2</b>3</a>\",\"/a/text()\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<a>1<b>2</b>3</a>\",\"/a/text()\"),2)" "VALUE!"
rel m "=FILTERXML(\"<a> 12 </a>\",\"/a\")" 12.0 0.0
error m "=INDEX(FILTERXML(\"<a> 12 </a>\",\"/a\"),2)" "REF!"
rel m "=FILTERXML(\"<a>\"&CHAR(10)&\" 12\"&CHAR(10)&\"</a>\",\"/a\")" 12.0 0.0
error m "=INDEX(FILTERXML(\"<a>\"&CHAR(10)&\" 12\"&CHAR(10)&\"</a>\",\"/a\"),2)" "REF!"
rel m "=FILTERXML(\"<a>1e3</a>\",\"/a\")" 1000.0 0.0
error m "=INDEX(FILTERXML(\"<a>1e3</a>\",\"/a\"),2)" "REF!"
rel m "=FILTERXML(\"<a>1E-2</a>\",\"/a\")" 0.01 0.0
error m "=INDEX(FILTERXML(\"<a>1E-2</a>\",\"/a\"),2)" "REF!"
is "bool" m "=FILTERXML(\"<a>TRUE</a>\",\"/a\")" (CellValue.Boolean true)
is "bool" m "=FILTERXML(\"<a>true</a>\",\"/a\")" (CellValue.Boolean true)
rel m "=FILTERXML(\"<a>$5</a>\",\"/a\")" 5.0 0.0
error m "=INDEX(FILTERXML(\"<a>$5</a>\",\"/a\"),2)" "REF!"
rel m "=FILTERXML(\"<a>50%</a>\",\"/a\")" 0.5 0.0
error m "=INDEX(FILTERXML(\"<a>50%</a>\",\"/a\"),2)" "REF!"
text m "=FILTERXML(\"<a>1,5</a>\",\"/a\")" "1,5"
rel m "=FILTERXML(\"<a>1,000</a>\",\"/a\")" 1000.0 0.0
error m "=INDEX(FILTERXML(\"<a>1,000</a>\",\"/a\"),2)" "REF!"
rel m "=FILTERXML(\"<a>2020-01-02</a>\",\"/a\")" 43832.0 0.0
error m "=INDEX(FILTERXML(\"<a>2020-01-02</a>\",\"/a\"),2)" "REF!"
rel m "=FILTERXML(\"<a>1/2/2020</a>\",\"/a\")" 43832.0 0.0
error m "=INDEX(FILTERXML(\"<a>1/2/2020</a>\",\"/a\"),2)" "REF!"
rel m "=FILTERXML(\"<a>12:30</a>\",\"/a\")" 0.5208333333333334 0.0
error m "=INDEX(FILTERXML(\"<a>12:30</a>\",\"/a\"),2)" "REF!"
rel m "=FILTERXML(\"<a>-0</a>\",\"/a\")" 0.0 0.0
error m "=INDEX(FILTERXML(\"<a>-0</a>\",\"/a\"),2)" "REF!"
rel m "=FILTERXML(\"<a>+5</a>\",\"/a\")" 5.0 0.0
error m "=INDEX(FILTERXML(\"<a>+5</a>\",\"/a\"),2)" "REF!"
rel m "=FILTERXML(\"<a>.5</a>\",\"/a\")" 0.5 0.0
error m "=INDEX(FILTERXML(\"<a>.5</a>\",\"/a\"),2)" "REF!"
rel m "=FILTERXML(\"<a>5.</a>\",\"/a\")" 5.0 0.0
error m "=INDEX(FILTERXML(\"<a>5.</a>\",\"/a\"),2)" "REF!"
text m "=FILTERXML(\"<a>0x10</a>\",\"/a\")" "0x10"
rel m "=FILTERXML(\"<a>(5)</a>\",\"/a\")" -5.0 0.0
error m "=INDEX(FILTERXML(\"<a>(5)</a>\",\"/a\"),2)" "REF!"
text m "=FILTERXML(\"<a>1 2</a>\",\"/a\")" "1 2"
rel m "=FILTERXML(\"<a>12345678901234567890</a>\",\"/a\")" 1.23456789012345e+19 0.0
error m "=INDEX(FILTERXML(\"<a>12345678901234567890</a>\",\"/a\"),2)" "REF!"
rel m "=FILTERXML(\"<a>0.1234567890123456789</a>\",\"/a\")" 0.123456789012345 0.0
error m "=INDEX(FILTERXML(\"<a>0.1234567890123456789</a>\",\"/a\"),2)" "REF!"
text m "=FILTERXML(\"<a>1e400</a>\",\"/a\")" "1e400"
text m "=FILTERXML(\"<a>#N/A</a>\",\"/a\")" "#N/A"
text m "=FILTERXML(\"<a>Infinity</a>\",\"/a\")" "Infinity"
text m "=FILTERXML(\"<a>NaN</a>\",\"/a\")" "NaN"
rel m "=FILTERXML(\"<a>00012</a>\",\"/a\")" 12.0 0.0
error m "=INDEX(FILTERXML(\"<a>00012</a>\",\"/a\"),2)" "REF!"
error m "=FILTERXML(\"<a></a>\",\"/a\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<a></a>\",\"/a\"),2)" "VALUE!"
error m "=FILTERXML(\"<a/>\",\"/a\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<a/>\",\"/a\"),2)" "VALUE!"
error m "=FILTERXML(\"<a><b/></a>\",\"/a\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<a><b/></a>\",\"/a\"),2)" "VALUE!"
text m "=FILTERXML(\"<a>&amp;&lt;&gt;&quot;&apos;</a>\",\"/a\")" "&<>\"'"
text m "=FILTERXML(\"<a>&#65;&#x42;</a>\",\"/a\")" "AB"
text m "=FILTERXML(\"<a><![CDATA[<x>]]></a>\",\"/a\")" "<x>"
text m "=FILTERXML(\"<a><!-- c -->x</a>\",\"/a\")" "x"
error m "=FILTERXML(\"<a><!-- c -->x</a>\",\"//comment()\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<a><!-- c -->x</a>\",\"//comment()\"),2)" "VALUE!"
rel m "=FILTERXML(\"<?xml version='1.0'?><a>1</a>\",\"/a\")" 1.0 0.0
error m "=INDEX(FILTERXML(\"<?xml version='1.0'?><a>1</a>\",\"/a\"),2)" "REF!"
rel m "=FILTERXML(\"<?xml version='1.0' encoding='UTF-8'?><a>1</a>\",\"/a\")" 1.0 0.0
error m "=INDEX(FILTERXML(\"<?xml version='1.0' encoding='UTF-8'?><a>1</a>\",\"/a\"),2)" "REF!"
error m "=FILTERXML(\"<a>1</a><b>2</b>\",\"/a\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<a>1</a><b>2</b>\",\"/a\"),2)" "VALUE!"
error m "=FILTERXML(\"<a>1\",\"/a\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<a>1\",\"/a\"),2)" "VALUE!"
error m "=FILTERXML(\"a\",\"/a\")" "VALUE!"
error m "=INDEX(FILTERXML(\"a\",\"/a\"),2)" "VALUE!"
error m "=FILTERXML(\"\",\"/a\")" "VALUE!"
error m "=INDEX(FILTERXML(\"\",\"/a\"),2)" "VALUE!"
error m "=FILTERXML(\"<a><b></a></b>\",\"/a\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<a><b></a></b>\",\"/a\"),2)" "VALUE!"
error m "=FILTERXML(\"<a x='1' x='2'/>\",\"/a\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<a x='1' x='2'/>\",\"/a\"),2)" "VALUE!"
rel m "=FILTERXML(\"<a xmlns='urn:x'><b>1</b></a>\",\"//b\")" 1.0 0.0
error m "=INDEX(FILTERXML(\"<a xmlns='urn:x'><b>1</b></a>\",\"//b\"),2)" "REF!"
rel m "=FILTERXML(\"<a xmlns='urn:x'><b>1</b></a>\",\"//*[local-name()='b']\")" 1.0 0.0
error m "=INDEX(FILTERXML(\"<a xmlns='urn:x'><b>1</b></a>\",\"//*[local-name()='b']\"),2)" "REF!"
rel m "=FILTERXML(\"<a xmlns:p='urn:p'><p:b>1</p:b></a>\",\"//p:b\")" 1.0 0.0
error m "=INDEX(FILTERXML(\"<a xmlns:p='urn:p'><p:b>1</p:b></a>\",\"//p:b\"),2)" "REF!"
rel m "=FILTERXML(\"<a xmlns:p='urn:p'><p:b>1</p:b></a>\",\"//*[local-name()='b']\")" 1.0 0.0
error m "=INDEX(FILTERXML(\"<a xmlns:p='urn:p'><p:b>1</p:b></a>\",\"//*[local-name()='b']\"),2)" "REF!"
rel m "=FILTERXML(\"<a xmlns:p='urn:p'><p:b>1</p:b></a>\",\"//*[name()='p:b']\")" 1.0 0.0
error m "=INDEX(FILTERXML(\"<a xmlns:p='urn:p'><p:b>1</p:b></a>\",\"//*[name()='p:b']\"),2)" "REF!"
rel m "=FILTERXML(\"<p:a xmlns:p='urn:p'>1</p:a>\",\"/*\")" 1.0 0.0
error m "=INDEX(FILTERXML(\"<p:a xmlns:p='urn:p'>1</p:a>\",\"/*\"),2)" "REF!"
text m "=FILTERXML(\"<a><b>x</b><b>y</b><b>z</b></a>\",\"//b\")" "x"
text m "=INDEX(FILTERXML(\"<a><b>x</b><b>y</b><b>z</b></a>\",\"//b\"),2)" "y"
rel m "=FILTERXML(\"<a><b>1</b><b>x</b><b>3</b></a>\",\"//b\")" 1.0 0.0
text m "=INDEX(FILTERXML(\"<a><b>1</b><b>x</b><b>3</b></a>\",\"//b\"),2)" "x"
error m "=FILTERXML(\"<!DOCTYPE a [<!ENTITY e 'hi'>]><a>&e;</a>\",\"/a\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<!DOCTYPE a [<!ENTITY e 'hi'>]><a>&e;</a>\",\"/a\"),2)" "VALUE!"
error m "=FILTERXML(\"<!DOCTYPE a><a>1</a>\",\"/a\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<!DOCTYPE a><a>1</a>\",\"/a\"),2)" "VALUE!"
text m "=FILTERXML(\"<a>café 中</a>\",\"/a\")" "café 中"
text m "=FILTERXML(\"<a b='x &amp; y'/>\",\"/a/@b\")" "x & y"
rel m "=FILTERXML(\"<a b='1'/>\",\"/a/@b\")" 1.0 0.0
error m "=INDEX(FILTERXML(\"<a b='1'/>\",\"/a/@b\"),2)" "REF!"
rel m "=FILTERXML(\"<a b=' 2 '/>\",\"/a/@b\")" 2.0 0.0
error m "=INDEX(FILTERXML(\"<a b=' 2 '/>\",\"/a/@b\"),2)" "REF!"
error m "=FILTERXML(\"<a>\"&CHAR(9)&\"<b>1</b>\"&CHAR(10)&\"</a>\",\"/a\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<a>\"&CHAR(9)&\"<b>1</b>\"&CHAR(10)&\"</a>\",\"/a\"),2)" "VALUE!"
rel m "=FILTERXML(\"<a>\"&CHAR(9)&\"<b>1</b>\"&CHAR(10)&\"</a>\",\"/a/node()\")" 1.0 0.0
error m "=INDEX(FILTERXML(\"<a>\"&CHAR(9)&\"<b>1</b>\"&CHAR(10)&\"</a>\",\"/a/node()\"),2)" "REF!"
text m "=FILTERXML(\"<a>  x  y  </a>\",\"/a\")" "x  y"
error m "=FILTERXML(\"<a>x</a>\",\"/a/@nope\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<a>x</a>\",\"/a/@nope\"),2)" "VALUE!"
error m "=FILTERXML(\"<a><b>1</b></a>\",\"name(/a/b)\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<a><b>1</b></a>\",\"name(/a/b)\"),2)" "VALUE!"
error m "=FILTERXML(\"<a><b>1</b></a>\",\"/a/b[1]/text()[1]\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<a><b>1</b></a>\",\"/a/b[1]/text()[1]\"),2)" "VALUE!"
rel m "=FILTERXML(\"<root><i>1</i><i>2</i><i>3</i><i>4</i><i>5</i><i>6</i><i>7</i></root>\",\"//i\")" 1.0 0.0
rel m "=INDEX(FILTERXML(\"<root><i>1</i><i>2</i><i>3</i><i>4</i><i>5</i><i>6</i><i>7</i></root>\",\"//i\"),2)" 2.0 0.0
rel m "=FILTERXML(\"<root><i>1</i><i>2</i><i>3</i><i>4</i><i>5</i><i>6</i><i>7</i></root>\",\"//i[position() mod 2 = 0]\")" 2.0 0.0
rel m "=INDEX(FILTERXML(\"<root><i>1</i><i>2</i><i>3</i><i>4</i><i>5</i><i>6</i><i>7</i></root>\",\"//i[position() mod 2 = 0]\"),2)" 4.0 0.0
rel m "=FILTERXML(\"<root><i>1</i><i>2</i></root>\",\"//i[. = 2] | //i[. = 1]\")" 1.0 0.0
rel m "=INDEX(FILTERXML(\"<root><i>1</i><i>2</i></root>\",\"//i[. = 2] | //i[. = 1]\"),2)" 2.0 0.0
error m "=FILTERXML(\"<r><a>x</a></r>\",\"//a/following-sibling::a\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<r><a>x</a></r>\",\"//a/following-sibling::a\"),2)" "VALUE!"
rel m "=FILTERXML(\"<a> <b>1</b> </a>\",\"/a/node()\")" 1.0 0.0
error m "=INDEX(FILTERXML(\"<a> <b>1</b> </a>\",\"/a/node()\"),2)" "REF!"
text m "=FILTERXML(\"<a> </a>\",\"/a\")" " "
error m "=FILTERXML(\"<a>1</a>\",\"/a/text()/text()\")" "VALUE!"
error m "=INDEX(FILTERXML(\"<a>1</a>\",\"/a/text()/text()\"),2)" "VALUE!"
error m "=FILTERXML(1,\"/a\")" "VALUE!"
error m "=FILTERXML(\"<a>1</a>\",1)" "VALUE!"
error m "=FILTERXML(TRUE,\"/a\")" "VALUE!"
error m "=FILTERXML(\"<a>1</a>\",1/0)" "DIV/0!"
error m "=FILTERXML(1/0,\"/a\")" "DIV/0!"
error m "=FILTERXML(\"<a>1</a>\",\"\")" "VALUE!"
is "bool" m "=FILTERXML(\"<a>TRUE</a>\",\"/a\")=TRUE" (CellValue.Boolean true)
rel m "=FILTERXML(\"<a>5</a>\",\"/a\")+1" 6.0 0.0
rel m "=FILTERXML({\"<a>1</a>\",\"<a>2</a>\"},\"/a\")" 1.0 0.0
error m "=INDEX(FILTERXML({\"<a>1</a>\",\"<a>2</a>\"},\"/a\"),1,2)" "REF!"

printfn "stage69 normal distribution, ETS, FILTERXML, EUROCONVERT, BAHTTEXT: %d passed, %d failed" pass fail
if fail > 0 then exit 1
