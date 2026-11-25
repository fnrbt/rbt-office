/// Coupon bonds of the function table: the COUP* family, PRICE, YIELD,
/// DURATION, MDURATION, ACCRINT, the odd first and last period prices and
/// yields, and the French depreciations AMORDEGRC and AMORLINC. One coupon
/// schedule and one day-count core serve them all (`Schedule`, `Basis`).
/// The rules where Excel is not the textbook are noted where they are, as
/// found against Excel 2019 (tools/excel-parity/make-bonds-probe.py).
module Rbt.Spreadsheet.Functions.FnBonds

open System
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.Functions.Kit

let moduleName = "Bonds"

// --- Calendar ----------------------------------------------------------------
// A date is its day number in the 1900 system (serial; the workbook's own
// system is shifted in and out at the arguments), its calendar fields worked
// out in whole numbers (proleptic Gregorian, so 1900 is no leap year: the
// bond functions of Excel do not see its phantom 29 February either).

[<Struct>]
type Date = { Serial: int; Year: int; Month: int; Day: int }

/// Days from 1970-01-01 of a calendar date (H. Hinnant's days_from_civil).
let private daysFromCivil (y: int) (m: int) (d: int) =
    let y = if m <= 2 then y - 1 else y
    let era = (if y >= 0 then y else y - 399) / 400
    let yoe = y - era * 400
    let doy = (153 * (if m > 2 then m - 3 else m + 9) + 2) / 5 + d - 1
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy
    era * 146097 + doe - 719468

let private epoch = daysFromCivil 1899 12 30

let isLeap (y: int) = (y % 4 = 0 && y % 100 <> 0) || y % 400 = 0
let daysInMonth (y: int) (m: int) =
    match m with
    | 2 -> if isLeap y then 29 else 28
    | 4 | 6 | 9 | 11 -> 30
    | _ -> 31

let dateOf (y: int) (m: int) (d: int) : Date = { Serial = daysFromCivil y m d - epoch; Year = y; Month = m; Day = d }

let dateOfSerial (serial: int) : Date =
    let z = serial + epoch + 719468
    let era = (if z >= 0 then z else z - 146096) / 146097
    let doe = z - era * 146097
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100)
    let mp = (5 * doy + 2) / 153
    let d = doy - (153 * mp + 2) / 5 + 1
    let m = if mp < 10 then mp + 3 else mp - 9
    let y = yoe + era * 400 + (if m <= 2 then 1 else 0)
    { Serial = serial; Year = y; Month = m; Day = d }

let isMonthEnd (d: Date) = d.Day = daysInMonth d.Year d.Month

/// The date `months` on (back when negative), its day held to the month's
/// length; the month's last day when `toMonthEnd`.
let addMonths (d: Date) (months: int) (toMonthEnd: bool) : Date =
    let total = d.Year * 12 + (d.Month - 1) + months
    let y = (if total >= 0 then total else total - 11) / 12
    let m = total - y * 12 + 1
    let last = daysInMonth y m
    dateOf y m (if toMonthEnd then last else min d.Day last)

// --- Day counts ----------------------------------------------------------------

/// 30/360 days from a to b by the US (NASD) rule of basis 0. A start on
/// February's last day counts as the 30th; an end on it does too when the
/// start is one as well, or always when `modifyEnd` (the coupon period's own
/// length). An end on a 31st is the 30th when `modifyEnd` or when the start's
/// own day is 30 or 31 - before February's end is moved: Excel counts 28 Feb
/// to 31 Aug as 181 days.
let days360Us (modifyEnd: bool) (a: Date) (b: Date) =
    let aFebEnd = a.Month = 2 && isMonthEnd a
    let bFebEnd = b.Month = 2 && isMonthEnd b
    let mutable d1 = a.Day
    let mutable d2 = b.Day
    if bFebEnd && (aFebEnd || modifyEnd) then d2 <- 30
    if d2 = 31 && (d1 >= 30 || modifyEnd) then d2 <- 30
    if d1 = 31 || aFebEnd then d1 <- 30
    float ((b.Year - a.Year) * 360 + (b.Month - a.Month) * 30 + (d2 - d1))

/// 30/360 days from a to b by the European rule of basis 4: a 31st is the 30th.
let days360Eu (a: Date) (b: Date) =
    let d1 = min a.Day 30
    let d2 = min b.Day 30
    float ((b.Year - a.Year) * 360 + (b.Month - a.Month) * 30 + (d2 - d1))

/// The days from a to b in a basis: 30/360 for 0 and 4, actual days otherwise.
let daysIn (basis: int) (a: Date) (b: Date) =
    match basis with
    | 0 -> days360Us false a b
    | 4 -> days360Eu a b
    | _ -> float (b.Serial - a.Serial)

/// The days of the year in a basis, for the year fraction from a to b
/// (YEARFRAC's denominator: basis 1 averages the years a span of more than a
/// year touches).
let yearDays (basis: int) (a: Date) (b: Date) =
    match basis with
    | 1 ->
        let withinYear =
            a.Year = b.Year || (b.Year = a.Year + 1 && (b.Month < a.Month || (b.Month = a.Month && b.Day <= a.Day)))
        if withinYear then
            if a.Year = b.Year then (if isLeap a.Year then 366.0 else 365.0)
            elif (isLeap a.Year && a.Month < 3) || (isLeap b.Year && (b.Month > 2 || (b.Month = 2 && b.Day = 29))) then 366.0
            else 365.0
        else
            let mutable total = 0.0
            for y in a.Year .. b.Year do total <- total + (if isLeap y then 366.0 else 365.0)
            total / float (b.Year - a.Year + 1)
    | 3 -> 365.0
    | _ -> 360.0

/// YEARFRAC of a to b (a <= b).
let yearFrac (basis: int) (a: Date) (b: Date) =
    match basis with
    | 1 -> float (b.Serial - a.Serial) / yearDays 1 a b
    | _ -> daysIn basis a b / yearDays basis a b

// --- The coupon schedule ----------------------------------------------------------

/// The coupon dates of a bond: back from maturity by whole periods, each
/// counted from maturity itself (so a 30th stays a 30th after a February),
/// on the month's last day throughout when maturity is a month's last day.
type Schedule = { Previous: Date; Next: Date; Remaining: int }

/// The coupon `k` periods before `anchor` (k < 0: after).
let couponDate (anchor: Date) (months: int) (k: int) = addMonths anchor (-k * months) (isMonthEnd anchor)

/// The coupon dates around `settlement` (the last on or before it, the next
/// after it) of the schedule ending at `maturity`, and how many coupons are
/// still to be paid.
let schedule (settlement: Date) (maturity: Date) (frequency: int) : Schedule =
    let months = 12 / frequency
    let span = (maturity.Year - settlement.Year) * 12 + maturity.Month - settlement.Month
    let mutable k = max 1 (span / months)
    while (couponDate maturity months k).Serial > settlement.Serial do k <- k + 1
    while k > 1 && (couponDate maturity months (k - 1)).Serial <= settlement.Serial do k <- k - 1
    { Previous = couponDate maturity months k; Next = couponDate maturity months (k - 1); Remaining = k }

/// The days of the coupon period from `previous` to `next` in a basis.
let periodDays (basis: int) (frequency: int) (previous: Date) (next: Date) =
    match basis with
    | 1 -> float (next.Serial - previous.Serial)
    | 3 -> 365.0 / float frequency
    | _ -> 360.0 / float frequency

/// The actual days of the coupon period holding settlement as COUPDAYS of
/// basis 1 counts them, which is not always from COUPPCD to COUPNCD: the next
/// coupon is found by stepping back from maturity a period at a time, each
/// step from the last (so the day drifts: a 30 August maturity has 28
/// February, then 28 August coupons), and the period's start is a period
/// before that, on the month's last day when that coupon is on one (30
/// September back to 31 March); while that start is after settlement, the
/// period before it is taken. (Found on 1608 settlements of six maturities:
/// make-bonds-probe.py's sweep.)
let actualCoupDays (settlement: Date) (maturity: Date) (frequency: int) =
    let months = 12 / frequency
    let mutable next = maturity
    let mutable step = maturity
    while step.Serial > settlement.Serial do
        next <- step
        step <- addMonths step (-months) false
    let back (d: Date) = addMonths d (-months) (isMonthEnd d)
    let mutable start = back next
    while start.Serial > settlement.Serial do
        next <- start
        start <- back next
    float (next.Serial - start.Serial)

/// The days of a coupon period, COUPDAYS: 360 or 365 a year over the
/// frequency, basis 1 the actual days (`actualCoupDays`).
let coupDays (basis: int) (frequency: int) (settlement: Date) (maturity: Date) =
    match basis with
    | 1 -> actualCoupDays settlement maturity frequency
    | 3 -> 365.0 / float frequency
    | _ -> 360.0 / float frequency
/// The days from the period's start to settlement, COUPDAYBS.
let coupDaysBS (basis: int) (settlement: Date) (s: Schedule) = daysIn basis s.Previous settlement
/// The days from settlement to the next coupon, COUPDAYSNC. Basis 0 counts
/// them as the period's 30/360 days (an end on February's last day the 30th)
/// less those before settlement.
let coupDaysNC (basis: int) (settlement: Date) (s: Schedule) =
    match basis with
    | 0 -> days360Us true s.Previous s.Next - days360Us false s.Previous settlement
    | _ -> daysIn basis settlement s.Next

// --- Arguments ---------------------------------------------------------------------

/// The last serial Excel's calendar has (9999-12-31).
let private lastSerial = 2958465.0

/// A date argument: a serial within the calendar, its time of day dropped.
let private dateAt (a: Args) (i: int) : Result<Date, string> =
    res {
        let! x = a.AtpNumber i
        let x = Math.Floor x + dateSystemDays a.Ctx.Model.Date1904
        do! check (x >= 0.0 && x <= lastSerial) Err.num
        return dateOfSerial (int x) }

let private serialOut (a: Args) (d: Date) = CellValue.Number (float d.Serial - dateSystemDays a.Ctx.Model.Date1904)

/// The frequency argument: 1, 2 or 4 (truncated), else #NUM!.
let private frequencyAt (a: Args) (i: int) : Result<int, string> =
    res {
        let! f = a.AtpNumber i
        let f = truncate f
        do! check (f = 1.0 || f = 2.0 || f = 4.0) Err.num
        return int f }

/// The basis argument (0 when omitted): 0 to 4 truncated, else #NUM!.
let private basisAt (a: Args) (i: int) : Result<int, string> =
    res {
        let! b = a.AtpNumberOr i 0.0
        let b = truncate b
        do! check (b >= 0.0 && b <= 4.0) Err.num
        return int b }

/// Settlement and maturity (arguments 0 and 1), settlement before maturity.
let private term (a: Args) =
    res {
        let! settlement = dateAt a 0
        let! maturity = dateAt a 1
        return settlement, maturity }

// --- Prices and yields ---------------------------------------------------------------

/// The coupon period around settlement as PRICE, YIELD and DURATION see it:
/// its days E (basis 1 the actual days from COUPPCD to COUPNCD, not
/// COUPDAYS's), the days A from its start to settlement, the days DSC to the
/// next coupon taken as E - A in every basis (not COUPDAYSNC: an actual/360
/// period of 181 days counts 180), and the coupons N still to come.
let private period (settlement: Date) (maturity: Date) (frequency: int) (basis: int) =
    let s = schedule settlement maturity frequency
    let e = periodDays basis frequency s.Previous s.Next
    let a = coupDaysBS basis settlement s
    e, a, e - a, s.Remaining

/// PRICE: the price per 100 of face value at a yield.
let price (settlement: Date) (maturity: Date) (rate: float) (yld: float) (redemption: float) (frequency: int) (basis: int) =
    let e, a, dsc, n = period settlement maturity frequency basis
    let f = float frequency
    let coupon = 100.0 * rate / f
    let accrued = 100.0 * rate / f * a / e
    if n = 1 then (redemption + coupon) / (1.0 + dsc / e * yld / f) - accrued
    else
        let factor k = Math.Pow (1.0 + yld / f, k - 1.0 + dsc / e)
        let mutable coupons = 0.0
        for k in 1 .. n do coupons <- coupons + coupon / factor (float k)
        redemption / factor (float n) + coupons - accrued

/// The root of `f` near `guess` by Newton's method (a numerical slope), as
/// Excel's iterative yields: None when it does not converge.
let solve (f: float -> float) (guess: float) : float option =
    let mutable x = guess
    let mutable result = None
    let mutable i = 0
    while result.IsNone && i < 100 do
        let fx = f x
        let h = max 1e-8 (abs x * 1e-8)
        let slope = (f (x + h) - fx) / h
        if slope = 0.0 || Double.IsNaN slope then i <- 100
        else
            let next = x - fx / slope
            if abs (next - x) < 1e-12 then result <- Some next
            x <- next
            i <- i + 1
    result

let private yieldOf (settlement: Date) (maturity: Date) (rate: float) (pr: float) (redemption: float) (frequency: int) (basis: int) =
    let s = schedule settlement maturity frequency
    let f = float frequency
    if s.Remaining = 1 then
        // The last period in closed form, its days as Excel counts them here:
        // the period 360/f days for bases 0 and 4 but its actual days for 1, 2
        // and 3; to redemption the basis's days (none: a yield of 0).
        let e = if basis = 0 || basis = 4 then 360.0 / f else float (s.Next.Serial - s.Previous.Serial)
        let a = coupDaysBS basis settlement s
        let dsr = daysIn basis settlement maturity
        let paid = pr / 100.0 + a / e * rate / f
        if dsr = 0.0 then Some 0.0
        else Some ((redemption / 100.0 + rate / f - paid) / paid * f * e / dsr)
    else solve (fun y -> price settlement maturity rate y redemption frequency basis - pr) 0.05

/// Macaulay duration in years, the sums in Excel's order (the time of the
/// k-th payment (k - 1) + DSC/E, of the last DSC/E + N - 1).
let duration (settlement: Date) (maturity: Date) (coupon: float) (yld: float) (frequency: int) (basis: int) =
    let e, _, dsc, n = period settlement maturity frequency basis
    let f = float frequency
    let x1 = dsc / e
    let x2 = x1 + float n - 1.0
    let x3 = yld / f + 1.0
    let x4 = Math.Pow (x3, x2)
    let mutable weighted = x2 * 100.0 / x4
    let mutable pv = 100.0 / x4
    let mutable term2 = 0.0
    let mutable term4 = 0.0
    for k in 1 .. n do
        let x5 = float k - 1.0 + x1
        let x7 = 100.0 * coupon / f / Math.Pow (x3, x5)
        term2 <- term2 + x7 * x5
        term4 <- term4 + x7
    weighted <- weighted + term2
    pv <- pv + term4
    weighted / pv / f

// --- Accrued interest -------------------------------------------------------------------

/// ACCRINT: the interest accrued from `issue` to `settlement`, counted in
/// the quasi-coupon periods of the schedule through `firstInterest` (q_k,
/// k periods before it). Excel's count, as found on 770 probes
/// (make-bonds-probe.py's ACCRINT section), is not the textbook sum of
/// A_i/NL_i:
/// - the stub: the days from issue to the first quasi-coupon date on or
///   after it, over that period's length NL;
/// - each whole period after it counts 1;
/// - settlement on or before first interest: the last period counts
///   (EA - DSC)/E, DSC the days from settlement to the next quasi-coupon
///   date and EA, E the days and the length of the period ending at first
///   interest (EA actual days for bases 1 and 2, 365/f for 3; DSC in 30/360
///   ends a 31st on the 30th). With settlement in the stub's own period the
///   total is the stub less DSC/E (DSC as the basis counts it).
/// - settlement after first interest: the days from the quasi-coupon date
///   before it, over E;
/// - calc_method FALSE: the stub and the days from the period start before
///   first interest to settlement over E (negative when settlement is
///   before that start, as Excel's is).
let accrint (issue: Date) (firstInterest: Date) (settlement: Date) (rate: float) (par: float) (frequency: int) (basis: int) (fromIssue: bool) =
    let months = 12 / frequency
    let q k = couponDate firstInterest months k
    let lengthOf (p: Date) (n: Date) = periodDays basis frequency p n
    // k: the first quasi-coupon date on or after issue is q_k.
    let mutable k = 0
    while (q (k + 1)).Serial >= issue.Serial do k <- k + 1
    while (q k).Serial < issue.Serial do k <- k - 1
    let q1 = q k
    let p0 = q (k + 1)
    let stub = daysIn basis issue q1 / lengthOf p0 q1
    let pf = q 1
    let e = lengthOf pf firstInterest
    let total =
        if not fromIssue then
            if q1.Serial <= pf.Serial then stub + daysIn basis pf settlement / e
            else daysIn basis issue settlement / lengthOf p0 q1
        elif settlement.Serial > firstInterest.Serial then
            let mutable j = 0
            while (q (j - 1)).Serial <= settlement.Serial do j <- j - 1
            stub + float (k - j) + daysIn basis (q j) settlement / e
        else
            let mutable j = 0
            while (q (j + 1)).Serial >= settlement.Serial do j <- j + 1
            let next = q j
            let c = k - j
            if c = 0 then stub - daysIn basis settlement next / e
            else
                let dsc = if basis = 0 then days360Us true settlement next else daysIn basis settlement next
                let ea =
                    match basis with
                    | 1 | 2 -> float (firstInterest.Serial - pf.Serial)
                    | 3 -> 365.0 / float frequency
                    | _ -> daysIn basis pf firstInterest
                stub + float (c - 1) + (ea - dsc) / e
    par * rate / float frequency * total

// --- Odd periods ---------------------------------------------------------------------------

/// The coupons after `firstCoupon` up to maturity, -1 when the first coupon
/// is not one of maturity's schedule.
let private couponsAfter (firstCoupon: Date) (maturity: Date) (frequency: int) =
    let months = 12 / frequency
    let mutable k = 0
    while (couponDate maturity months k).Serial > firstCoupon.Serial do k <- k + 1
    if (couponDate maturity months k).Serial = firstCoupon.Serial then k else -1

/// ODDFPRICE: a bond whose first period is short or long. The textbook
/// formulas with Excel's counts (found on 1020 probes, make-bonds-probe.py's
/// ODDFPRICE section):
/// - E is the actual days (basis 1) of the quasi-coupon period holding
///   settlement, its basis's 360/f or 365/f otherwise; a first period of more
///   than E days (issue to first coupon, in the basis) is long.
/// - Long: the quasi-coupon periods are stepped back from the first coupon a
///   period at a time, each from the last (the day drifts, no month ends
///   kept); every period but the one holding issue counts its full NL days;
///   DSC is E less the days from the period's start to settlement, the actual
///   days to the next quasi-coupon date for bases 2 and 3.
/// - Nq, the whole periods from settlement to the first coupon, are counted
///   by stepping forward from settlement - on the month's last days when the
///   first coupon is on one, the step from settlement first going to its
///   month's end and counting one: a first coupon on a month's end thus often
///   discounts every payment one period further than the calendar says.
let oddFirstPrice (settlement: Date) (maturity: Date) (issue: Date) (firstCoupon: Date) (rate: float) (yld: float) (redemption: float) (frequency: int) (basis: int) =
    let f = float frequency
    let months = 12 / frequency
    let quasi = schedule settlement firstCoupon frequency
    let e = periodDays basis frequency quasi.Previous quasi.Next
    let n = couponsAfter firstCoupon maturity frequency
    let dfc = daysIn basis issue firstCoupon
    let factor = 1.0 + yld / f
    let coupon = 100.0 * rate / f
    if dfc <= e then
        let dsc = daysIn basis settlement firstCoupon
        let a = daysIn basis issue settlement
        let mutable total = redemption / Math.Pow (factor, float n + dsc / e)
        total <- total + coupon * dfc / e / Math.Pow (factor, dsc / e)
        for k in 1 .. n do total <- total + coupon / Math.Pow (factor, float k + dsc / e)
        total - coupon * a / e
    else
        let mutable nc = 1
        while (couponDate firstCoupon months nc).Serial > issue.Serial do nc <- nc + 1
        let mutable dcnl = 0.0
        let mutable anl = 0.0
        let mutable late = firstCoupon
        for i in 1 .. nc do
            let early = addMonths late (-months) false
            let nl = periodDays basis frequency early late
            let start = if issue.Serial > early.Serial then issue else early
            let dc = if i < nc then nl else daysIn basis start late
            let finish = if settlement.Serial < late.Serial then settlement else late
            let a = if finish.Serial > start.Serial then daysIn basis start finish else 0.0
            dcnl <- dcnl + dc / nl
            anl <- anl + a / nl
            late <- early
        let dsc =
            if basis = 2 || basis = 3 then float (quasi.Next.Serial - settlement.Serial)
            else e - coupDaysBS basis settlement quasi
        let nq =
            let monthEnd =
                if not (isMonthEnd firstCoupon) && firstCoupon.Month <> 2 && firstCoupon.Day > 28 then isMonthEnd settlement
                else isMonthEnd firstCoupon
            let start = addMonths settlement 0 monthEnd
            let mutable count = if settlement.Serial < start.Serial then 1 else 0
            let mutable d = addMonths start months monthEnd
            while d.Serial < firstCoupon.Serial do
                count <- count + 1
                d <- addMonths d months monthEnd
            count
        let mutable total = redemption / Math.Pow (factor, float (n + nq) + dsc / e)
        total <- total + coupon * dcnl / Math.Pow (factor, float nq + dsc / e)
        for k in 1 .. n do total <- total + coupon / Math.Pow (factor, float (k + nq) + dsc / e)
        total - coupon * anl

/// The sums of ODDLPRICE and ODDLYIELD over the quasi-coupon periods from the
/// last coupon on: the days counted (DC), accrued (A) and to come (DSC), each
/// over its period's own days NL. As Excel counts them (found on 1134 probes,
/// make-bonds-probe.py's ODDLPRICE section): the periods are as many as
/// COUPNUM counts from the last coupon to maturity, stepped forward a period
/// at a time from the last coupon (the day drifts); NL is the period's days
/// in the basis - actual days for bases 2 and 3 too, not 360/f or 365/f, and
/// in 30/360 an end on February's last day or on a 31st the 30th (the
/// days counted too); every period but the last counts its NL days.
let private oddLastSums (settlement: Date) (maturity: Date) (lastInterest: Date) (frequency: int) (basis: int) =
    let months = 12 / frequency
    let nc = (schedule lastInterest maturity frequency).Remaining
    let mutable dcnl = 0.0
    let mutable anl = 0.0
    let mutable dscnl = 0.0
    let mutable early = lastInterest
    for i in 1 .. nc do
        let late = addMonths early months false
        let days (a: Date) (b: Date) = if basis = 0 then days360Us true a b else daysIn basis a b
        let nl = days early late
        let dc = if i < nc then nl else days early maturity
        let a =
            if late.Serial < settlement.Serial then dc
            elif early.Serial < settlement.Serial then daysIn basis early settlement
            else 0.0
        let start = if settlement.Serial > early.Serial then settlement else early
        let finish = if maturity.Serial < late.Serial then maturity else late
        let dsc = if finish.Serial > start.Serial then daysIn basis start finish else 0.0
        dcnl <- dcnl + dc / nl
        anl <- anl + a / nl
        dscnl <- dscnl + dsc / nl
        early <- late
    dcnl, anl, dscnl

let oddLastPrice settlement maturity lastInterest rate yld redemption frequency basis =
    let dcnl, anl, dscnl = oddLastSums settlement maturity lastInterest frequency basis
    let f = float frequency
    let x = 100.0 * rate / f
    (dcnl * x + redemption) / (dscnl * yld / f + 1.0) - anl * x

let oddLastYield settlement maturity lastInterest rate pr redemption frequency basis =
    let dcnl, anl, dscnl = oddLastSums settlement maturity lastInterest frequency basis
    let f = float frequency
    let x = 100.0 * rate / f
    let paid = anl * x + pr
    (dcnl * x + redemption - paid) / paid * f / dscnl

// --- French depreciation -----------------------------------------------------------------
// As Excel computes them (found on 1125 probes each, make-bonds-probe.py's
// AMOR section). The first period's fraction of a year is YEARFRAC's except
// in basis 1, where it is the actual days over the purchase year's length,
// and a purchase on the first period's end counts a whole year.

/// The year fraction of the first period.
let private firstFraction (bought: Date) (first: Date) (basis: int) =
    if bought.Serial = first.Serial then 1.0
    elif basis = 1 then float (first.Serial - bought.Serial) / (if isLeap bought.Year then 366.0 else 365.0)
    else yearFrac basis bought first

/// Excel's ROUND to a whole number: to 15 significant places first, then
/// halves away from zero (1637.4999999999998 is 1638).
let private roundCharge (x: float) =
    let x =
        if x = 0.0 || Double.IsNaN x || Double.IsInfinity x then x
        else
            let scale = Math.Pow (10.0, 14.0 - Math.Floor (Math.Log10 (abs x)))
            Math.Round (x * scale) / scale
    roundHalfAway x

let private amorlinc (cost: float) (bought: Date) (first: Date) (salvage: float) (period: int) (rate: float) (basis: int) =
    let one = cost * rate
    let first0 = min (firstFraction bought first basis * rate * cost) (cost - salvage)
    if period = 0 then first0
    else
        // What is left to depreciate: a whole rate's worth a period, the rest in the last.
        let mutable rest = cost - salvage - first0
        let mutable result = 0.0
        for k in 1 .. period do
            let charge = if rest >= one then one elif rest > 0.0 then rest else 0.0
            if k = period then result <- charge
            rest <- rest - one
        result

/// AMORDEGRC: the rate times a coefficient of the life 1/rate (1.5 up to 4
/// years, 2 up to 6, 2.5 beyond; up to 2 years #NUM!), each charge rounded to
/// a whole number. The book value runs on unrounded after the first period
/// (value less rate times value), so a charge can round differently from the
/// rate times the rounded values. The last two of the ceil(life) periods (one
/// fewer when bought on the first period's end) each take half the value left
/// before them, later ones nothing - unless the first half would fall on the
/// first period, when no halving happens. A value below salvage takes nothing.
let private amordegrc (cost: float) (bought: Date) (first: Date) (salvage: float) (period: int) (rate: float) (basis: int) =
    let life = 1.0 / rate
    let coefficient = if life <= 4.0 then 1.5 elif life <= 6.0 then 2.0 else 2.5
    let r = rate * coefficient
    let charge0 = roundCharge (firstFraction bought first basis * r * cost)
    let years = Math.Ceiling life
    let periods = int years - (if bought.Serial = first.Serial then 1 else 0)
    let halving = periods - 2 >= 1
    let mutable result = if period = 0 then Some charge0 else None
    let mutable value = cost - charge0
    let mutable half = 0.0
    let mutable n = 0
    while result.IsNone && n < period do
        n <- n + 1
        let charge =
            if value < salvage then 0.0
            elif halving && n > periods - 1 then 0.0
            elif halving && n >= periods - 2 then
                if half = 0.0 then half <- roundCharge (value * 0.5)
                half
            else roundCharge (r * value)
        if n = period then result <- Some charge
        value <- if halving && n >= periods - 2 then value - charge else value - r * value
    defaultArg result 0.0

// --- The functions -----------------------------------------------------------------------

/// A COUP* function: settlement, maturity, frequency, basis.
let private coupon (name: string) (f: Args -> Date -> Date -> int -> int -> Schedule -> CellValue) =
    scalar name 3 4 "|S" (fun a -> res {
        let! settlement, maturity = term a
        let! frequency = frequencyAt a 2
        let! basis = basisAt a 3
        do! check (settlement.Serial < maturity.Serial) Err.num
        return f a settlement maturity frequency basis (schedule settlement maturity frequency) })

let functions : FunctionDef list =
    [ coupon "COUPDAYBS" (fun _ settlement _ _ basis s -> CellValue.Number (coupDaysBS basis settlement s))
      coupon "COUPDAYS" (fun _ settlement maturity frequency basis _ -> CellValue.Number (coupDays basis frequency settlement maturity))
      coupon "COUPDAYSNC" (fun _ settlement _ _ basis s -> CellValue.Number (coupDaysNC basis settlement s))
      coupon "COUPNCD" (fun a _ _ _ _ s -> serialOut a s.Next)
      coupon "COUPPCD" (fun a _ _ _ _ s -> serialOut a s.Previous)
      coupon "COUPNUM" (fun _ _ _ _ _ s -> CellValue.Number (float s.Remaining))
      scalar "PRICE" 6 7 "|S" (fun a -> res {
        let! settlement, maturity = term a
        let! rate = a.AtpNumber 2
        let! yld = a.AtpNumber 3
        let! redemption = a.AtpNumber 4
        let! frequency = frequencyAt a 5
        let! basis = basisAt a 6
        do! check (settlement.Serial < maturity.Serial && rate >= 0.0 && yld >= 0.0 && redemption > 0.0) Err.num
        return! numOk (price settlement maturity rate yld redemption frequency basis) })
      scalar "YIELD" 6 7 "|S" (fun a -> res {
        let! settlement, maturity = term a
        let! rate = a.AtpNumber 2
        let! pr = a.AtpNumber 3
        let! redemption = a.AtpNumber 4
        let! frequency = frequencyAt a 5
        let! basis = basisAt a 6
        do! check (settlement.Serial < maturity.Serial && rate >= 0.0 && pr > 0.0 && redemption > 0.0) Err.num
        match yieldOf settlement maturity rate pr redemption frequency basis with
        | Some y -> return! numOk y
        | None -> return! fail Err.num })
      scalar "DURATION" 5 6 "|S" (fun a -> res {
        let! settlement, maturity = term a
        let! coupon = a.AtpNumber 2
        let! yld = a.AtpNumber 3
        let! frequency = frequencyAt a 4
        let! basis = basisAt a 5
        do! check (settlement.Serial < maturity.Serial && coupon >= 0.0 && yld >= 0.0) Err.num
        return! numOk (duration settlement maturity coupon yld frequency basis) })
      scalar "MDURATION" 5 6 "|S" (fun a -> res {
        let! settlement, maturity = term a
        let! coupon = a.AtpNumber 2
        let! yld = a.AtpNumber 3
        let! frequency = frequencyAt a 4
        let! basis = basisAt a 5
        do! check (settlement.Serial < maturity.Serial && coupon >= 0.0 && yld >= 0.0) Err.num
        return! numOk (duration settlement maturity coupon yld frequency basis / (yld / float frequency + 1.0)) })
      scalar "ACCRINT" 6 8 "|S" (fun a -> res {
        let! issue = dateAt a 0
        let! firstInterest = dateAt a 1
        let! settlement = dateAt a 2
        let! rate = a.AtpNumber 3
        let! par = a.AtpNumberOr 4 1000.0
        let! frequency = frequencyAt a 5
        let! basis = basisAt a 6
        let! fromIssue = a.BoolOr 7 true
        do! check (issue.Serial < settlement.Serial && rate > 0.0 && par > 0.0) Err.num
        return! numOk (accrint issue firstInterest settlement rate par frequency basis fromIssue) })
      scalar "ODDFPRICE" 8 9 "|S" (fun a -> res {
        let! settlement, maturity = term a
        let! issue = dateAt a 2
        let! firstCoupon = dateAt a 3
        let! rate = a.AtpNumber 4
        let! yld = a.AtpNumber 5
        let! redemption = a.AtpNumber 6
        let! frequency = frequencyAt a 7
        let! basis = basisAt a 8
        do! check (issue.Serial < settlement.Serial && settlement.Serial < firstCoupon.Serial && firstCoupon.Serial < maturity.Serial) Err.num
        do! check (rate >= 0.0 && yld >= 0.0 && redemption > 0.0 && couponsAfter firstCoupon maturity frequency >= 0) Err.num
        return! numOk (oddFirstPrice settlement maturity issue firstCoupon rate yld redemption frequency basis) })
      scalar "ODDFYIELD" 8 9 "|S" (fun a -> res {
        let! settlement, maturity = term a
        let! issue = dateAt a 2
        let! firstCoupon = dateAt a 3
        let! rate = a.AtpNumber 4
        let! pr = a.AtpNumber 5
        let! redemption = a.AtpNumber 6
        let! frequency = frequencyAt a 7
        let! basis = basisAt a 8
        do! check (issue.Serial < settlement.Serial && settlement.Serial < firstCoupon.Serial && firstCoupon.Serial < maturity.Serial) Err.num
        do! check (rate >= 0.0 && pr > 0.0 && redemption > 0.0 && couponsAfter firstCoupon maturity frequency >= 0) Err.num
        match solve (fun y -> oddFirstPrice settlement maturity issue firstCoupon rate y redemption frequency basis - pr) 0.05 with
        | Some y -> return! numOk y
        | None -> return! fail Err.num })
      scalar "ODDLPRICE" 7 8 "|S" (fun a -> res {
        let! settlement, maturity = term a
        let! lastInterest = dateAt a 2
        let! rate = a.AtpNumber 3
        let! yld = a.AtpNumber 4
        let! redemption = a.AtpNumber 5
        let! frequency = frequencyAt a 6
        let! basis = basisAt a 7
        do! check (lastInterest.Serial < settlement.Serial && settlement.Serial < maturity.Serial) Err.num
        do! check (rate >= 0.0 && yld >= 0.0 && redemption > 0.0) Err.num
        return! numOk (oddLastPrice settlement maturity lastInterest rate yld redemption frequency basis) })
      scalar "ODDLYIELD" 7 8 "|S" (fun a -> res {
        let! settlement, maturity = term a
        let! lastInterest = dateAt a 2
        let! rate = a.AtpNumber 3
        let! pr = a.AtpNumber 4
        let! redemption = a.AtpNumber 5
        let! frequency = frequencyAt a 6
        let! basis = basisAt a 7
        do! check (lastInterest.Serial < settlement.Serial && settlement.Serial < maturity.Serial) Err.num
        do! check (rate >= 0.0 && pr > 0.0 && redemption > 0.0) Err.num
        return! numOk (oddLastYield settlement maturity lastInterest rate pr redemption frequency basis) })
      scalar "AMORDEGRC" 6 7 "|S" (fun a -> res {
        let! cost = a.AtpNumber 0
        let! bought = dateAt a 1
        let! first = dateAt a 2
        let! salvage = a.AtpNumber 3
        let! period = a.AtpNumber 4
        let! rate = a.AtpNumber 5
        let! basis = basisAt a 6
        do! check (cost >= 0.0 && salvage >= 0.0 && salvage <= cost && period >= 0.0 && rate > 0.0 && bought.Serial <= first.Serial && basis <> 2) Err.num
        do! check (1.0 / rate > 2.0) Err.num
        return! numOk (amordegrc cost bought first salvage (int period) rate basis) })
      scalar "AMORLINC" 6 7 "|S" (fun a -> res {
        let! cost = a.AtpNumber 0
        let! bought = dateAt a 1
        let! first = dateAt a 2
        let! salvage = a.AtpNumber 3
        let! period = a.AtpNumber 4
        let! rate = a.AtpNumber 5
        let! basis = basisAt a 6
        do! check (cost >= 0.0 && salvage >= 0.0 && salvage <= cost && period >= 0.0 && rate > 0.0 && bought.Serial <= first.Serial && basis <> 2) Err.num
        return! numOk (amorlinc cost bought first salvage (int period) rate basis) }) ]

let install () = registerFunctions moduleName functions
