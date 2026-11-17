/// Financial functions of the function table: depreciation, fractional dollar
/// prices, interest-rate conversions, Treasury bills and discounted securities.
/// (Coupon bonds - PRICE, YIELD, the COUP* family - come in a later module.)
module Rbt.Spreadsheet.Functions.FnFinancial

open System
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.Functions.Kit

let moduleName = "Financial"

/// The year fraction between two dates in a day-count basis, as YEARFRAC gives it.
let private yearFraction (a: Args) (start: float) (finish: float) (basis: float) =
    coerceNumber (a.Ev.Call "YEARFRAC" [ Number start; Number finish; Number basis ] a.Ctx)

/// A date argument: a serial, its fraction of a day dropped.
let private dateArg (a: Args) (i: int) = a.Number i |> Result.map Math.Floor

/// The day-count basis argument (0 when omitted): 0 to 4, else #NUM!.
let private basisArg (a: Args) (i: int) =
    a.NumberOr i 0.0 |> Result.bind (fun b -> let b = truncate b in if b < 0.0 || b > 4.0 then Result.Error Err.num else Ok b)

/// The declining-balance depreciation of a period (DDB; VDB's steps).
let private ddbOf (cost: float) (salvage: float) (life: float) (period: float) (factor: float) =
    let rate = factor / life
    let rate, oldValue =
        if rate >= 1.0 then 1.0, (if period = 1.0 then cost else 0.0)
        else rate, cost * Math.Pow (1.0 - rate, period - 1.0)
    let newValue = cost * Math.Pow (1.0 - rate, period)
    let depreciation = if newValue < salvage then oldValue - salvage else oldValue - newValue
    max 0.0 depreciation

/// VDB's depreciation of the first `period` periods, switching to straight
/// line when that gives more.
let private vdbInterval (cost: float) (salvage: float) (life: float) (life1: float) (period: float) (factor: float) =
    let intEnd = Math.Ceiling (period - 1e-12)
    let loopEnd = int intEnd
    let mutable total = 0.0
    let mutable remaining = cost - salvage
    let mutable nowStraight = false
    let mutable straight = 0.0
    for i in 1 .. loopEnd do
        let mutable term = 0.0
        if not nowStraight then
            let ddb = ddbOf cost salvage life (float i) factor
            straight <- remaining / (life1 - float (i - 1))
            if straight > ddb then
                term <- straight
                nowStraight <- true
            else
                term <- ddb
                remaining <- remaining - ddb
        else term <- straight
        if i = loopEnd then term <- term * (period + 1.0 - intEnd)
        total <- total + term
    total

let private vdb (cost: float) (salvage: float) (life: float) (start: float) (finish: float) (factor: float) (noSwitch: bool) =
    let near (x: float) (y: float) = abs (x - y) < 1e-12
    let intStart = Math.Floor (start + 1e-12)
    let intEnd = Math.Ceiling (finish - 1e-12)
    if noSwitch then
        let mutable total = 0.0
        for i in int intStart + 1 .. int intEnd do
            let mutable term = ddbOf cost salvage life (float i) factor
            if i = int intStart + 1 then term <- term * (min finish (intStart + 1.0) - start)
            elif i = int intEnd then term <- term * (finish + 1.0 - intEnd)
            total <- total + term
        total
    else
        let mutable part = 0.0
        if not (near start intStart) then
            let value = cost - vdbInterval cost salvage life life intStart factor
            part <- part + (start - intStart) * vdbInterval value salvage life (life - intStart) 1.0 factor
        if not (near finish intEnd) then
            let tempStart = intEnd - 1.0
            let value = cost - vdbInterval cost salvage life life tempStart factor
            part <- part + (intEnd - finish) * vdbInterval value salvage life (life - tempStart) (intEnd - tempStart) factor
        let cost = cost - vdbInterval cost salvage life life intStart factor
        vdbInterval cost salvage life (life - intStart) (intEnd - intStart) factor - part

/// The number of digits of a fraction's denominator (DOLLARDE, DOLLARFR).
let private denominatorScale (fraction: float) = Math.Pow (10.0, Math.Ceiling (Math.Log10 fraction))

/// The days from settlement to maturity of a Treasury bill: within a year, else #NUM!.
let private billDays (a: Args) =
    res {
        let! settlement = dateArg a 0
        let! maturity = dateArg a 1
        do! check (settlement < maturity) Err.num
        do! check (maturity - settlement <= 365.0) Err.num
        return maturity - settlement }

/// Settlement and maturity (arguments 0 and 1), checked, and the year fraction
/// between them in the basis of argument `basisAt`.
let private securityTerm (a: Args) (basisAt: int) =
    res {
        let! settlement = dateArg a 0
        let! maturity = dateArg a 1
        let! basis = basisArg a basisAt
        do! check (settlement < maturity) Err.num
        let! fraction = yearFraction a settlement maturity basis
        return settlement, maturity, basis, fraction }

let functions : FunctionDef list =
    [ scalar "DB" 4 5 "|S" (fun a -> res {
        let! cost = a.Number 0
        let! salvage = a.Number 1
        let! life = a.Number 2
        let! period = a.Number 3
        let! month = a.NumberOr 4 12.0
        let period = truncate period
        let month = truncate month
        do! check (cost >= 0.0 && salvage >= 0.0 && life > 0.0 && period > 0.0 && month >= 1.0 && month <= 12.0) Err.num
        do! check (period <= life + (if month < 12.0 then 1.0 else 0.0)) Err.num
        if cost = 0.0 then return CellValue.Number 0.0
        else
            let rate = Math.Round (1.0 - Math.Pow (salvage / cost, 1.0 / life), 3)
            let first = cost * rate * month / 12.0
            let result =
                if period = 1.0 then first
                else
                    let mutable total = first
                    let mutable current = 0.0
                    let mutable p = 2.0
                    while p <= min period life do
                        current <- (cost - total) * rate
                        total <- total + current
                        p <- p + 1.0
                    if period > life then (cost - total) * rate * (12.0 - month) / 12.0 else current
            return! numOk result })
      scalar "DDB" 4 5 "|S" (fun a -> res {
        let! cost = a.Number 0
        let! salvage = a.Number 1
        let! life = a.Number 2
        let! period = a.Number 3
        let! factor = a.NumberOr 4 2.0
        do! check (cost >= 0.0 && salvage >= 0.0 && life > 0.0 && period > 0.0 && period <= life && factor > 0.0) Err.num
        return! numOk (ddbOf cost salvage life period factor) })
      scalar "SYD" 4 4 "|S" (fun a -> res {
        let! cost = a.Number 0
        let! salvage = a.Number 1
        let! life = a.Number 2
        let! period = a.Number 3
        do! check (life > 0.0 && period > 0.0 && period <= life) Err.num
        return! numOk ((cost - salvage) * (life - period + 1.0) * 2.0 / (life * (life + 1.0))) })
      scalar "VDB" 5 7 "|S" (fun a -> res {
        let! cost = a.Number 0
        let! salvage = a.Number 1
        let! life = a.Number 2
        let! start = a.Number 3
        let! finish = a.Number 4
        let! factor = a.NumberOr 5 2.0
        let! noSwitch = a.BoolOr 6 false
        do! check (start >= 0.0 && finish >= start && finish <= life && cost >= 0.0 && salvage <= cost && salvage >= 0.0 && factor > 0.0 && life > 0.0) Err.num
        return! numOk (vdb cost salvage life start finish factor noSwitch) })
      scalar "DOLLARDE" 2 2 "|S" (fun a -> res {
        let! dollar = a.Number 0
        let! fraction = a.Number 1
        let fraction = truncate fraction
        do! check (fraction >= 0.0) Err.num
        do! check (fraction > 0.0) Err.div0
        let whole = truncate dollar
        return! numOk (whole + (dollar - whole) * denominatorScale fraction / fraction) })
      scalar "DOLLARFR" 2 2 "|S" (fun a -> res {
        let! dollar = a.Number 0
        let! fraction = a.Number 1
        let fraction = truncate fraction
        do! check (fraction >= 0.0) Err.num
        do! check (fraction > 0.0) Err.div0
        let whole = truncate dollar
        return! numOk (whole + (dollar - whole) * fraction / denominatorScale fraction) })
      scalar "EFFECT" 2 2 "|S" (fun a -> res {
        let! rate = a.Number 0
        let! periods = a.Number 1
        let periods = truncate periods
        do! check (rate > 0.0 && periods >= 1.0) Err.num
        return! numOk (Math.Pow (1.0 + rate / periods, periods) - 1.0) })
      scalar "NOMINAL" 2 2 "|S" (fun a -> res {
        let! rate = a.Number 0
        let! periods = a.Number 1
        let periods = truncate periods
        do! check (rate > 0.0 && periods >= 1.0) Err.num
        return! numOk (periods * (Math.Pow (1.0 + rate, 1.0 / periods) - 1.0)) })
      scalar "FVSCHEDULE" 2 2 "SR" (fun a -> res {
        let! principal = a.Number 0
        let rates = a.Flat 1
        match rates |> Array.tryPick (fun v -> match v with CellValue.Error e -> Some e | CellValue.String _ -> Some Err.value | _ -> None) with
        | Some e -> return! fail e
        | None ->
            let factor = rates |> Array.fold (fun product v -> match v with CellValue.Number r -> product * (1.0 + r) | _ -> product) 1.0
            return! numOk (principal * factor) })
      scalar "ISPMT" 4 4 "|S" (fun a -> res {
        let! rate = a.Number 0
        let! period = a.Number 1
        let! periods = a.Number 2
        let! pv = a.Number 3
        do! check (periods <> 0.0) Err.div0
        return! numOk (pv * rate * (period / periods - 1.0)) })
      scalar "PDURATION" 3 3 "|S" (fun a -> res {
        let! rate = a.Number 0
        let! pv = a.Number 1
        let! fv = a.Number 2
        do! check (rate > 0.0 && pv > 0.0 && fv > 0.0) Err.num
        return! numOk ((Math.Log fv - Math.Log pv) / Math.Log (1.0 + rate)) })
      scalar "RRI" 3 3 "|S" (fun a -> res {
        let! periods = a.Number 0
        let! pv = a.Number 1
        let! fv = a.Number 2
        do! check (periods > 0.0) Err.num
        do! check (pv <> 0.0) Err.num
        return! numOk (Math.Pow (fv / pv, 1.0 / periods) - 1.0) })
      scalar "TBILLEQ" 3 3 "|S" (fun a -> res {
        let! days = billDays a
        let! discount = a.Number 2
        do! check (discount > 0.0) Err.num
        if days <= 182.0 then return! numOk (365.0 * discount / (360.0 - discount * days))
        else
            let price = 100.0 * (1.0 - discount * days / 360.0)
            do! check (price > 0.0) Err.num
            let term = days / 365.0
            let root = term * term - (2.0 * term - 1.0) * (1.0 - 100.0 / price)
            do! check (root >= 0.0) Err.num
            return! numOk ((-term + Math.Sqrt root) / (term - 0.5)) })
      scalar "TBILLPRICE" 3 3 "|S" (fun a -> res {
        let! days = billDays a
        let! discount = a.Number 2
        do! check (discount > 0.0) Err.num
        let price = 100.0 * (1.0 - discount * days / 360.0)
        do! check (price > 0.0) Err.num
        return! numOk price })
      scalar "TBILLYIELD" 3 3 "|S" (fun a -> res {
        let! days = billDays a
        let! price = a.Number 2
        do! check (price > 0.0) Err.num
        return! numOk ((100.0 - price) / price * 360.0 / days) })
      scalar "DISC" 4 5 "|S" (fun a -> res {
        let! _, _, _, fraction = securityTerm a 4
        let! price = a.Number 2
        let! redemption = a.Number 3
        do! check (price > 0.0 && redemption > 0.0) Err.num
        return! numOk ((redemption - price) / redemption / fraction) })
      scalar "INTRATE" 4 5 "|S" (fun a -> res {
        let! _, _, _, fraction = securityTerm a 4
        let! investment = a.Number 2
        let! redemption = a.Number 3
        do! check (investment > 0.0 && redemption > 0.0) Err.num
        return! numOk ((redemption - investment) / investment / fraction) })
      scalar "RECEIVED" 4 5 "|S" (fun a -> res {
        let! _, _, _, fraction = securityTerm a 4
        let! investment = a.Number 2
        let! discount = a.Number 3
        do! check (investment > 0.0 && discount > 0.0) Err.num
        let denominator = 1.0 - discount * fraction
        do! check (denominator > 0.0) Err.num
        return! numOk (investment / denominator) })
      scalar "PRICEDISC" 4 5 "|S" (fun a -> res {
        let! _, _, _, fraction = securityTerm a 4
        let! discount = a.Number 2
        let! redemption = a.Number 3
        do! check (discount > 0.0 && redemption > 0.0) Err.num
        return! numOk (redemption - discount * redemption * fraction) })
      scalar "YIELDDISC" 4 5 "|S" (fun a -> res {
        let! _, _, _, fraction = securityTerm a 4
        let! price = a.Number 2
        let! redemption = a.Number 3
        do! check (price > 0.0 && redemption > 0.0) Err.num
        return! numOk ((redemption - price) / price / fraction) })
      scalar "PRICEMAT" 5 6 "|S" (fun a -> res {
        let! settlement, maturity, basis, toMaturity = securityTerm a 5
        let! issue = dateArg a 2
        let! rate = a.Number 3
        let! yld = a.Number 4
        do! check (rate >= 0.0 && yld >= 0.0 && issue < settlement) Err.num
        let! issued = yearFraction a issue maturity basis
        let! accrued = yearFraction a issue settlement basis
        return! numOk ((100.0 + issued * rate * 100.0) / (1.0 + toMaturity * yld) - accrued * rate * 100.0) })
      scalar "YIELDMAT" 5 6 "|S" (fun a -> res {
        let! settlement, maturity, basis, toMaturity = securityTerm a 5
        let! issue = dateArg a 2
        let! rate = a.Number 3
        let! price = a.Number 4
        do! check (rate >= 0.0 && price > 0.0 && issue < settlement) Err.num
        let! issued = yearFraction a issue maturity basis
        let! accrued = yearFraction a issue settlement basis
        let paid = price / 100.0 + accrued * rate
        return! numOk (((1.0 + issued * rate) - paid) / paid / toMaturity) })
      scalar "ACCRINTM" 3 5 "|S" (fun a -> res {
        let! issue = dateArg a 0
        let! settlement = dateArg a 1
        let! rate = a.Number 2
        let! par = a.NumberOr 3 1000.0
        let! basis = basisArg a 4
        do! check (rate > 0.0 && par > 0.0 && issue < settlement) Err.num
        let! fraction = yearFraction a issue settlement basis
        return! numOk (par * rate * fraction) }) ]

let install () = registerFunctions moduleName functions
