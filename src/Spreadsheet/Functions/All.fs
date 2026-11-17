/// .NET: every module of the function table, installed at start-up (Formula
/// reads `installed` the first time it looks a function up). The browser
/// never imports this file: there each module is loaded when it is needed.
module Rbt.Spreadsheet.Functions.All

let installed : bool =
    FnMath.install ()
    FnDatabase.install ()
    FnEngineering.install ()
    FnStatistical.install ()
    FnFinancial.install ()
    FnText.install ()
    FnArrays.install ()
    FnInformation.install ()
    true
