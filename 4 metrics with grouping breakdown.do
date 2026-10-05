clear all
set more off
*========================= USER INPUTS =========================
* ===================== USER INPUTS =====================
local filepath "C:\Users\wb367460\OneDrive - WBG\Working Files_1210-2021\WDI\Other request and tests\icp_tk_extrapolation\key statistics check\P_Data_Extract_From_WDI_Database_Archives_PPP benchmark changes.xlsx"
local sheetname "2021_GDP"

* Dataset variables
local var_time      "Time"
local var_series    "SeriesCode"
local var_income    "IncomeClassification"
local var_icpregion "ICPregion"
local var_quintile  "Quintile"

* Vintage tags (detect column names containing these yyyymm tokens)
local tag_extrapolation "2024 Mar [202403]"
local tag_benchmark "2024 May [202405]"
local CURRENT_IS        "BENCHMARK"   // or "EXTRAPOLATION"

* Optional filters (leave blank to skip)
local year .
local scode ""

*========================= LOGGING (quick) =========================
capture log close _all
local logpath "`c(pwd)'\log_`sheetname'_4metrics_obs.txt"
log using "`logpath'", text replace name(log4_`sheetname')

*========================= IMPORT =========================
import excel using "`filepath'", sheet("`sheetname'") firstrow clear
tempfile original_data

* MINIMAL FIX: create persistent normalized Quintile helper once
capture drop __qnorm_all
gen str20 __qnorm_all = lower(strtrim(`var_quintile'))

save `original_data', replace

*========================= IDENTIFY VINTAGE COLUMNS (by yyyymm tokens) =========================
local tok_extra ""
local tok_bench ""
if regexm("`tag_extrapolation'", "([0-9]{6})") local tok_extra = regexs(1)
if regexm("`tag_benchmark'", "([0-9]{6})") local tok_bench = regexs(1)
unab allvars: _all
local v_tag_extrapolation ""
local v_tag_benchmark ""
foreach v of local allvars {
    if "`v_tag_extrapolation'"=="" & regexm("`v'", "`tok_extra'") local v_tag_extrapolation "`v'"
    if "`v_tag_benchmark'"=="" & regexm("`v'", "`tok_bench'") local v_tag_benchmark "`v'"
}
di as txt "Detected EXTRAP: `v_tag_extrapolation' | BENCH: `v_tag_benchmark'"

*========================= MAP ROLES =========================
local var_current ""
local var_previous ""
if "`CURRENT_IS'" == "BENCHMARK" {
    local var_current "`v_tag_benchmark'"
    local var_previous "`v_tag_extrapolation'"
}
else {
    local var_current "`v_tag_extrapolation'"
    local var_previous "`v_tag_benchmark'"
}
di as txt "Current=`var_current' | Previous=`var_previous'"

*========================= DEFINE CASES =========================
local all_cases All LM H UM L "Group 5" "Group 4" "Group 3" "Group 2" "Group 1" a_AFR b_ASI c_CIS e_LAT f_CAR g_WAS h_SPP
local income_cases LM H UM L
local quintile_cases "Group 5" "Group 4" "Group 3" "Group 2" "Group 1"
local icpregion_cases a_AFR b_ASI c_CIS e_LAT f_CAR g_WAS h_SPP

capture program drop _is_in_list
program define _is_in_list, rclass
    syntax, word(string) list(string)
    local padded " `list' "
    local target " `word' "
    if strpos("`padded'", "`target'")>0 return scalar isin = 1
    else return scalar isin = 0
end

*========================= QUICK DIAGNOSTIC FOR QUINTILE STORAGE =========================
di as txt "Distinct raw values of `var_quintile' (first 25 shown):"
capture noisily tab `var_quintile' if _n<=., m

*========================= SAFE INIT: RESULTS DATASET (long) =========================
tempfile results4
preserve
    clear
    set obs 0
    gen str20 metric = ""
    gen str20 case = ""
    gen double value = .
    save `results4', replace
restore

*========================= MAIN LOOP OVER 17 CASES =========================
foreach case of local all_cases {
    di as txt _n(2)
    di as txt "-----------------------------------------------------------------"
    di as txt "Processing case: `case'"
    di as txt "-----------------------------------------------------------------"

    use `original_data', clear

    * Optional filters
    if "`scode'" != "" {
        gen __series_lc = lower(trim(`var_series'))
        keep if __series_lc == lower(trim("`scode'"))
        drop __series_lc
    }
    if !missing(`year') keep if `var_time' == `year'

    * Route to single filter
    local income ""
    local quintile ""
    local icpregion ""
    if "`case'"!="All" {
        quietly _is_in_list, word("`case'") list("`income_cases'")
        if r(isin)==1 local income "`case'"
        else {
            quietly _is_in_list, word("`case'") list("`quintile_cases'")
            if r(isin)==1 local quintile "`case'"
            else {
                quietly _is_in_list, word("`case'") list("`icpregion_cases'")
                if r(isin)==1 local icpregion "`case'"
            }
        }
    }

    * Robust Income filter
    if "`income'"!="" {
        gen __income_norm = lower(trim(`var_income'))
        keep if __income_norm == lower(trim("`income'"))
        drop __income_norm
    }

    * Robust ICP region filter
    if "`icpregion'"!="" {
        gen __icp_norm = lower(trim(`var_icpregion'))
        keep if __icp_norm == lower(trim("`icpregion'"))
        drop __icp_norm
    }

    * Robust Quintile filter (numeric or string)
    if "`quintile'"!="" {
        capture confirm numeric variable `var_quintile'
        if _rc==0 {
            * Numeric quintile: map Group 1..5 -> 1..5
            local qnum = .
            if "`quintile'"=="Group 1" local qnum = 1
            if "`quintile'"=="Group 2" local qnum = 2
            if "`quintile'"=="Group 3" local qnum = 3
            if "`quintile'"=="Group 4" local qnum = 4
            if "`quintile'"=="Group 5" local qnum = 5
            if `qnum'<. {
                quietly count
                local N0 = r(N)
                keep if `var_quintile' == `qnum'
                quietly count
                di as txt "Quintile (numeric) `quintile' -> `qnum' | N: `N0' -> " %9.0f r(N)
            }
            else {
                di as err "Quintile mapping failed (numeric var) for case: `quintile'"
            }
        }
        else {
            * MINIMAL FIX: use persistent normalized helper created after import
            local qnorm = lower(trim("`quintile'"))
            quietly count
            local N0 = r(N)
            keep if __qnorm_all == "`qnorm'"
            quietly count
            di as txt "Quintile (string) target=`qnorm' | N: `N0' -> " %9.0f r(N)
        }
    }

    * Base sample used for metrics
    keep if !missing(`var_current') & !missing(`var_previous')
    count
    di as txt "Rows after filtering: " %9.0f r(N)
    scalar N_base = _N

    * Empty-case: append missing for 5 rows (MAE, MSE, RMSE, obs, R2(pred))
    if _N==0 {
        preserve
            use `results4', clear
            set obs `= _N + 5'
            replace metric = "MAE" in -5
            replace case = "`case'" in -5
            replace value = . in -5
            replace metric = "MSE" in -4
            replace case = "`case'" in -4
            replace value = . in -4
            replace metric = "RMSE" in -3
            replace case = "`case'" in -3
            replace value = . in -3
            replace metric = "obs" in -2
            replace case = "`case'" in -2
            replace value = . in -2
            replace metric = "R2(pred)" in -1
            replace case = "`case'" in -1
            replace value = . in -1
            save `results4', replace
        restore
        continue
    }

    * Ensure numeric
    capture confirm numeric variable `var_current'
    if _rc destring `var_current', replace force
    capture confirm numeric variable `var_previous'
    if _rc destring `var_previous', replace force

    * Error terms
    gen double err = `var_current' - `var_previous'
    gen double abserr = abs(err)
    gen double sqerr = err^2

    * 4 metrics
    quietly summarize abserr
    scalar MAE = r(mean)
    quietly summarize sqerr
    scalar MSE = r(mean)
    scalar RMSE = sqrt(MSE)

    quietly summarize `var_current'
    scalar Ncur = r(N)
    scalar VarCur = r(Var)
    scalar SST = VarCur * max(Ncur-1,0)
    quietly summarize sqerr
    scalar SSE = r(mean) * r(N)
    scalar R2_pred = .
    if SST>0 & r(N)>0 scalar R2_pred = 1 - (SSE / SST)

    * Append rows for this case (5 rows)
    preserve
        use `results4', clear
        set obs `= _N + 5'
        replace metric = "MAE" in -5
        replace case = "`case'" in -5
        replace value = MAE in -5
        replace metric = "MSE" in -4
        replace case = "`case'" in -4
        replace value = MSE in -4
        replace metric = "RMSE" in -3
        replace case = "`case'" in -3
        replace value = RMSE in -3
        replace metric = "obs" in -2
        replace case = "`case'" in -2
        replace value = N_base in -2
        replace metric = "R2(pred)" in -1
        replace case = "`case'" in -1
        replace value = R2_pred in -1
        save `results4', replace
    restore
}

*========================= ASSEMBLE AND EXPORT FINAL OUTPUT TABLE (5 rows) =========================
use `results4', clear
keep metric case value
gen str40 casekey = case
replace casekey = subinstr(casekey, " ", "_", .)
replace casekey = subinstr(casekey, "-", "_", .)
replace casekey = subinstr(casekey, "(", "", .)
replace casekey = subinstr(casekey, ")", "", .)
replace casekey = subinstr(casekey, ":", "", .)
replace casekey = subinstr(casekey, "/", "_", .)
replace casekey = subinstr(casekey, "\", "_", .)

collapse (mean) value, by(metric casekey)
reshape wide value, i(metric) j(casekey) string

capture confirm variable valueAll
if _rc==0 {
    rename valueAll All
    rename valueLM LM
    rename valueH H
    rename valueUM UM
    rename valueL L
    rename valueGroup_5 Group_5
    rename valueGroup_4 Group_4
    rename valueGroup_3 Group_3
    rename valueGroup_2 Group_2
    rename valueGroup_1 Group_1
    rename valuea_AFR a_AFR
    rename valueb_ASI b_ASI
    rename valuec_CIS c_CIS
    rename valuee_LAT e_LAT
    rename valuef_CAR f_CAR
    rename valueg_WAS g_WAS
    rename valueh_SPP h_SPP
}
else {
    capture confirm variable valueall
    if _rc==0 rename valueall All
    capture confirm variable valuelm
    if _rc==0 rename valuelm LM
    capture confirm variable valueh
    if _rc==0 rename valueh H
    capture confirm variable valueum
    if _rc==0 rename valueum UM
    capture confirm variable valuel
    if _rc==0 rename valuel L
    capture confirm variable valuegroup_5
    if _rc==0 rename valuegroup_5 Group_5
    capture confirm variable valuegroup_4
    if _rc==0 rename valuegroup_4 Group_4
    capture confirm variable valuegroup_3
    if _rc==0 rename valuegroup_3 Group_3
    capture confirm variable valuegroup_2
    if _rc==0 rename valuegroup_2 Group_2
    capture confirm variable valuegroup_1
    if _rc==0 rename valuegroup_1 Group_1
    capture confirm variable valuea_afr
    if _rc==0 rename valuea_afr a_AFR
    capture confirm variable valueb_asi
    if _rc==0 rename valueb_asi b_ASI
    capture confirm variable valuec_cis
    if _rc==0 rename valuec_cis c_CIS
    capture confirm variable valuee_lat
    if _rc==0 rename valuee_lat e_LAT
    capture confirm variable valuef_car
    if _rc==0 rename valuef_car f_CAR
    capture confirm variable valueg_was
    if _rc==0 rename valueg_was g_WAS
    capture confirm variable valueh_spp
    if _rc==0 rename valueh_spp h_SPP
}

rename metric Metrics
* Row order
gen byte _ord = .
replace _ord = 1 if Metrics=="MAE"
replace _ord = 2 if Metrics=="MSE"
replace _ord = 3 if Metrics=="RMSE"
replace _ord = 4 if Metrics=="obs"
replace _ord = 5 if Metrics=="R2(pred)"
sort _ord Metrics
drop _ord

order Metrics All LM H UM L Group_5 Group_4 Group_3 Group_2 Group_1 a_AFR b_ASI c_CIS e_LAT f_CAR g_WAS h_SPP

local outxlsx "`c(pwd)'\four_metrics_with_obs_`sheetname'.xlsx"
export excel using "`outxlsx'", firstrow(varlabels) replace
di as txt "Exported 5-row (4 metrics + obs) table to: `outxlsx'"

capture log close log4_`sheetname'
