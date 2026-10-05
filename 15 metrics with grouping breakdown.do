clear all
set more off

* ============================ USER INPUTS ===================================
//local filepath "C:\Users\wb367460\OneDrive - WBG\Working Files_1210-2021\WDI\Other request and tests\icp_tk_extrapolation\key statistics check\P_Data_Extract_From_WDI_Database_Archives_PPP benchmark changes.xlsx"
local filepath "C:\Users\wb589699\OneDrive - WBG\Desktop\WDIvsBenchmark_analysis\P_Data_Extract_From_WDI_Database_Archives_PPP benchmark replacement in WDI.xlsx"

local sheetname "2021_GDP"

* Dataset variables
local var_time "Time"
local var_series "SeriesCode"
local var_income "IncomeClassification"
local var_icpregion "ICPregion"
local var_quintile "Quintile"

* Vintage tags (detect column names containing these yyyymm tokens)
local tag_extrapolation "2024 Mar [202403]"
local tag_benchmark "2024 May [202405]"
local CURRENT_IS "BENCHMARK" // or "EXTRAPOLATION"

* Optional filters (leave blank to skip)
local year .
local scode ""

* ============================ LOGGING (quick) ===============================
capture log close _all
local logpath "`c(pwd)'\log_`sheetname'_15metrics_obs.txt"
log using "`logpath'", text replace name(log15_`sheetname')

* ============================ IMPORT ========================================
import excel using "`filepath'", sheet("`sheetname'") firstrow clear
tempfile original_data
save `original_data', replace

* =================== IDENTIFY VINTAGE COLUMNS (by yyyymm tokens) =========
local tok_extra ""
local tok_bench ""
if regexm("`tag_extrapolation'", "([0-9]{6})") local tok_extra = regexs(1)
if regexm("`tag_benchmark'", "([0-9]{6})") local tok_bench = regexs(1)

unab allvars: _all
local v_tag_extrapolation ""
local v_tag_benchmark ""
foreach v of local allvars {
    if `"`v_tag_extrapolation'"' == "" & regexm("`v'", "`tok_extra'") local v_tag_extrapolation "`v'"
    if `"`v_tag_benchmark'"' == "" & regexm("`v'", "`tok_bench'") local v_tag_benchmark "`v'"
}
di as txt "Detected EXTRAP: `v_tag_extrapolation' | BENCH: `v_tag_benchmark'"

* ============================ MAP ROLES =====================================
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

* ============================ DEFINE CASES ==================================
local all_cases All LM H UM L "Group 5" "Group 4" "Group 3" "Group 2" "Group 1" a_AFR b_ASI c_CIS e_LAT f_CAR g_WAS h_SPP
local income_cases LM H UM L
local quintile_cases "Group 5" "Group 4" "Group 3" "Group 2" "Group 1"
local icpregion_cases a_AFR b_ASI c_CIS e_LAT f_CAR g_WAS h_SPP

capture program drop _is_in_list
program define _is_in_list, rclass
    syntax, word(string) list(string)
    local padded " `list' "
    local target " `word' "
    if strpos("`padded'", "`target'") > 0 return scalar isin = 1
    else return scalar isin = 0
end

* ===================== SAFE INIT: RESULTS DATASET (long) ====================
tempfile results15
preserve
    clear
    set obs 0
    gen str30 metric = ""
    gen str20 case = ""
    gen double value = .
    save `results15', replace
restore

* ======================= MAIN LOOP OVER 17 CASES ==========================
foreach case of local all_cases {
    di as txt _n(2)
    di as txt "Processing case: `case'"
    di as txt "----------------------------------------------------------------"

    use `original_data', clear

    * --- CORRECTED ORDER OF OPERATIONS ---
    
    * 1. Apply universal filters first to establish the base working dataset.
    if "`scode'" != "" {
        gen __series_lc = lower(trim(`var_series'))
        keep if __series_lc == lower(trim("`scode'"))
        drop __series_lc
    }
    if !missing(`year') keep if `var_time' == `year'

    * Establish the base sample of rows with valid data for both vintages.
    keep if !missing(`var_current') & !missing(`var_previous')

    * 2. Now, apply the specific filter for the current case (if not "All").
    if "`case'" != "All" {
        
        quietly _is_in_list, word("`case'") list("`income_cases'")
        if r(isin)==1 {
            gen __inc = lower(trim(`var_income'))
            keep if __inc == lower(trim("`case'"))
            drop __inc
        }
        else {
            quietly _is_in_list, word("`case'") list("`quintile_cases'")
            if r(isin)==1 {
                gen __q = lower(trim(`var_quintile'))
                keep if __q == lower(trim("`case'"))
                drop __q
            }
            else {
                quietly _is_in_list, word("`case'") list("`icpregion_cases'")
                if r(isin)==1 {
                    gen __icp = lower(trim(`var_icpregion'))
                    keep if __icp == lower(trim("`case'"))
                    drop __icp
                }
            }
        }
    }
    * --- END OF CORRECTION ---

    count
    di as txt "Rows after filtering: " %9.0f r(N)
    scalar N_base = _N

    * If empty, append missing and continue
    if _N==0 {
        preserve
            use `results15', clear
            set obs `=_N+16'
            replace metric = "MAPE:" in -16
            replace case = "`case'" in -16
            replace value = . in -16
            replace metric = "sMAPE:" in -15
            replace case = "`case'" in -15
            replace value = . in -15
            replace metric = "RMSPE:" in -14
            replace case = "`case'" in -14
            replace value = . in -14
            replace metric = "obs" in -13
            replace case = "`case'" in -13
            replace value = . in -13
            replace metric = "Log MAE:" in -12
            replace case = "`case'" in -12
            replace value = . in -12
            replace metric = "Log RMSE:" in -11
            replace case = "`case'" in -11
            replace value = . in -11
            replace metric = "Log R2:" in -10
            replace case = "`case'" in -10
            replace value = . in -10
            replace metric = "Log Cor:" in -9
            replace case = "`case'" in -9
            replace value = . in -9
            replace metric = "MLE:" in -8
            replace case = "`case'" in -8
            replace value = . in -8
            replace metric = "Med Ratio:" in -7
            replace case = "`case'" in -7
            replace value = . in -7
            replace metric = "Cov10:" in -6
            replace case = "`case'" in -6
            replace value = . in -6
            replace metric = "Cov20:" in -5
            replace case = "`case'" in -5
            replace value = . in -5
            replace metric = "Med APE:" in -4
            replace case = "`case'" in -4
            replace value = . in -4
            replace metric = "Med |Log Err|:" in -3
            replace case = "`case'" in -3
            replace value = . in -3
            replace metric = "T-MAPE(5%):" in -2
            replace case = "`case'" in -2
            replace value = . in -2
            replace metric = "T-Log RMSE(5%):" in -1
            replace case = "`case'" in -1
            replace value = . in -1
            save `results15', replace
        restore
        continue
    }

    * Ensure numeric
    capture confirm numeric variable `var_current'
    if _rc != 0 destring `var_current', replace force
    capture confirm numeric variable `var_previous'
    if _rc != 0 destring `var_previous', replace force

    * Define A and P; components
    gen double A = `var_current' 
	gen double P = `var_previous'
   
    gen byte Apos = (A>0)
    gen double ape = .
    replace ape = abs((A - P) / A) if Apos

    gen double smape_denom = (abs(A) + abs(P)) / 2
    gen double smape = .
    replace smape = abs(A - P) / smape_denom if smape_denom>0

    gen byte log_ok = (A>0 & P>0)
    gen double lnA = ln(A) if log_ok
    gen double lnP = ln(P) if log_ok
    gen double ln_err = lnA - lnP if log_ok
    gen double ln_abserr = abs(ln_err) if log_ok
    gen double ln_sqerr = ln_err^2 if log_ok

    gen double ratio = P/A if Apos
    gen double pe_sq = ((A - P) / A)^2 if Apos

    * SCALARS FOR METRICS
    quietly summarize ape if Apos
    scalar MAPE = r(mean)
    quietly summarize pe_sq if Apos
    scalar RMSPE = sqrt(r(mean))
    quietly summarize smape if smape_denom>0
    scalar sMAPE = r(mean)

    quietly summarize ln_abserr if log_ok
    scalar Log_MAE = r(mean)
    quietly summarize ln_sqerr if log_ok
    scalar Log_RMSE = sqrt(r(mean))

    scalar Log_R2 = .
    quietly summarize lnA if log_ok
    scalar Nlog = r(N)
    scalar Var_lnA = r(Var)
    scalar SST_log = Var_lnA * max(Nlog-1,0)
    quietly summarize ln_sqerr if log_ok
    scalar SSE_log = r(mean) * r(N)
    if SST_log>0 & r(N)>0 {
        scalar Log_R2 = 1 - (SSE_log / SST_log)
    }

    scalar Log_Cor = .
    quietly corr lnA lnP if log_ok
    if r(N)>0 scalar Log_Cor = r(rho)

    quietly summarize ln_err if log_ok
    scalar MLE = -r(mean)
    quietly summarize ratio if Apos, detail
    scalar Med_Ratio = r(p50)

    gen byte cov10 = (abs(ratio - 1) <= 0.10) if Apos
    gen byte cov20 = (abs(ratio - 1) <= 0.20) if Apos
    quietly summarize cov10 if Apos
    scalar Cov10 = r(mean)
    quietly summarize cov20 if Apos
    scalar Cov20 = r(mean)

    quietly summarize ape if Apos, detail
    scalar Med_APE = r(p50)
    quietly summarize ln_abserr if log_ok, detail
    scalar Med_AbsLog = r(p50)

    * Trimmed MAPE (5%), guarded
    scalar T_MAPE_5 = .
    preserve
        keep if Apos & !missing(ape)
        count
        local Nape = r(N)
        if `Nape' > 1 {
            sort ape
            local trimN = floor(`Nape'*0.05)
            local lo = `trimN' + 1
            local hi = `Nape' - `trimN'
            if `lo' <= `hi' {
                quietly summarize ape in `lo'/`hi'
                scalar T_MAPE_5 = r(mean)
            }
        }
    restore

    * Trimmed Log RMSE (5%), guarded
    scalar T_LogRMSE_5 = .
    preserve
        keep if log_ok & !missing(ln_sqerr)
        count
        local Nlna = r(N)
        if `Nlna' > 1 {
            sort ln_sqerr
            local trimN2 = floor(`Nlna'*0.05)
            local lo2 = `trimN2' + 1
            local hi2 = `Nlna' - `trimN2'
            if `lo2' <= `hi2' {
                quietly summarize ln_sqerr in `lo2'/`hi2'
                scalar T_LogRMSE_5 = sqrt(r(mean))
            }
        }
    restore

    * ================== SAFE APPEND: 15 metrics + obs =======================
    preserve
        use `results15', clear
        set obs `=_N+16'
        replace metric = "MAPE:" in -16
        replace case = "`case'" in -16
        replace value = MAPE in -16
        replace metric = "sMAPE:" in -15
        replace case = "`case'" in -15
        replace value = sMAPE in -15
        replace metric = "RMSPE:" in -14
        replace case = "`case'" in -14
        replace value = RMSPE in -14
        replace metric = "obs" in -13
        replace case = "`case'" in -13
        replace value = N_base in -13
        replace metric = "Log MAE:" in -12
        replace case = "`case'" in -12
        replace value = Log_MAE in -12
        replace metric = "Log RMSE:" in -11
        replace case = "`case'" in -11
        replace value = Log_RMSE in -11
        replace metric = "Log R2:" in -10
        replace case = "`case'" in -10
        replace value = Log_R2 in -10
        replace metric = "Log Cor:" in -9
        replace case = "`case'" in -9
        replace value = Log_Cor in -9
        replace metric = "MLE:" in -8
        replace case = "`case'" in -8
        replace value = MLE in -8
        replace metric = "Med Ratio:" in -7
        replace case = "`case'" in -7
        replace value = Med_Ratio in -7
        replace metric = "Cov10:" in -6
        replace case = "`case'" in -6
        replace value = Cov10 in -6
        replace metric = "Cov20:" in -5
        replace case = "`case'" in -5
        replace value = Cov20 in -5
        replace metric = "Med APE:" in -4
        replace case = "`case'" in -4
        replace value = Med_APE in -4
        replace metric = "Med |Log Err|:" in -3
        replace case = "`case'" in -3
        replace value = Med_AbsLog in -3
        replace metric = "T-MAPE(5%):" in -2
        replace case = "`case'" in -2
        replace value = T_MAPE_5 in -2
        replace metric = "T-Log RMSE(5%):" in -1
        replace case = "`case'" in -1
        replace value = T_LogRMSE_5 in -1
        save `results15', replace
    restore
}

* ======== ASSEMBLE AND EXPORT FINAL OUTPUT TABLE (16 rows) ================
use `results15', clear
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

* Rename variables to valid Stata names
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

* Enforce row order
gen byte _ord = .
replace _ord = 1 if Metrics=="MAPE:"
replace _ord = 2 if Metrics=="sMAPE:"
replace _ord = 3 if Metrics=="RMSPE:"
replace _ord = 4 if Metrics=="obs"
replace _ord = 5 if Metrics=="Log MAE:"
replace _ord = 6 if Metrics=="Log RMSE:"
replace _ord = 7 if Metrics=="Log R2:"
replace _ord = 8 if Metrics=="Log Cor:"
replace _ord = 9 if Metrics=="MLE:"
replace _ord = 10 if Metrics=="Med Ratio:"
replace _ord = 11 if Metrics=="Cov10:"
replace _ord = 12 if Metrics=="Cov20:"
replace _ord = 13 if Metrics=="Med APE:"
replace _ord = 14 if Metrics=="Med |Log Err|:"
replace _ord = 15 if Metrics=="T-MAPE(5%):"
replace _ord = 16 if Metrics=="T-Log RMSE(5%):"
sort _ord
drop _ord

* Column order
order Metrics All LM H UM L Group_5 Group_4 Group_3 Group_2 Group_1 a_AFR b_ASI c_CIS e_LAT f_CAR g_WAS h_SPP

* Export
local outxlsx "`c(pwd)'\fifteen_metrics_with_obs_`sheetname'.xlsx"
export excel using "`outxlsx'", firstrow(varlabels) replace
di as txt "Exported 16-row (15 metrics + obs) table to: `outxlsx'"

* ============================ CLOSE LOG =====================================
capture log close log15_`sheetname'
