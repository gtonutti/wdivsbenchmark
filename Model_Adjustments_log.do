clear all
set more off

cd "C:\Users\wb589699\OneDrive - WBG\Desktop\WDIvsBenchmark_analysis"

use ICPvsWDI_2021, clear
append using ICPvsWDI_2017

merge m:1 Country_Code Time using WDI_tradedata, nogen keep(match)
* Year dummy for 2021
gen year2021 = (Time == 2021)


tabulate IncomeClassification, gen(inc_)
* This generates inc_1 inc_2 inc_3 inc_4 — check which maps to which group
* label them clearly
gen inc_H  = (IncomeClassification == "H")
gen inc_UM = (IncomeClassification == "UM")
gen inc_LM = (IncomeClassification == "LM")
gen inc_L  = (IncomeClassification == "L")


* Create interaction terms between pli_ext and each income group
gen pli_ext_H  = pli_ext * inc_H
gen pli_ext_UM = pli_ext * inc_UM
gen pli_ext_LM = pli_ext * inc_LM
gen pli_ext_L  = pli_ext * inc_L

* ============================================================
* STEP 1 — Generate log variables and year trend
* ============================================================
gen ln_gdppc =ln(GDPpcCurrent_USD)
gen ln_pli_bench = ln(pli_bench)
gen ln_pli_ext   = ln(pli_ext)
gen ln_pli_ext_gdppc = ln_pli_ext * ln_gdppc
gen pli_ext_gdppc = pli_ext * ln_gdppc

bys Country_Code Category (Time): gen year_trend=_n+1


* Cross validation approach: splitting my data into two folds, training and 
set seed 12345
gen random_draw = runiform()
bysort IncomeClassification Time Category (random_draw): gen rank = _n
bysort IncomeClassification Time Category: gen group_n = _N
gen fold5 = ceil(rank / group_n * 5)

* Check balance
tab IncomeClassification fold
tab year2021 fold


* ============================================================
* STEP 2 — Cross-validation loop
* ============================================================
* We will store RMSE and MAE for raw pli_ext and pli_corrected

gen pli_corrected_cv = .
gen pli_corrected_cv_norm =. 
gen ln_pli_corrected_cv = .
* Generatin an inflation estimate for th regression model
egen inflation=rowtotal(Inflation_2011 Inflation_2017)
gen log_inflation=ln(inflation)

gen obs_weight = 1
replace obs_weight = 2 if Time == 2021


levelsof Category, local(categ)

foreach c of local categ {
forvalues f = 1/5 {
    

    * Define training and validation folds
    local val_fold = `f'
    
	

    * --- Estimate model on training fold ---
      reg ln_pli_bench ln_pli_ext ln_pli_ext_gdppc ln_gdppc log_inflation year_trend [aweight=obs_weight] ///
        if  fold5 != `val_fold' & Category=="`c'", robust

    * Apply correction to validation fold
    replace ln_pli_corrected_cv = _b[_cons]                 ///
                  + _b[ln_pli_ext]       * ln_pli_ext        ///
                  + _b[ln_pli_ext_gdppc] * ln_pli_ext_gdppc  ///
                  + _b[ln_gdppc]         * ln_gdppc           ///
                  + _b[log_inflation]        * log_inflation           ///
                  + _b[year_trend]       * year_trend          ///
                  if fold5 == `val_fold' & Category == "`c'"
}

* Convert back to levels
replace  pli_corrected_cv = exp(ln_pli_corrected_cv) if Category == "`c'"
quietly summarize pli_corrected_cv if Country_Code == "USA" & Category == "`c'"
replace pli_corrected_cv_norm = pli_corrected_cv / r(mean) if Category == "`c'"
}

* ============================================================
* STEP 3 — Compare corrected vs raw pli_ext against benchmark
* ============================================================

* Generate error terms
gen pct_error_raw       = ((pli_ext             - pli_bench) / pli_bench) * 100
gen sq_pct_error_raw    = pct_error_raw^2
gen abs_pct_error_raw   = abs(pct_error_raw)

* Percentage errors — corrected (normalized)
gen pct_error_cor       = ((pli_corrected_cv_norm - pli_bench) / pli_bench) * 100
gen sq_pct_error_cor    = pct_error_cor^2
gen abs_pct_error_cor   = abs(pct_error_cor)

gen IncomeClass_label = cond(IncomeClassification == "H",  "High Income", ///
                        cond(IncomeClassification == "UM", "Upper Middle Income", ///
                        cond(IncomeClassification == "LM", "Lower Middle Income", ///
                        cond(IncomeClassification == "L",  "Low Income", ""))))
						
gen series="HHC" if Category == "PA.NUS.PRVT.PP"
replace series="GDP" if missing(series)

* ============================================================
* SETUP
* ============================================================
local series_list GDP HHC
local years 2017 2021
local income_groups H UM LM L

* ============================================================
* LOOP OVER SERIES AND YEARS
* ============================================================
foreach s of local series_list {
foreach y of local years {

    * --- Determine sheet name ---
    local sheetname "`s'_`y'"

    * --- Open file and sheet ---
    putexcel set "Modeladjustment_diagnostic_log.xlsx", sheet("`sheetname'") modify

    * --- Write header ---
    putexcel A1 = "Income Group"
    putexcel B1 = "RMSPE Raw"
    putexcel C1 = "RMSPE Corrected"
    putexcel D1 = "MAPE Raw"
    putexcel E1 = "MAPE Corrected"
    putexcel F1 = "RMSPE Change (%)"
    putexcel G1 = "MAPE Change (%)"

    putexcel A1:G1, bold border(bottom, medium)

    * --- Fill rows by income group ---
    local row = 2
    foreach inc of local income_groups {

        putexcel A`row' = "`inc'"

        * RMSPE raw
        quietly summarize sq_pct_error_raw if IncomeClassification == "`inc'" ///
            & series == "`s'" & Time == `y' & ICPregion != "d_EUO"
        local rmspe_raw = sqrt(r(mean))

        * RMSPE corrected
        quietly summarize sq_pct_error_cor if IncomeClassification == "`inc'" ///
            & series == "`s'" & Time == `y' & ICPregion != "d_EUO"
        local rmspe_cor = sqrt(r(mean))

        * MAPE raw
        quietly summarize abs_pct_error_raw if IncomeClassification == "`inc'" ///
            & series == "`s'" & Time == `y' & ICPregion != "d_EUO"
        local mape_raw = r(mean)

        * MAPE corrected
        quietly summarize abs_pct_error_cor if IncomeClassification == "`inc'" ///
            & series == "`s'" & Time == `y' & ICPregion != "d_EUO"
        local mape_cor = r(mean)

        * Percentage changes
        local rmspe_chg = round((`rmspe_cor' - `rmspe_raw') / `rmspe_raw' * 100, 0.01)
        local mape_chg  = round((`mape_cor'  - `mape_raw')  / `mape_raw'  * 100, 0.01)

        * Round for display
        local rmspe_raw = round(`rmspe_raw', 0.001)
        local rmspe_cor = round(`rmspe_cor', 0.001)
        local mape_raw  = round(`mape_raw',  0.001)
        local mape_cor  = round(`mape_cor',  0.001)

        putexcel B`row' = `rmspe_raw'
        putexcel C`row' = `rmspe_cor'
        putexcel D`row' = `mape_raw'
        putexcel E`row' = `mape_cor'
        putexcel F`row' = `rmspe_chg'
        putexcel G`row' = `mape_chg'

        local row = `row' + 1
    }

    * --- Mean row across all income groups ---
    putexcel A`row' = "All", bold border(top, medium)

    * RMSPE raw
    quietly summarize sq_pct_error_raw if series == "`s'" ///
        & Time == `y' & ICPregion != "d_EUO"
    local rmspe_raw_all = round(sqrt(r(mean)), 0.001)

    * RMSPE corrected
    quietly summarize sq_pct_error_cor if series == "`s'" ///
        & Time == `y' & ICPregion != "d_EUO"
    local rmspe_cor_all = round(sqrt(r(mean)), 0.001)

    * MAPE raw
    quietly summarize abs_pct_error_raw if series == "`s'" ///
        & Time == `y' & ICPregion != "d_EUO"
    local mape_raw_all = round(r(mean), 0.001)

    * MAPE corrected
    quietly summarize abs_pct_error_cor if series == "`s'" ///
        & Time == `y' & ICPregion != "d_EUO"
    local mape_cor_all = round(r(mean), 0.001)

    * Percentage changes
    local rmspe_chg_all = round((`rmspe_cor_all' - `rmspe_raw_all') / `rmspe_raw_all' * 100, 0.01)
    local mape_chg_all  = round((`mape_cor_all'  - `mape_raw_all')  / `mape_raw_all'  * 100, 0.01)

    putexcel B`row' = `rmspe_raw_all', bold
    putexcel C`row' = `rmspe_cor_all', bold
    putexcel D`row' = `mape_raw_all',  bold
    putexcel E`row' = `mape_cor_all',  bold
    putexcel F`row' = `rmspe_chg_all', bold
    putexcel G`row' = `mape_chg_all',  bold
	
		foreach col in A B C D E F G  {
    putexcel `col'1, overwrite border(bottom, thin, black)
	}

* --- Thick border after All Countries ---
	foreach col in A B C D E F G {
    putexcel `col'5, overwrite border(bottom, thin, black)
	}

}
}





levelsof Category, local(categ) 
foreach c of local categ {
	
* RMSE


quietly summarize sq_pct_error_raw if Category=="`c'"
local rmspe_raw = sqrt(r(mean)) 

quietly summarize sq_pct_error_cor if Category=="`c'"
local rmspe_cor = sqrt(r(mean))


* MAE
quietly summarize abs_pct_error_raw if Category=="`c'"
local mape_raw = r(mean)
quietly summarize abs_pct_error_cor if Category=="`c'"
local mape_cor = r(mean)

di              " `c'"
di "         Raw pli_ext    Corrected"
di "RMSPE:   `rmspe_raw'    `rmspe_cor'"
di "MAPE:    `mape_raw'     `mape_cor'"
di "========================================"


}


foreach ser in "PA.NUS.PPP" "PA.NUS.PRVT.PP" {
    di "=== `ser' - Non-EUO ==="
    quietly summarize sq_pct_error_raw if Category == "`ser'" & ICPregion != "d_EUO"
    local rmspe_raw = sqrt(r(mean))
    quietly summarize sq_pct_error_cor if Category == "`ser'" & ICPregion != "d_EUO"
    local rmspe_cor = sqrt(r(mean))
    quietly summarize abs_pct_error_raw if Category == "`ser'" & ICPregion != "d_EUO"
    local mape_raw = r(mean)
    quietly summarize abs_pct_error_cor if Category == "`ser'" & ICPregion != "d_EUO"
    local mape_cor = r(mean)
    di "RMSPE raw: `rmspe_raw'  corrected: `rmspe_cor'"
    di "MAPE  raw: `mape_raw'   corrected: `mape_cor'"
}



foreach y in "2017" "2021" {
	foreach ser in "PA.NUS.PPP" "PA.NUS.PRVT.PP" {
    di "=== `y' `ser' - Non-EUO ==="
    quietly summarize sq_pct_error_raw if Category == "`ser'" & ICPregion != "d_EUO" & Time==`y'
    local rmspe_raw = sqrt(r(mean))
    quietly summarize sq_pct_error_cor if Category == "`ser'" & ICPregion != "d_EUO" & Time==`y'
    local rmspe_cor = sqrt(r(mean))
    quietly summarize abs_pct_error_raw if Category == "`ser'" & ICPregion != "d_EUO" & Time==`y'
    local mape_raw = r(mean)
    quietly summarize abs_pct_error_cor if Category == "`ser'" & ICPregion != "d_EUO" & Time==`y'
    local mape_cor = r(mean)
    di "RMSPE raw: `rmspe_raw'  corrected: `rmspe_cor'"
    di "MAPE  raw: `mape_raw'   corrected: `mape_cor'"
}
}
* ============================================================
* STEP 4 — Break down performance by income group
* ============================================================
**# Bookmark #1
levelsof Category, local(categ) 
foreach c of local categ {
foreach inc in "H" "UM" "LM" "L" {
    quietly summarize sq_pct_error_raw       if IncomeClassification == "`inc'" & Category=="`c'" & ICPregion != "d_EUO"
    local rmse_raw_g = sqrt(r(mean))
    quietly summarize sq_pct_error_cor if IncomeClassification == "`inc'" & Category=="`c'" & ICPregion != "d_EUO"
    local rmse_cor_g = sqrt(r(mean))
    di "Income group `inc' — `c' RMSE raw: `rmse_raw_g'  corrected: `rmse_cor_g'"
	
	quietly summarize abs_pct_error_raw       if IncomeClassification == "`inc'" & Category=="`c'" & ICPregion != "d_EUO"
    local mae_raw_g = (r(mean))
    quietly summarize abs_pct_error_cor if IncomeClassification == "`inc'" & Category=="`c'" & ICPregion != "d_EUO"
    local mae_corrected_g = (r(mean))
    di "Income group `inc' — `c' MAE raw: `mae_raw_g'  corrected: `mae_corrected_g'"
}
}

//Model seems to work best with low income countries for both GDP and especially for HHC.
foreach y in "2017" "2021" {
levelsof Category, local(categ) 
foreach c of local categ {
foreach inc in "H" "UM" "LM" "L" { 
    quietly summarize sq_pct_error_raw       if IncomeClassification == "`inc'" & Category=="`c'" & ICPregion != "d_EUO" & Time==`y'
    local rmse_raw_g = sqrt(r(mean))
    quietly summarize sq_pct_error_cor if IncomeClassification == "`inc'" & Category=="`c'" & ICPregion != "d_EUO" & Time==`y'
    local rmse_cor_g = sqrt(r(mean))
    di "`y' Income group `inc' — `c' RMSE raw: `rmse_raw_g'  corrected: `rmse_cor_g'"
	
	quietly summarize abs_pct_error_raw       if IncomeClassification == "`inc'" & Category=="`c'" & ICPregion != "d_EUO" & Time==`y'
    local mae_raw_g = (r(mean))
    quietly summarize abs_pct_error_cor if IncomeClassification == "`inc'" & Category=="`c'" & ICPregion != "d_EUO" & Time==`y'
    local mae_corrected_g = (r(mean))
    di "`y' Income group `inc' — `c' MAE raw: `mae_raw_g'  corrected: `mae_corrected_g'"
}
}
}

forvalues f = 1/2 {
    di "=== Validation fold `f' ==="
    quietly summarize sq_pct_error_cor if ICPregion != "d_EUO" & fold == `f'
    di "RMSPE corrected: " sqrt(r(mean))
    quietly summarize abs_pct_error_cor if ICPregion != "d_EUO" & fold == `f'
    di "MAPE corrected: " r(mean)
}



gen pct_error_raw_abs = abs(pct_error_raw)
gen pct_error_cor_abs = abs(pct_error_cor)

twoway ///
    (scatter abs_pct_error_cor abs_pct_error_raw if ICPregion != "d_EUO" & Time==2021 & Category=="PA.NUS.PPP", mlabel(Country_Code) msize(small)) ///
    (function y = x, range(0 50) lpattern(dash) lcolor(gray)), ///
    xtitle("Absolute % Error — Raw Extrapolation") ///
    ytitle("Absolute % Error — Corrected") ///
    title("Prediction Error: Raw vs Corrected (Non-EUO Countries)") ///
    note("Points below the 45° line indicate improvement from correction") ///
    legend(order(2 "45° line") pos(5) ring(1))


*********************************************************************************************************************************************
***********************************************************PLOTTING ERRORS*******************************************************************
******************************************************************************************************************************************


foreach y in "2017" "2021" {
levelsof series, local(ser)

foreach s of local ser {
	
levelsof IncomeClass_label, local(incomes)

foreach inc of local incomes {

 * --- Compute axis range from pli_extrap distribution ---
    quietly summarize abs_pct_error_raw if ICPregion != "d_EUO" & Time==`y' & series=="`s'" & IncomeClass_label=="`inc'", detail
    local xmin = round(r(min), 0.01)
    local xmax = round(r(max) + 0.05, 0.01)
    local xmin_plot = `xmin' - 0.05

    local range = `xmax' - `xmin'
    local xq1  = round(`xmin' + `range' * 1/4, 0.01)
    local xmed = round(`xmin' + `range' * 2/4, 0.01)
    local xq3  = round(`xmin' + `range' * 3/4, 0.01)

    * --- Clean label (remove "c_" prefix if present) ---
    

    twoway ///
        (scatter abs_pct_error_cor abs_pct_error_raw if ICPregion != "d_EUO" & Time==`y' & series=="`s'" & IncomeClass_label=="`inc'", mlabel(Country_Code) msize(small)) ///
        (function y = x, range(`xmin_plot' `xmax') lpattern(dash) lcolor(gray)), ///
        xlabel(`xmin' `xq1' `xmed' `xq3' `xmax', format(%4.2f)) ///
        ylabel(`xmin' `xq1' `xmed' `xq3' `xmax', format(%4.2f)) ///
        xscale(range(`xmin_plot' `xmax')) yscale(range(`xmin_plot' `xmax')) ///
		plotregion(margin(zero)) ///
        graphregion(margin(t=8 b=2 l=2 r=8)) ///
		xtitle("Absolute % Error — Raw Extrapolation") ///
        ytitle("Absolute % Error — Corrected") ///
        title("`s' `y'' Prediction Error: Raw vs Corrected - `inc'") ///
        note("Points below the 45° line indicate improvement from correction") ///
        legend(order(2 "45° line") pos(5) ring(1))

    graph export "charts\Erros_`y'_`s'_`inc'.png", replace
}
}
}
	

	
foreach y in 2017 2021 {
levelsof series, local(ser)

foreach s of local ser {
	
levelsof IncomeClass_label, local(incomes)

foreach inc of local incomes {

    * --- Compute axis range across both pli_ext and pli_corrected ---
    quietly summarize pli_bench if ICPregion != "d_EUO" & Time == `y' & series == "`s'" & IncomeClass_label == "`inc'", detail
    local xmin = round(r(min), 0.01)
    local xmax = round(r(max) + 0.05, 0.01)

    * Also check pli_ext and pli_corrected don't exceed the range
    quietly summarize pli_ext if ICPregion != "d_EUO" & Time == `y' & series == "`s'" & IncomeClass_label == "`inc'", detail
    if round(r(min), 0.01) < `xmin' local xmin = round(r(min), 0.01)
    if round(r(max), 0.01) > `xmax' local xmax = round(r(max) + 0.05, 0.01)

    quietly summarize pli_corrected_cv_norm if ICPregion != "d_EUO" & Time == `y' & series == "`s'" & IncomeClass_label == "`inc'", detail
    if round(r(min), 0.01) < `xmin' local xmin = round(r(min), 0.01)
    if round(r(max), 0.01) > `xmax' local xmax = round(r(max) + 0.05, 0.01)

    local xmin_plot = `xmin' - 0.05
    local range = `xmax' - `xmin'
    local xq1  = round(`xmin' + `range' * 1/4, 0.01)
    local xmed = round(`xmin' + `range' * 2/4, 0.01)
    local xq3  = round(`xmin' + `range' * 3/4, 0.01)

    twoway ///
        (scatter pli_ext       pli_bench if ICPregion != "d_EUO" & Time == `y' & series == "`s'" & IncomeClass_label == "`inc'", ///
            mlabel(Country_Code) msize(small) mcolor(navy) mlabcolor(navy) msymbol(circle)) ///
        (scatter pli_corrected_cv_norm pli_bench if ICPregion != "d_EUO" & Time == `y' & series == "`s'" & IncomeClass_label == "`inc'", ///
            mlabel(Country_Code) msize(small) mcolor(orange) mlabcolor(orange) msymbol(triangle)) ///
        (function y = x, range(`xmin_plot' `xmax') lpattern(dash) lcolor(gray)), ///
        xlabel(`xmin' `xq1' `xmed' `xq3' `xmax', format(%4.2f)) ///
        ylabel(`xmin' `xq1' `xmed' `xq3' `xmax', format(%4.2f)) ///
        xscale(range(`xmin_plot' `xmax')) yscale(range(`xmin_plot' `xmax')) ///
        plotregion(margin(zero)) ///
        graphregion(margin(t=8 b=2 l=2 r=8)) ///
        xtitle("Benchmark PLI") ///
        ytitle("Extrapolated PLI") ///
        title("`s' `y' — Raw and Model-Adjusted PLIs vs Benchmark", size(small)) ///
        subtitle("`inc'", size(small)) ///
        note("Points on the 45° line indicate perfect prediction") ///
        legend(order(1 "Raw extrapolation" 2 "Model adjusted" 3 "45° line") pos(5) ring(1)) ///
        xsize(7) ysize(5)

    graph export "charts\PLIsrawandmodel_`y'_`s'_`inc'.png", replace
}
}
}

*********************************************************************************************************************************************
***********************************************************HIT RATE ANALYS*******************************************************************
******************************************************************************************************************************************


* ============================================================
* SETUP — define thresholds and groups
* ============================================================
local thresholds 5 10 15
local income_groups H UM LM L
local series_list GDP HHC
local years 2017 2021

* Sheet name mapping
local sheet_GDP_2017    "GDP_2017"
local sheet_GDP_2021    "GDP_2021"
local sheet_HHC_2017 "HHC_2017"
local sheet_HHC_2021 "HHC_2021"

* ============================================================
* LOOP OVER SERIES AND YEARS
* ============================================================
foreach c of local series_list {
foreach y of local years {

    local sheetname = "`sheet_`c'_`y''"

    * --- Open file and sheet ---
    putexcel set "Modeladjustment_hitrate_log.xlsx", sheet("`sheetname'") modify

    * --- Write header ---
    putexcel A1 = "Income Group"
    putexcel B1 = "Raw — Within 5%"
    putexcel C1 = "Corrected — Within 5%"
    putexcel D1 = "Raw — Within 10%"
    putexcel E1 = "Corrected — Within 10%"
    putexcel F1 = "Raw — Within 15%"
    putexcel G1 = "Corrected — Within 15%"

    * --- Format header row (bold) ---
    putexcel A1:G1, bold border(bottom, medium)

    * --- Fill rows by income group ---
    local row = 2
    foreach inc of local income_groups {

        putexcel A`row' = "`inc'"

        local col = 2   // start from column B
        foreach threshold of local thresholds {

            * Raw hit rate
            quietly count if IncomeClassification == "`inc'" ///
                & series == "`c'" & Time == `y'           ///
                & ICPregion != "d_EUO"                      ///
                & abs_pct_error_raw <= `threshold'
            local n_raw = r(N)

            quietly count if IncomeClassification == "`inc'" ///
                & series == "`c'" & Time == `y'           ///
                & ICPregion != "d_EUO"
            local n_total = r(N)

            local share_raw = round(`n_raw' / `n_total' * 100, 0.1)

            * Corrected hit rate
            quietly count if IncomeClassification == "`inc'" ///
                & series == "`c'" & Time == `y'           ///
                & ICPregion != "d_EUO"                      ///
                & abs_pct_error_cor <= `threshold'
            local n_cor = r(N)

            local share_cor = round(`n_cor' / `n_total' * 100, 0.1)

            * Write to Excel
            local col_raw = word("B C D E F G", `col' - 1)
            local col_cor = word("B C D E F G", `col')

            putexcel `col_raw'`row' = `share_raw'
            putexcel `col_cor'`row' = `share_cor'

            local col = `col' + 2
        }

        local row = `row' + 1
    }

    * --- Add total row ---
    putexcel A`row' = "All"

    local col = 2
    foreach threshold of local thresholds {

        quietly count if series == "`c'" & Time == `y' ///
            & ICPregion != "d_EUO"                       ///
            & abs_pct_error_raw <= `threshold'
        local n_raw = r(N)

        quietly count if series == "`c'" & Time == `y' ///
            & ICPregion != "d_EUO"
        local n_total = r(N)

        local share_raw = round(`n_raw' / `n_total' * 100, 0.1)

        quietly count if series == "`c'" & Time == `y' ///
            & ICPregion != "d_EUO"                       ///
            & abs_pct_error_cor <= `threshold'
        local n_cor = r(N)

        local share_cor = round(`n_cor' / `n_total' * 100, 0.1)

        local col_raw = word("B C D E F G", `col' - 1)
        local col_cor = word("B C D E F G", `col')

        putexcel `col_raw'`row' = `share_raw', bold
        putexcel `col_cor'`row' = `share_cor', bold

        local col = `col' + 2
    }

    putexcel A`row', bold border(top, medium)
	
	foreach col in A B C D E F G  {
    putexcel `col'1, overwrite border(bottom, thin, black)
	}

* --- Thick border after All Countries ---
	foreach col in A B C D E F G {
    putexcel `col'5, overwrite border(bottom, thin, black)
	}

}
}









***********MODEL 2 No inflation

levelsof Category, local(categ)

foreach c of local categ {
forvalues f = 1/2 {

    * Define training and validation folds
    local val_fold = `f'
    local train_fold = 3 - `f'   // if f=1, train=2; if f=2, train=1
    * --- Estimate model on training fold ---
    reg pli_bench inc_UM inc_LM inc_L ///
                  pli_ext_H pli_ext_UM pli_ext_LM pli_ext_L ///
                  year2021 ///
                  if fold == `train_fold' & Category=="`c'", robust

    * --- Apply correction to validation fold ---
    replace pli_corrected_cv = _b[_cons] + _b[pli_ext_H]  * pli_ext  if fold == `val_fold' & IncomeClassification == "H" & Category=="`c'"
    replace pli_corrected_cv = _b[_cons] + _b[inc_UM]  + _b[pli_ext_UM] * pli_ext  if fold == `val_fold' & IncomeClassification == "UM" & Category=="`c'"
    replace pli_corrected_cv = _b[_cons] + _b[inc_LM]  + _b[pli_ext_LM] * pli_ext  if fold == `val_fold' & IncomeClassification == "LM" & Category=="`c'"
    replace pli_corrected_cv = _b[_cons] + _b[inc_L]   + _b[pli_ext_L]  * pli_ext  if fold == `val_fold' & IncomeClassification == "L" & Category=="`c'"
	
	quietly summarize pli_corrected if Country_Code == "USA"
    local usa_predicted = r(mean)

* Rescale so USA = 1
  replace pli_corrected_norm = pli_corrected / `usa_predicted'
}
}

* ============================================================
* STEP 3 — Compare corrected vs raw pli_ext against benchmark
* ============================================================

* Generate error terms
replace error_corrected = pli_corrected_norm - pli_bench

replace sq_error_corrected = error_corrected^2

replace abs_error_corrected = abs(error_corrected)

levelsof Category, local(categ) 
foreach c of local categ {
	
* RMSE
quietly summarize sq_error_raw if Category=="`c'"
local rmse_raw = sqrt(r(mean)) 

quietly summarize sq_error_corrected if Category=="`c'"
local rmse_corrected = sqrt(r(mean)) 

* MAE


quietly summarize abs_error_raw if Category=="`c'"
local mae_raw = r(mean)

quietly summarize abs_error_corrected if Category=="`c'"
local mae_corrected = r(mean)

* Display results
di         " `c'"
di "         Raw pli_ext    Corrected"
di "RMSE:    `rmse_raw'     `rmse_corrected'"
di "MAE:     `mae_raw'      `mae_corrected'"
di "========================================"

}


levelsof Category, local(categ) 
foreach c of local categ {
foreach inc in "H" "UM" "LM" "L" {
    quietly summarize sq_error_raw       if IncomeClassification == "`inc'" & Category=="`c'"
    local rmse_raw_g = sqrt(r(mean))
    quietly summarize sq_error_corrected if IncomeClassification == "`inc'" & Category=="`c'"
    local rmse_cor_g = sqrt(r(mean))
    di "Income group `inc' — `c' RMSE raw: `rmse_raw_g'  corrected: `rmse_cor_g'"
	
	quietly summarize abs_error_raw       if IncomeClassification == "`inc'" & Category=="`c'"
    local mae_raw_g = sqrt(r(mean))
    quietly summarize abs_error_corrected if IncomeClassification == "`inc'" & Category=="`c'"
    local mae_corrected_g = sqrt(r(mean))
    di "Income group `inc' — `c' MAE raw: `mae_raw_g'  corrected: `mae_corrected_g'"
}
}


*****************MODEL 3 GDP as continuous variable*******************************

reg pli_bench inc_UM inc_LM inc_L ///
                  pli_ext_H pli_ext_UM pli_ext_LM pli_ext_L ///
                  year2021 ///
				  if Category=="PA.NUS.PPP"


reg pli_bench pli_ext GDPpcPPP inflation ///
                  year2021 ///
				  if Category=="PA.NUS.PPP"

gen ln_gdppc=ln(GDPpcPPP)
gen pli_ext_gdppc = pli_ext * ln_gdppc


levelsof Category, local(categ)
foreach c of local categ {
forvalues f = 1/2 {

 * Define training and validation folds
    local val_fold = `f'
    local train_fold = 3 - `f'   // if f=1, train=2; if f=2, train=1
      
    reg pli_bench pli_ext pli_ext_gdppc ln_gdppc inflation year2021  if fold == `train_fold' & Category=="`c'", robust

    * --- Apply correction to validation fold ---
    replace pli_corrected_cv = _b[_cons] + _b[pli_ext]  * pli_ext +  _b[pli_ext_gdppc] * pli_ext_gdppc + _b[ln_gdppc]*ln_gdppc+ _b[inflation]*inflation  if    fold == `val_fold' 
	
	quietly summarize pli_corrected if Country_Code == "USA"
    local usa_predicted = r(mean)

* Rescale so USA = 1
  replace pli_corrected_norm = pli_corrected / `usa_predicted'
}
}
* ============================================================
* STEP 3 — Compare corrected vs raw pli_ext against benchmark
* ============================================================

* Generate error terms
replace error_corrected = pli_corrected_norm - pli_bench

replace sq_error_corrected = error_corrected^2

replace abs_error_corrected = abs(error_corrected)

gen pct_error_raw       = abs(pli_ext  - pli_bench) / pli_bench * 100
gen pct_error_corrected = abs(pli_corrected_cv - pli_bench) / pli_bench * 100

levelsof Category, local(categ) 
foreach c of local categ {
	
* RMSE
quietly summarize sq_error_raw if Category=="`c'"
local rmse_raw = sqrt(r(mean)) 

quietly summarize sq_error_corrected if Category=="`c'"
local rmse_corrected = sqrt(r(mean)) 

* MAE


quietly summarize abs_error_raw if Category=="`c'"
local mae_raw = r(mean)

quietly summarize abs_error_corrected if Category=="`c'"
local mae_corrected = r(mean)

* MAPE

quietly summarize pct_error_raw if Category=="`c'"
local mape_raw = r(mean)

quietly summarize pct_error_corrected if Category=="`c'"
local mape_corrected = r(mean)

* Display results
di         " `c'"
di "         Raw pli_ext    Corrected"
di "RMSE:    `rmse_raw'     `rmse_corrected'"
di "MAPE:     `mape_raw'      `mape_corrected'"
di "========================================"

}


levelsof Category, local(categ) 
foreach c of local categ {
foreach inc in "H" "UM" "LM" "L" {
    quietly summarize sq_error_raw       if IncomeClassification == "`inc'" & Category=="`c'"
    local rmse_raw_g = sqrt(r(mean))
    quietly summarize sq_error_corrected if IncomeClassification == "`inc'" & Category=="`c'"
    local rmse_cor_g = sqrt(r(mean))
    di "Income group `inc' — `c' RMSE raw: `rmse_raw_g'  corrected: `rmse_cor_g'"
	
	quietly summarize pct_error_raw       if IncomeClassification == "`inc'" & Category=="`c'"
    local mae_raw_g = sqrt(r(mean))
    quietly summarize pct_error_corrected if IncomeClassification == "`inc'" & Category=="`c'"
    local mape_corrected_g = sqrt(r(mean))
    di "Income group `inc' — `c' MAE raw: `mae_raw_g'  corrected: `mae_corrected_g'"
}
}











//High income as the reference group
reg pli_bench inc_UM inc_LM inc_L ///
              pli_ext_H pli_ext_UM pli_ext_LM pli_ext_L ///
              year2021 if Category=="PA.NUS.PPP", robust
			  
			  
scalar alpha_H  = _b[_cons]
scalar alpha_UM = _b[_cons] + _b[inc_UM]
scalar alpha_LM = _b[_cons] + _b[inc_LM]
scalar alpha_L  = _b[_cons] + _b[inc_L]

scalar beta_H  = _b[pli_ext_H]
scalar beta_UM = _b[pli_ext_UM]
scalar beta_LM = _b[pli_ext_LM]
scalar beta_L  = _b[pli_ext_L]


gen pli_corrected = .
replace pli_corrected = alpha_H  + beta_H  * pli_ext if IncomeClassification == "H" & Time==2021 & Category=="PA.NUS.PPP"
replace pli_corrected = alpha_UM + beta_UM * pli_ext if IncomeClassification == "UM" & Time==2021 & Category=="PA.NUS.PPP"
replace pli_corrected = alpha_LM + beta_LM * pli_ext if IncomeClassification == "LM" & Time==2021 & Category=="PA.NUS.PPP"
replace pli_corrected = alpha_L  + beta_L  * pli_ext if IncomeClassification == "L" & Time==2021 & Category=="PA.NUS.PPP"


 quietly summarize pli_corrected if Time==2021 & Category=="PA.NUS.PPP", detail
    local xmin = round(r(min),   0.01)
    local xq1  = round(r(max)/4,   0.01)
    local xmed = round(r(max)/2,   0.01)
    local xq3  = round(r(max)*2/3,   0.01)
    local xmax = round(r(max) + 0.1, 0.01)
	local xmin_plot = `xmin' - 0.01

twoway ///
        (scatter pli_bench pli_corrected if Time==2021 & Category=="PA.NUS.PPP", mlabel(Country_Code)) ///
        (lfit    pli_bench pli_corrected if Time==2021 & Category=="PA.NUS.PPP") ///
        (function y = x, range(`xmin' `xmax') lpattern(dash) lcolor(gray)), ///
        xlabel(`xmin' `xq1' `xmed' `xq3' `xmax', format(%4.2f)) ///
        ylabel(`xmin' `xq1' `xmed' `xq3' `xmax', format(%4.2f)) ///
        xscale(range(`xmin_plot' `xmax')) yscale(range(`xmin_plot' `xmax')) ///
		plotregion(margin(zero)) ///
        graphregion(margin(t=8 b=2 l=2 r=8)) ///
		xtitle("Extrapolation corrected")  ///
		ytitle("Benchmark")  ///
		xsize(7) ysize(5) ///
        title(" 2021 PLI Benchmark vs Extrapolation corrected ",  size(medsmall)) ///
        legend(order(1 "Observations" 2 "Fitted line" 3 "45° line") pos(11) ring(0))

    graph export "charts\_2021\PLIreg2021_benchmarkvscorrected.png", replace
	
 quietly summarize pli_ext if Time==2021 & Category=="PA.NUS.PPP", detail
    local xmin = round(r(min),   0.01)
    local xq1  = round(r(max)/4,   0.01)
    local xmed = round(r(max)/2,   0.01)
    local xq3  = round(r(max)*2/3,   0.01)
    local xmax = round(r(max) + 0.1, 0.01)
	local xmin_plot = `xmin' - 0.01

twoway ///
        (scatter pli_bench pli_ext if Time==2021 & Category=="PA.NUS.PPP", mlabel(Country_Code)) ///
        (lfit    pli_bench pli_ext if Time==2021 & Category=="PA.NUS.PPP") ///
        (function y = x, range(`xmin' `xmax') lpattern(dash) lcolor(gray)), ///
        xlabel(`xmin' `xq1' `xmed' `xq3' `xmax', format(%4.2f)) ///
        ylabel(`xmin' `xq1' `xmed' `xq3' `xmax', format(%4.2f)) ///
        xscale(range(`xmin_plot' `xmax')) yscale(range(`xmin_plot' `xmax')) ///
		plotregion(margin(zero)) ///
        graphregion(margin(t=8 b=2 l=2 r=8)) ///
		xtitle("Extrapolation")  ///
		ytitle("Benchmark")  ///
		xsize(7) ysize(5) ///
        title(" 2021 PLI Benchmark vs Extrapolation ",  size(medsmall)) ///
        legend(order(1 "Observations" 2 "Fitted line" 3 "45° line") pos(11) ring(0))

    graph export "charts\_2021\PLIreg2021_benchmarkvsextrapolated.png", replace

	
gen pct_error_pli    = (pli_ext - pli_bench) / pli_bench * 100 
gen abs_pct_error_pli = abs(pct_error_pli)

gen sqr_error_pli = (pli_ext - pli_bench)^2
gen abs_error_pli = abs(pli_ext - pli_bench)

sum pct_error_pli
sum abs_pct_error_pli if 
sum abs_error_pli


gen pct_error_pli_cor    = (pli_corrected - pli_bench) / pli_bench * 100 
gen abs_pct_error_pli_cor = abs(pct_error_pli_cor)

gen sqr_error_pli_cor = (pli_corrected - pli_bench)^2
gen abs_error_pli = abs(pli_ext - pli_bench)