clear all
set more off

cd "C:\Users\wb589699\OneDrive - WBG\Desktop\WDIvsBenchmark_analysis"


**********************************************************************************************************************************************************************************
***********************************************************************************2021*******************************************************************************************
***********************************************************************************************************************************************************************************


use ICPvsWDI_2021, clear


* ============================================================================
* 7. GENERATE EXCEL TABLES WITH REGRESSION ESTIMATES
* ============================================================================
* Treat Georgia and Ukraine as Eurostat-OECD region for subsequent analysis
replace ICPregion = "d_EUO" if ICPregion == "h_SPP"
replace ICPregion="e_LAC" if inlist(ICPregion, "e_LAT", "f_CAR")

*Putexcel option
putexcel set "mincer_zarnowitz_results.xlsx", replace
levelsof Category, local(series)

levelsof IncomeClassification, local(inc_levels)
levelsof ICPregion, local(reg_levels)

gen pct_error_pli    = (pli_ext - pli_bench) / pli_bench * 100
gen abs_pct_error_pli = abs(pct_error_pli)

gen sqr_error_pli = (pli_ext - pli_bench)^2
gen abs_error_pli = abs(pli_ext - pli_bench)

* Write headers
foreach ser in `series' {
putexcel set "mincer_zarnowitz_results.xlsx",  sheet(2021_`ser') modify
putexcel A1="Group" B1="Coefficient" C1="Intercept" D1="R2" E1="Adj. R2" F1="RMSE" G1="MPE" H1="MAPE" I1="MAE" J1="N.Obs"

local r = 2

* Define a helper: rounds and appends stars based on p-value
* Stars: *** p<0.01, ** p<0.05, * p<0.10

    reg pli_bench pli_ext if Category=="`ser'"
    local b_coef  = string(round(_b[pli_ext], 0.0001), "%9.4f")
    local b_cons  = string(round(_b[_cons], 0.0001), "%9.4f")
    local p_coef  = 2*ttail(e(df_r), abs(_b[pli_ext]/_se[pli_ext]))
    local p_cons  = 2*ttail(e(df_r), abs(_b[_cons]/_se[_cons]))
    
    * Append stars
    if `p_coef' < 0.01  local b_coef "`b_coef'***"
    else if `p_coef' < 0.05 local b_coef "`b_coef'**"
    else if `p_coef' < 0.10 local b_coef "`b_coef'*"
    
    if `p_cons' < 0.01  local b_cons "`b_cons'***"
    else if `p_cons' < 0.05 local b_cons "`b_cons'**"
    else if `p_cons' < 0.10 local b_cons "`b_cons'*"
	
	summarize pct_error_pli if Category=="`ser'"
    scalar MPE  = r(mean)

    summarize abs_pct_error_pli if Category=="`ser'"
    scalar MAPE = r(mean)
    
	summarize abs_error_pli if Category=="`ser'"
    scalar MAE = r(mean)

    putexcel A`r' = "All Countries" B`r' = "`b_coef'" C`r' = "`b_cons'" ///
             D`r' = `e(r2)' E`r' = `e(r2_a)' F`r' = `e(rmse)' G`r'=MPE H`r'=MAPE I`r'=MAE J`r' = `e(N)'

    local ++r



* Define a helper: rounds and appends stars based on p-value
* Stars: *** p<0.01, ** p<0.05, * p<0.10

foreach inc in `inc_levels' {
    reg pli_bench pli_ext if IncomeClassification == "`inc'" & Category=="`ser'"
    local b_coef  = string(round(_b[pli_ext], 0.0001), "%9.4f")
    local b_cons  = string(round(_b[_cons], 0.0001), "%9.4f")
    local p_coef  = 2*ttail(e(df_r), abs(_b[pli_ext]/_se[pli_ext]))
    local p_cons  = 2*ttail(e(df_r), abs(_b[_cons]/_se[_cons]))
    
    * Append stars
    if `p_coef' < 0.01  local b_coef "`b_coef'***"
    else if `p_coef' < 0.05 local b_coef "`b_coef'**"
    else if `p_coef' < 0.10 local b_coef "`b_coef'*"
    
    if `p_cons' < 0.01  local b_cons "`b_cons'***"
    else if `p_cons' < 0.05 local b_cons "`b_cons'**"
    else if `p_cons' < 0.10 local b_cons "`b_cons'*"
	
	summarize pct_error_pli if IncomeClassification == "`inc'" & Category=="`ser'"
    scalar MPE  = r(mean)

    summarize abs_pct_error_pli if IncomeClassification == "`inc'" & Category=="`ser'"
    scalar MAPE = r(mean)
    
	summarize abs_error_pli if IncomeClassification == "`inc'" & Category=="`ser'"
    scalar MAE = r(mean)

    putexcel A`r' = "`inc'" B`r' = "`b_coef'" C`r' = "`b_cons'" ///
             D`r' = `e(r2)' E`r' = `e(r2_a)' F`r' = `e(rmse)' G`r'=MPE H`r'=MAPE I`r'=MAE J`r' = `e(N)'
    local ++r
}

foreach reg in `reg_levels' {
    reg pli_bench pli_ext if ICPregion == "`reg'"  & Category=="`ser'"
    local b_coef  = string(round(_b[pli_ext], 0.0001), "%9.4f")
    local b_cons  = string(round(_b[_cons], 0.0001), "%9.4f")
    local p_coef  = 2*ttail(e(df_r), abs(_b[pli_ext]/_se[pli_ext]))
    local p_cons  = 2*ttail(e(df_r), abs(_b[_cons]/_se[_cons]))
    
    * Append stars
    if `p_coef' < 0.01  local b_coef "`b_coef'***" 
    else if `p_coef' < 0.05 local b_coef "`b_coef'**"
    else if `p_coef' < 0.10 local b_coef "`b_coef'*"
    
    if `p_cons' < 0.01  local b_cons "`b_cons'***"
    else if `p_cons' < 0.05 local b_cons "`b_cons'**"
    else if `p_cons' < 0.10 local b_cons "`b_cons'*"

    summarize pct_error_pli  if ICPregion == "`reg'"  & Category=="`ser'"
    scalar MPE  = r(mean)

    summarize abs_pct_error_pli  if ICPregion == "`reg'"  & Category=="`ser'"
    scalar MAPE = r(mean)
    
	summarize abs_error_pli  if ICPregion == "`reg'"  & Category=="`ser'"
    scalar MAE = r(mean)

    putexcel A`r' = "`reg'" B`r' = "`b_coef'" C`r' = "`b_cons'" ///
             D`r' = `e(r2)' E`r' = `e(r2_a)' F`r' = `e(rmse)' G`r'=MPE H`r'=MAPE I`r'=MAE J`r' = `e(N)'
    local ++r
}

foreach col in A B C D E F G H I J {
    putexcel `col'1, overwrite border(bottom, double, black)
}

* --- Thick border after All Countries ---
foreach col in A B C D E F G H I J {
    putexcel `col'2, overwrite border(bottom, thin, black)
}

* --- Thick border after income block (Low Income = row 6) ---
foreach col in A B C D E F G H I J {
    putexcel `col'6, overwrite border(bottom, thin, black)
}

foreach col in A B C D E F G H I J {
    putexcel `col'12, overwrite border(bottom, double, black)
}

}


//Plotting the data
***
//Transforming some variables in log
gen log_inflation=ln(Inflation_2017)
gen log_abs_pct_error=ln(abs_pct_error)
gen log_gdp_pc=ln(GDPpcPPP)



centile Inflation_2017, centile(95)
local cutoff = r(c_1)
* Step 2: Regression excluding top 5%
reg pct_error Inflation_2017 if Inflation_2017 <= `cutoff'

* Step 3: Graph excluding top 5%
twoway (scatter pct_error Inflation_2017 if Inflation_2017 <= `cutoff' & Category == "PA.NUS.PPP" , mlabel(Country_Code)) ///
       (lfit pct_error Inflation_2017 if Inflation_2017 <= `cutoff' & Category == "PA.NUS.PPP") , ///
    ytitle("Percentage Error (%)") ///
    xtitle("Inflation Rate 2017-2021") ///
    title("Percentage Errors vs. Inflation Rates (excl. top 5% outliers)") 
graph export "charts\_2021\PCTerrorVSInflation.png", replace	
	
twoway (scatter log_abs_pct_error log_inflation if Category == "PA.NUS.PPP", mlabel(Country_Code)) ///
(lfit log_abs_pct_error log_inflation if Category == "PA.NUS.PPP"), ///
ytitle("Percentage Error (log))") ///
    xtitle("Inflation Rate 2017 (log)") ///
    title("Percentage Errors vs. Inflation Rates (log)") 
graph export "charts\_2021PCTerrorVSInflation_log.png", replace

	   
twoway (scatter pct_error log_gdp_pc if Category == "PA.NUS.PPP" , mlabel(Country_Code)) ///
       (lfit abs_pct_error log_gdp_pc if Category == "PA.NUS.PPP" ), ///
	   ytitle("Percentage Error (%)") ///
    xtitle("GDP pc (PPP log)") ///
    title("Percentage Errors vs. GDP pc") 
graph export "charts\_2021\PCTerrorVSGDP.png", replace	


* ============================================================
* LOOP OVER ICPregion
* ============================================================
gen series="HHC" if Category == "PA.NUS.PRVT.PP"
replace series="GDP" if missing(series)

levelsof series, local(ser)

foreach s of local ser {
	
levelsof ICPregion, local(regions)

foreach reg of local regions {

    * --- Compute axis range from pli_extrap distribution ---
    quietly summarize pli_ext if ICPregion == "`reg'" & series == "`s'", detail
    local xmin = round(r(min), 0.01)
    local xmax = round(r(max) + 0.05, 0.01)
    local xmin_plot = `xmin' - 0.05

    local range = `xmax' - `xmin'
    local xq1  = round(`xmin' + `range' * 1/4, 0.01)
    local xmed = round(`xmin' + `range' * 2/4, 0.01)
    local xq3  = round(`xmin' + `range' * 3/4, 0.01)

    * --- Clean label (remove "c_" prefix if present) ---
    local reg_label = substr("`reg'", 3, .)

    twoway ///
        (scatter pli_bench pli_ext if ICPregion == "`reg'" & series == "`s'", mlabel(Country_Code)) ///
        (lfit    pli_bench pli_ext if ICPregion == "`reg'" & series == "`s'") ///
        (function y = x, range(`xmin_plot' `xmax') lpattern(dash) lcolor(gray)), ///
        xlabel(`xmin' `xq1' `xmed' `xq3' `xmax', format(%4.2f)) ///
        ylabel(`xmin' `xq1' `xmed' `xq3' `xmax', format(%4.2f)) ///
        xscale(range(`xmin_plot' `xmax')) yscale(range(`xmin_plot' `xmax')) ///
		plotregion(margin(zero)) ///
        graphregion(margin(t=8 b=2 l=2 r=8)) ///
		xtitle("Extrapolation")  ///
		ytitle("Benchmark")  ///
		xsize(7) ysize(5) ///
        title("`s' 2021 PLI Benchmark vs Extrapolated — `reg_label'",  size(medsmall)) ///
        legend(order(1 "Observations" 2 "Fitted line" 3 "45° line") pos(11) ring(0))

    graph export "charts\_2021\PLIreg2021_`s'_`reg_label'.png", replace
}
}

* ============================================================
* LOOP OVER IncomeClassification
* ============================================================
gen IncomeClass_label = cond(IncomeClassification == "H",  "High Income", ///
                        cond(IncomeClassification == "UM", "Upper Middle Income", ///
                        cond(IncomeClassification == "LM", "Lower Middle Income", ///
                        cond(IncomeClassification == "L",  "Low Income", ""))))
levelsof series, local(ser)

foreach s of local ser {
	
levelsof IncomeClass_label, local(incomes)

foreach inc of local incomes {

 * --- Compute axis range from pli_extrap distribution ---
    quietly summarize pli_ext if IncomeClass_label == "`inc'" & series == "`s'", detail
    local xmin = round(r(min), 0.01)
    local xmax = round(r(max) + 0.05, 0.01)
    local xmin_plot = `xmin' - 0.05

    local range = `xmax' - `xmin'
    local xq1  = round(`xmin' + `range' * 1/4, 0.01)
    local xmed = round(`xmin' + `range' * 2/4, 0.01)
    local xq3  = round(`xmin' + `range' * 3/4, 0.01)

    * --- Clean label (remove "c_" prefix if present) ---
    

    twoway ///
        (scatter pli_bench pli_ext if IncomeClass_label == "`inc'" & series == "`s'", mlabel(Country_Code)) ///
        (lfit    pli_bench pli_ext if IncomeClass_label == "`inc'" & series == "`s'") ///
        (function y = x, range(`xmin_plot' `xmax') lpattern(dash) lcolor(gray)), ///
        xlabel(`xmin' `xq1' `xmed' `xq3' `xmax', format(%4.2f)) ///
        ylabel(`xmin' `xq1' `xmed' `xq3' `xmax', format(%4.2f)) ///
        xscale(range(`xmin_plot' `xmax')) yscale(range(`xmin_plot' `xmax')) ///
		plotregion(margin(zero)) ///
        graphregion(margin(t=8 b=2 l=2 r=8)) ///
		xtitle("Extrapolation")  ///
		ytitle("Benchmark")  ///
		xsize(7) ysize(5) ///
        title("`s' 2021 PLI Benchmark vs Extrapolated — `inc'", size(medsmall)) ///
        legend(order(1 "Observations" 2 "Fitted line" 3 "45° line") pos(11) ring(0))

    graph export "charts\_2021\PLIreg2021_`s'_`inc'.png", replace
}
}
******************************************************************************************************************************
*************************************************2017*************************************************************************
******************************************************************************************************************************

use ICPvsWDI_2017, clear

//Mincer-Zarnowitz Regression

replace ICPregion="d_EUO" if ICPregion=="h_SPP" 
replace ICPregion="g_WAS" if Country_Code=="IRN" 
replace ICPregion="e_LAC" if inlist(ICPregion, "e_LAT", "f_CAR")

gen pct_error_pli    = (pli_ext - pli_bench) / pli_bench * 100
gen abs_pct_error_pli = abs(pct_error_pli)

gen sqr_error_pli = (pli_ext - pli_bench)^2
gen abs_error_pli = abs(pli_ext - pli_bench)

//Putexcel option
levelsof Category, local(series)
levelsof IncomeClassification, local(inc_levels)
levelsof ICPregion, local(reg_levels)
* Write headers
foreach ser in `series' {
putexcel set "mincer_zarnowitz_results.xlsx",  sheet(2017_`ser') modify
putexcel A1="Group" B1="Coefficient" C1="Intercept" D1="R2" E1="Adj. R2" F1="RMSE" G1="MPE" H1="MAPE" I1="MAE" J1="N.Obs"

local r = 2

* Define a helper: rounds and appends stars based on p-value
* Stars: *** p<0.01, ** p<0.05, * p<0.10

    reg pli_bench pli_ext if Category=="`ser'"
    local b_coef  = string(round(_b[pli_ext], 0.0001), "%9.4f")
    local b_cons  = string(round(_b[_cons], 0.0001), "%9.4f")
    local p_coef  = 2*ttail(e(df_r), abs(_b[pli_ext]/_se[pli_ext]))
    local p_cons  = 2*ttail(e(df_r), abs(_b[_cons]/_se[_cons]))
    
    * Append stars
    if `p_coef' < 0.01  local b_coef "`b_coef'***"
    else if `p_coef' < 0.05 local b_coef "`b_coef'**"
    else if `p_coef' < 0.10 local b_coef "`b_coef'*"
    
    if `p_cons' < 0.01  local b_cons "`b_cons'***"
    else if `p_cons' < 0.05 local b_cons "`b_cons'**"
    else if `p_cons' < 0.10 local b_cons "`b_cons'*"
	
	summarize pct_error_pli if Category=="`ser'"
    scalar MPE  = r(mean)

    summarize abs_pct_error_pli if Category=="`ser'"
    scalar MAPE = r(mean)
    
	summarize abs_error_pli if Category=="`ser'"
    scalar MAE = r(mean)

    putexcel A`r' = "All Countries" B`r' = "`b_coef'" C`r' = "`b_cons'" ///
             D`r' = `e(r2)' E`r' = `e(r2_a)' F`r' = `e(rmse)' G`r'=MPE H`r'=MAPE I`r'=MAE J`r' = `e(N)'

    local ++r



* Define a helper: rounds and appends stars based on p-value
* Stars: *** p<0.01, ** p<0.05, * p<0.10

foreach inc in `inc_levels' {
    reg pli_bench pli_ext if IncomeClassification == "`inc'" & Category=="`ser'"
    local b_coef  = string(round(_b[pli_ext], 0.0001), "%9.4f")
    local b_cons  = string(round(_b[_cons], 0.0001), "%9.4f")
    local p_coef  = 2*ttail(e(df_r), abs(_b[pli_ext]/_se[pli_ext]))
    local p_cons  = 2*ttail(e(df_r), abs(_b[_cons]/_se[_cons]))
    
    * Append stars
    if `p_coef' < 0.01  local b_coef "`b_coef'***"
    else if `p_coef' < 0.05 local b_coef "`b_coef'**"
    else if `p_coef' < 0.10 local b_coef "`b_coef'*"
    
    if `p_cons' < 0.01  local b_cons "`b_cons'***"
    else if `p_cons' < 0.05 local b_cons "`b_cons'**"
    else if `p_cons' < 0.10 local b_cons "`b_cons'*"
	
	summarize pct_error_pli if IncomeClassification == "`inc'" & Category=="`ser'"
    scalar MPE  = r(mean)

    summarize abs_pct_error_pli if IncomeClassification == "`inc'" & Category=="`ser'"
    scalar MAPE = r(mean)
    
	summarize abs_error_pli if IncomeClassification == "`inc'" & Category=="`ser'"
    scalar MAE = r(mean)

    putexcel A`r' = "`inc'" B`r' = "`b_coef'" C`r' = "`b_cons'" ///
             D`r' = `e(r2)' E`r' = `e(r2_a)' F`r' = `e(rmse)' G`r'=MPE H`r'=MAPE I`r'=MAE J`r' = `e(N)'
    local ++r
}

foreach reg in `reg_levels' {
    reg pli_bench pli_ext if ICPregion == "`reg'"  & Category=="`ser'"
    local b_coef  = string(round(_b[pli_ext], 0.0001), "%9.4f")
    local b_cons  = string(round(_b[_cons], 0.0001), "%9.4f")
    local p_coef  = 2*ttail(e(df_r), abs(_b[pli_ext]/_se[pli_ext]))
    local p_cons  = 2*ttail(e(df_r), abs(_b[_cons]/_se[_cons]))
    
    * Append stars
    if `p_coef' < 0.01  local b_coef "`b_coef'***" 
    else if `p_coef' < 0.05 local b_coef "`b_coef'**"
    else if `p_coef' < 0.10 local b_coef "`b_coef'*"
    
    if `p_cons' < 0.01  local b_cons "`b_cons'***"
    else if `p_cons' < 0.05 local b_cons "`b_cons'**"
    else if `p_cons' < 0.10 local b_cons "`b_cons'*"

    summarize pct_error_pli  if ICPregion == "`reg'"  & Category=="`ser'"
    scalar MPE  = r(mean)

    summarize abs_pct_error_pli  if ICPregion == "`reg'"  & Category=="`ser'"
    scalar MAPE = r(mean)
    
	summarize abs_error_pli  if ICPregion == "`reg'"  & Category=="`ser'"
    scalar MAE = r(mean)

    putexcel A`r' = "`reg'" B`r' = "`b_coef'" C`r' = "`b_cons'" ///
             D`r' = `e(r2)' E`r' = `e(r2_a)' F`r' = `e(rmse)' G`r'=MPE H`r'=MAPE I`r'=MAE J`r' = `e(N)'
    local ++r
}

foreach col in A B C D E F G H I J {
    putexcel `col'1, overwrite border(bottom, double, black)
}

* --- Thick border after All Countries ---
foreach col in A B C D E F G H I J {
    putexcel `col'2, overwrite border(bottom, thin, black)
}

* --- Thick border after income block (Low Income = row 6) ---
foreach col in A B C D E F G H I J {
    putexcel `col'6, overwrite border(bottom, thin, black)
}

foreach col in A B C D E F G H I J {
    putexcel `col'12, overwrite border(bottom, double, black)
}

}

* ============================================================
* LOOP OVER ICPregion
* ============================================================
gen series="HHC" if Category == "PA.NUS.PRVT.PP"
replace series="GDP" if missing(series)

levelsof series, local(ser)

foreach s of local ser {
	
levelsof ICPregion, local(regions)

foreach reg of local regions {

    * --- Compute axis range from pli_extrap distribution ---
    quietly summarize pli_ext if ICPregion == "`reg'" & series == "`s'", detail
    local xmin = round(r(min), 0.01)
    local xmax = round(r(max) + 0.05, 0.01)
    local xmin_plot = `xmin' - 0.05

    local range = `xmax' - `xmin'
    local xq1  = round(`xmin' + `range' * 1/4, 0.01)
    local xmed = round(`xmin' + `range' * 2/4, 0.01)
    local xq3  = round(`xmin' + `range' * 3/4, 0.01)
    * --- Clean label (remove "c_" prefix if present) ---
    local reg_label = substr("`reg'", 3, .)

    twoway ///
        (scatter pli_bench pli_ext if ICPregion == "`reg'" & series == "`s'", mlabel(Country_Code)) ///
        (lfit    pli_bench pli_ext if ICPregion == "`reg'" & series == "`s'") ///
        (function y = x, range(`xmin_plot' `xmax') lpattern(dash) lcolor(gray)), ///
        xlabel(`xmin' `xq1' `xmed' `xq3' `xmax', format(%4.2f)) ///
        ylabel(`xmin' `xq1' `xmed' `xq3' `xmax', format(%4.2f)) ///
        xscale(range(`xmin_plot' `xmax')) yscale(range(`xmin_plot' `xmax')) ///
		plotregion(margin(zero)) ///
        graphregion(margin(t=8 b=2 l=2 r=8)) ///
		xtitle("Extrapolation")  ///
		ytitle("Benchmark")  ///
		xsize(7) ysize(5) ///
        title("`s' 2017 PLI Benchmark vs Extrapolated — `reg_label'",  size(medsmall)) ///
        legend(order(1 "Observations" 2 "Fitted line" 3 "45° line") pos(11) ring(0))

    graph export "charts\_2017\PLIreg2017_`s'_`reg_label'.png", replace
}
}

* ============================================================
* LOOP OVER IncomeClassification
* ============================================================
gen IncomeClass_label = cond(IncomeClassification == "H",  "High Income", ///
                        cond(IncomeClassification == "UM", "Upper Middle Income", ///
                        cond(IncomeClassification == "LM", "Lower Middle Income", ///
                        cond(IncomeClassification == "L",  "Low Income", ""))))
levelsof series, local(ser)

foreach s of local ser {
	
levelsof IncomeClass_label, local(incomes)

foreach inc of local incomes {

 * --- Compute axis range from pli_extrap distribution ---
    quietly summarize pli_ext if IncomeClass_label == "`inc'" & series == "`s'", detail
    local xmin = round(r(min), 0.01)
    local xmax = round(r(max) + 0.05, 0.01)
    local xmin_plot = `xmin' - 0.05

    local range = `xmax' - `xmin'
    local xq1  = round(`xmin' + `range' * 1/4, 0.01)
    local xmed = round(`xmin' + `range' * 2/4, 0.01)
    local xq3  = round(`xmin' + `range' * 3/4, 0.01)

    * --- Clean label (remove "c_" prefix if present) ---
    

    twoway ///
        (scatter pli_bench pli_ext if IncomeClass_label == "`inc'" & series == "`s'", mlabel(Country_Code)) ///
        (lfit    pli_bench pli_ext if IncomeClass_label == "`inc'" & series == "`s'") ///
        (function y = x, range(`xmin_plot' `xmax') lpattern(dash) lcolor(gray)), ///
        xlabel(`xmin' `xq1' `xmed' `xq3' `xmax', format(%4.2f)) ///
        ylabel(`xmin' `xq1' `xmed' `xq3' `xmax', format(%4.2f)) ///
        xscale(range(`xmin_plot' `xmax')) yscale(range(`xmin_plot' `xmax')) ///
		plotregion(margin(zero)) ///
        graphregion(margin(t=8 b=2 l=2 r=8)) ///
		xtitle("Extrapolation")  ///
		ytitle("Benchmark")  ///
		xsize(7) ysize(5) ///
        title("`s' 2017 PLI Benchmark vs Extrapolated — `inc'", size(medsmall)) ///
        legend(order(1 "Observations" 2 "Fitted line" 3 "45° line") pos(11) ring(0))

    graph export "charts\_2017\PLIreg2027_`s'_`inc'.png", replace
}
}


