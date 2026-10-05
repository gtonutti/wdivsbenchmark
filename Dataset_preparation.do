clear all
set more off

cd "C:\Users\wb589699\OneDrive - WBG\Desktop\WDIvsBenchmark_analysis"

**********************************************************************************************************************************************************************************
**********************************************************************AUXILIARY DATA*********************************************************************
***********************************************************************************************************************************************************************************


* ============================================================================
* 1. EXCHANGE RATES
*    Source: Economy SNA database (Atlas and official exchange rates)
* ============================================================================

odbc load, exec("SELECT * FROM Economy.dbo.FACT AS Fact WHERE Fact.Indicator_Code IN ('PA.NUS.ATLS', 'PA.NUS.FCRF')") dsn("DCS_prod") clear

replace Time_Code = strtrim(Time_Code)

* For Zimbabwe, replace Atlas rate with the official (FCRF) rate
bysort Country_Code Time_Code: gen    xr_dec     = Fact_Value * (Indicator_Code == "PA.NUS.FCRF")
bysort Country_Code Time_Code: egen   max_xr_dec = max(xr_dec)
replace Fact_Value = max_xr_dec if Country_Code == "ZWE"

keep if inlist(Time_Code, "YR2017", "YR2021") & Indicator_Code == "PA.NUS.ATLS"
keep  Time_Code Country_Code Fact_Value
rename Fact_Value xr

save XR, replace

import excel "MER_WDI_Database_Archives.xlsx", clear firstrow
drop SeriesName
drop if missing(SeriesCode)

replace SeriesCode = "XR_official" if SeriesCode == "PA.NUS.FCRF"
replace SeriesCode = "XR_dec"  if SeriesCode == "PA.NUS.ATLS"

reshape long  YR, i(CountryName CountryCode  VersionName VersionCode SeriesCode) j(Time) string
replace YR="" if YR==".."
drop if (VersionCode=="202004" & Time=="2021") |  (VersionCode=="202403" & Time=="2017")
reshape wide  YR, i(CountryName CountryCode  VersionName VersionCode) j(SeriesCode) string

ren (CountryName CountryCode YR*) (Country_Name Country_Code *)
destring, replace
tostring Time, gen(Time_Code)
replace Time_Code="YR"+Time_Code
drop if missing(XR_official) & missing(XR_dec)
save XR,replace



* ============================================================================
* 2. EUROSTAT-OECD BENCHMARK PPPs
*    Source: Eurostat-OECD March 2026 release
* ============================================================================

import excel "EUO_benchmark.xlsx", clear firstrow
rename (Country Series) (Country_Code Category)

save EUO_benchmark, replace


* ============================================================================
* 3. INCOME CLASSIFICATIONS
*    Source: World Bank FY2023 and FY2019 classifications
*    Note:   FY23 classification used for the 2021 benchmark round;
*            FY19 classification used for the 2017 benchmark round.
* ============================================================================

import excel "Income_classificationFY23.xlsx", clear firstrow
gen Time_Code = "YR2021"
save Income_classificationFY23, replace

import excel "Income_classificationFY19.xlsx", clear firstrow
gen Time_Code = "YR2017"
save Income_classificationFY19, replace


* ============================================================================
* 4. CPI / INFLATION DATA
*    Source: SNA economy database, pulled April 2026
* ============================================================================

import delimited "Inflation20212026-05-12.csv", clear

rename (country date category ratio_2017 ratio_2011) ///
       (Country_Code Time Category Inflation_2017 Inflation_2011)

* Drop observations missing inflation for the relevant reference year
drop if (Inflation_2017 == "NA" & Time == 2021) | (Inflation_2011 == "NA" & Time == 2017)

destring Inflation_2017, replace

* Harmonise time code format
tostring Time, gen(Time_Code)
replace Time_Code = "YR" + Time_Code

* Harmonise category codes to WDI indicator codes
replace Category = "PA.NUS.PPP"      if Category == "GDP"
replace Category = "PA.NUS.PRVT.PP"  if Category == "HHC"

save Inflation_20172021, replace


* ============================================================================
* 5. GDP PER CAPITA — WDI
*    Source: WDI (constant and current USD series)
* ============================================================================

import excel "GDPpc_WDI.xlsx", clear firstrow

rename (CountryName CountryCode) (Country_Name Country_Code)
keep Country_Name Country_Code SeriesCode VersionCode YR*
drop if missing(Country_Code)

reshape long YR, i(Country_Name Country_Code VersionCode SeriesCode) j(Time)
drop if YR == ".."
destring YR, replace
rename YR GDPpc

replace SeriesCode = "Constant_USD2010" if SeriesCode == "NY.GDP.PCAP.KD"
replace SeriesCode = "Current_USD"  if SeriesCode == "NY.GDP.PCAP.CD"
replace SeriesCode = "Current_LCU"  if SeriesCode == "NY.GDP.PCAP.CN"

drop if VersionCode=="202405" & Time==2017
drop VersionCode

reshape wide GDPpc, i(Country_Name Country_Code Time) j(SeriesCode) string

drop if missing(GDPpcCurrent_LCU)
tostring Time, gen(Time_Code)
replace Time_Code = "YR" + Time_Code

save GDPpc_WDI, replace


* ============================================================================
* 6. GDP PER CAPITA — ICP
*    Source: ICP data bank (PPP-adjusted)
* ============================================================================

import excel "GDP_pc_ICP.xlsx", clear firstrow

rename (CountryName CountryCode) (Country_Name Country_Code)
keep Country_Name Country_Code YR*
drop if missing(Country_Code)

reshape long YR, i(Country_Name Country_Code) j(Time)
destring YR, replace
rename YR GDPpcPPP

tostring Time, gen(Time_Code)
replace Time_Code = "YR" + Time_Code

* Correct country code for Russia
replace Country_Code = "RUS" if Country_Code == "RUT"

save GDP_pc_ICP, replace


* ============================================================================
* 7. EXTRAPOLATED PPPs AND FINAL DATASET BUILD
*    Source: SNA deflators, April 2026
* ============================================================================

import delimited "ExtrapolationOECD_andmissing2017and2021_2026-05-12.csv", clear

rename (country date category ppp exr) ///
       (Country_Code Time Category extrap xr)

tostring Time, gen(Time_Code)
replace Time_Code = "YR" + Time_Code

* Harmonise category codes to WDI indicator codes
replace Category = "PA.NUS.PPP"     if Category == "GDP"
replace Category = "PA.NUS.PRVT.PP" if Category == "HHC"

* Tag data vintage; flag countries where source data was obtained externally
gen Vintage_def = "WDI May 2026"

replace Vintage_def = "IFS/national bureau of statistics (pulled externally)" ///
    if inlist(Country_Code, "ARG", "COD", "DJI", "HTI", "IRN", "LBN", "MMR", "UZB") ///
    & Time == 2017

replace Vintage_def = "IFS/national bureau of statistics (pulled externally)" ///
    if inlist(Country_Code, "ARG", "COD", "TJK", "BMU", "SWZ", "CUW") ///
    & Time == 2021

* Drop observations missing extrapolated or benchmark values
drop if extrap == "NA" | benchmark == "NA"

destring xr, replace
destring extrap, replace
destring benchmark, replace

* Correct region classification for Russia
replace region_code = "c_CIS" if Country_Code == "RUS"


* Merge in Eurostat-OECD benchmarks
//merge 1:1 Country_Code Time_Code Category using EUO_benchmark, update nogen
//drop v1
rename region_code ICPregion

* Merge in income classifications and per capita income data
merge m:1 Country_Code using Income_classificationFY23, keep(match) nogen
merge m:1 Country_Code Time_Code using GDP_pc_ICP,                keep(match) nogen
merge m:1 Country_Code Time_Code using GDPpc_WDI,                 keep(match) nogen

* Drop Kosovo (not an ICP participant)
drop if Country_Code == "XKX"

destring GDPpcPPP, replace

save ExtrapolationandEUO_20172021, replace


* ============================================================================
* 8. Trade data
*    Source: WDI and UNCTAD
* ============================================================================

import excel "WDI_tradedata.xlsx", clear firstrow

rename (CountryName CountryCode) (Country_Name Country_Code)
drop if missing(Country_Code)

reshape long YR, i(Country_Name Country_Code SeriesName SeriesCode) j(Time)

replace SeriesCode="TermsofTrade" if SeriesCode=="TT.PRI.MRCH.XD.WD"
replace SeriesCode="TradetoGDP"  if SeriesCode=="NE.TRD.GNFS.ZS"
replace SeriesCode="EXPcapacityIMP"  if SeriesCode=="NY.EXP.CAPM.KN"

replace YR="" if YR==".."
destring YR, replace
tostring Time, gen(Time_Code)
replace Time_Code = "YR" + Time_Code

drop SeriesName
reshape wide YR, i(Country_Name Country_Code Time Time_Code) j(SeriesCode) string
ren YR* *
bys Country_Code (Time): gen id=_n
replace TermsofTrade=TermsofTrade[_n-1] if Country_Code=="SDN" & Time==2012
bys Country_Code (Time): egen TT_2011=total(TermsofTrade*(Time==2011))
bys Country_Code (Time): egen TT_2017=total(TermsofTrade*(Time==2017))
gen TermsofTrade_rebased=TermsofTrade/TT_2011*100 if Time<=2017
replace TermsofTrade_rebased=TermsofTrade/TT_2017*100 if Time>2017
replace TermsofTrade_rebased=TermsofTrade_rebased-100
keep if Time==2017 | Time==2021
keep Country_Code Time TradetoGDP TermsofTrade TermsofTrade_rebased EXPcapacityIMP

save WDI_tradedata, replace




**********************************************************************************************************************************************************************************
***********************************************************************************2021*******************************************************************************************
***********************************************************************************************************************************************************************************

* ============================================================================
* 1. IMPORT WDI ARCHIVE DATA (2021 ROUND — GDP AND HHC)
*    Source: WDI Database Archives — "PPP benchmark replacement in WDI" extract
* ============================================================================

local filepath "C:\Users\wb589699\OneDrive - WBG\Desktop\WDIvsBenchmark_analysis\P_Data_Extract_From_WDI_Database_Archives_PPP benchmark replacement in WDI.xlsx"

import excel using "`filepath'", sheet("2021_HHC") firstrow clear
save HHC_2021, replace

import excel using "`filepath'", sheet("2021_GDP") firstrow clear
append using HHC_2021

gen Vintage_def = "WDI March 2024"
rename (CountryCode TimeCode SeriesCode) (Country_Code Time_Code Category)
drop if missing(Country_Code)
destring Time, replace



* Rename vintage columns: Mar2024 = extrapolation, May2024 = benchmark
rename (Mar202403 May202405) (extrap benchmark)

* Drop the extrapolation column (only benchmark is retained from this source)



* ============================================================================
* 2. MERGE WITH MAIN EXTRAPOLATION DATASET
* ============================================================================

merge 1:1 Time_Code Country_Code Category using ExtrapolationandEUO_20172021, update nogen
drop if Time == 2017

* Merge in exchange rates
merge m:1 Country_Code Time_Code using XR, nogen keep(match)
* ============================================================================
* 3. MANUAL EXCHANGE RATE CORRECTIONS
*    Zimbabwe:  use the correct PPP-based XR for GDP; set XR to 1 for HHC
*    Liberia:   benchmark is in USD, so divide by XR then reset XR to 1
* ============================================================================

replace benchmark = benchmark / xr   if Country_Code == "LBR" & Category == "PA.NUS.PRVT.PP"
replace xr        = 1                if Country_Code == "LBR"

replace extrap = extrap/7.5345  if Country_Code=="HRV" //2021 PPPs were estimated in EUO, but original 2017 used for extrapolation still in old denomination.

replace XR_official=XR_dec if missing(XR_official) | Country_Code=="ZWE" | Country_Code=="LBR" | Country_Code=="PSE" | Country_Code=="LBN"  


* ============================================================================
* 4. SAMPLE RESTRICTIONS
* ============================================================================

* Drop non-benchmark and 2017 observations
drop if inlist(ICPregion, "j_NBM", "new_NBM") | Time_Code == "YR2017"

* Drop small economies not included in the analysis
drop if inlist(Country_Code, "AIA", "FSM", "KIR", "MSR", "PLW")
drop if inlist(Country_Code, "PNG", "SLB", "SSD", "TON", "VUT", "WSM",  "HTI", "MMR")
drop if inlist(Country_Code, "BHS", "BRB", "IRN", "SXM")


drop if Country_Code=="ZWE" &  Category == "PA.NUS.PRVT.PP"

///DEC NA

* ============================================================================
* 5. FILL MISSING CLASSIFICATION AND REGION VARIABLES
*    Use modal value within country to fill gaps from the merge
* ============================================================================

bysort Country_Code: egen temp1 = mode(IncomeClassification)
replace IncomeClassification = temp1 if missing(IncomeClassification)
drop temp1

bysort Country_Code: egen temp2 = mode(ICPregion)
replace ICPregion = temp2 if missing(ICPregion)
drop temp2


* ============================================================================
* 6. FINALISE AND MERGE INFLATION DATA
* ============================================================================

replace CountryName = Country_Name if missing(CountryName)

merge 1:1 Country_Code Time_Code Category using Inflation_20172021, nogen keep(match)


* ============================================================================
* 7. GENERATE PLI VARIABLES AND EXPORT
* ============================================================================

* Price Level Indices (PPP divided by exchange rate)
gen pli_ext   = extrap    / XR_dec if Vintage_def=="WDI March 2024"
replace pli_ext=extrap/ xr if missing(pli_ext)

gen pli_bench = benchmark / XR_dec if Vintage_def=="WDI March 2024"
replace pli_bench=benchmark/ xr if missing(pli_bench)

ren (XR_official xr) (XR_DECMay2024 XR_May2026)

sort Country_Code Category

keep  Time Category Country_Code extrap benchmark XR_DECMay2024 XR_May2026 pli_ext pli_bench  ///
       IncomeClassification ICPregion     ///
       GDPpcPPP GDPpcConstant_USD2010 GDPpcCurrent_LCU GDPpcCurrent_USD   ///
       Vintage Inflation_2017

order Time Category Country_Code extrap benchmark XR_DECMay2024 XR_May2026 pli_ext pli_bench  ///
       IncomeClassification ICPregion    ///
       GDPpcPPP GDPpcConstant_USD2010 GDPpcCurrent_LCU GDPpcCurrent_USD ///
       Inflation_2017  Vintage

export excel "ICPvsWDI_dataset_analysis.xlsx", sheet("_2021", replace) firstrow(variables)
save ICPvsWDI_2021, replace



******************************************************************************************************************************
*************************************************2017*************************************************************************
******************************************************************************************************************************

* ============================================================================
* 1. IMPORT WDI ARCHIVE DATA (2017 ROUND — GDP AND HHC)
*    Source: WDI Database Archives — "PPP benchmark replacement in WDI" extract
* ============================================================================

local filepath "C:\Users\wb589699\OneDrive - WBG\Desktop\WDIvsBenchmark_analysis\P_Data_Extract_From_WDI_Database_Archives_PPP benchmark replacement in WDI.xlsx"

import excel using "`filepath'", sheet("2017_HHC") firstrow clear
save HHC_2017, replace

import excel using "`filepath'", sheet("2017_GDP") firstrow clear
append using HHC_2017

rename (CountryCode TimeCode SeriesCode) (Country_Code Time_Code Category)
drop if missing(Country_Code)

gen Vintage_def="WDI Apr 2020"

* Rename vintage columns: Apr2020 = extrapolation, May2020 = benchmark
rename (Apr202004 May202005) (extrap benchmark)

* Drop the extrapolation column (will be populated from main dataset via update merge)

destring Time, replace


* ============================================================================
* 2. MERGE WITH MAIN EXTRAPOLATION DATASET
* ============================================================================

merge 1:1 Time_Code Country_Code Category using ExtrapolationandEUO_20172021, update nogen
* Merge in exchange rates
merge m:1 Country_Code Time_Code using XR, nogen keep(match)

drop if Time==2021

* ============================================================================
* 3. SAMPLE RESTRICTIONS
* ============================================================================

* Drop non-benchmark and 2021 observations
drop if inlist(ICPregion, "j_NBM", "new_NBM") | Time_Code == "YR2021"
drop if Time == 2021

* Drop countries excluded from the 2017 analysis
drop if inlist(Country_Code, "LBN", "UZB", "GTM", "DJI", "AIA", "FSM", "KIR", "MSR", "PLW")
drop if inlist(Country_Code, "PNG", "SLB", "SSD", "TON", "VUT", "WSM", "SOM")
* ============================================================================
* 4. MANUAL EXCHANGE RATE AND EXTRAPOLATION CORRECTIONS
*    Zimbabwe:  reset XR to the correct 2017 value; override GDP extrapolation;
*               set XR to 1 after corrections so PLI is expressed in local terms
*    Palestine: correct XR to the 2017 value
*    Sierra Leone: benchmark reported in old Leone (pre-redenomination), divide by 1000
* ============================================================================

* Palestine correction
* Sierra Leone: benchmark in old Leone units (pre-1000 redenomination)
replace extrap=extrap/3.45280 if Country_Code=="LTU" 
replace extrap=extrap/0.702804 if Country_Code=="LVA" 


replace XR_official=XR_dec if missing(XR_official) | Country_Code=="ZWE" | Country_Code=="LBR" | Country_Code=="LBN"  
replace XR_dec=xr if Country_Code=="PSE" 
replace XR_official=xr if Country_Code=="PSE" 


* ============================================================================
* 5. MERGE INFLATION DATA
* ============================================================================

merge 1:1 Country_Code Time_Code Category using Inflation_20172021, nogen keep(match)

destring Inflation_2011, replace

* Note: CPI index (2011–2017) missing for: ARG, BRB, COD, CYM, DJI, HTI,
*       IRN, LBN, MMR, SXM, TCA, TJK, UZB


* ============================================================================
* 6. FILL MISSING CLASSIFICATION AND REGION VARIABLES
*    Use modal value within country to fill gaps from the merge
* ============================================================================

bysort Country_Code: egen temp1 = mode(IncomeClassification)
replace IncomeClassification = temp1 if missing(IncomeClassification)
drop temp1

bysort Country_Code: egen temp2 = mode(ICPregion)
replace ICPregion = temp2 if missing(ICPregion)
drop temp2

replace CountryName = Country_Name if missing(CountryName)



* ============================================================================
* 7. GENERATE PLI VARIABLES AND EXPORT
* ============================================================================

* Price Level Indices (PPP divided by exchange rate)
gen pli_ext   = extrap    / XR_dec 
gen pli_bench = benchmark / XR_dec

ren (XR_dec xr) (XR_DECApr2020 XR_May2026)
sort Country_Code Category

keep  Time Category Country_Code extrap benchmark XR_DECApr2020 XR_May2026 pli_ext pli_bench  ///
       IncomeClassification ICPregion                      ///
      GDPpcPPP GDPpcConstant_USD2010 GDPpcCurrent_LCU GDPpcCurrent_USD                        ///
       Vintage Inflation_2011

order Time Category Country_Code extrap benchmark XR_DECApr2020 XR_May2026 pli_ext pli_bench  ///
       IncomeClassification ICPregion                      ///
      GDPpcPPP GDPpcConstant_USD2010 GDPpcCurrent_LCU GDPpcCurrent_USD                        ///
      Inflation_2011  Vintage

export excel "ICPvsWDI_dataset_analysis.xlsx", sheet("_2017", replace) firstrow(variables)
save ICPvsWDI_2017, replace
