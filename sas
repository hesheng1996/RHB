%let position_dt="31JUL2023"D;

data _null_;
     call symput('DATE',put(intnx('month',&position_dt,0,'E'),date9.));
     call symput('prev_date',put(intnx('month',&position_dt,-1,'E'),date9.));
run;

%put &DATE;
****************************************************************************************************************
proc sort data= RPT_HEADER_2
    out= RPT_HEADER_2 ;by CUST_ID descending FINANCIAL_YEAR_END  descending ID    ; run;

proc sort data= RPT_HEADER_2
          out= RPT_HEADER_3 nodupkey ;by CUST_ID ; run;


proc sort data= WORK.WORKING_MM_0622
          out= group1 nodupkey dupout=dups ;by CUSTOMER_NAME REMARK AMOUNT; run;

****************************************************************************************************************
case when substr(ID_NUMBER,1,3) = '000' then substr(ID_NUMBER,4,(length(ID_NUMBER)-3))
     when substr(ID_NUMBER,1,2) = '00' then substr(ID_NUMBER,3,(length(ID_NUMBER)-2))
     when substr(ID_NUMBER,1,1) = '0' then substr(ID_NUMBER,2,(length(ID_NUMBER)-1))
     else ID_NUMBER end as ID_NUMBER_NEW

****************************************************************************************************************
proc transpose data=final_chk1 out= final_chk1 (drop = _NAME_) ;
by ADJ_DIVISION_CD;
ID CUST_ID_map;
var count;
run;

****************************************************************************************************************
%MACRO REGRESSION(INDEP_VAR);
 
DATA &INDEP_VAR._DATA (KEEP=NPL_RATIO &INDEP_VAR);
SET CONS_EWS_INPUT;
RUN;
 
/*NO LAG*/
DATA &INDEP_VAR._LAG_DATA_0;
SET &INDEP_VAR._DATA;
WHERE &INDEP_VAR NE .;
RUN;

/* LAG 1Q*/
DATA &INDEP_VAR._LAG_DATA_1;
SET &INDEP_VAR._DATA;
LAG_1Q_&INDEP_VAR = LAG3(&INDEP_VAR);
RUN;


DATA &INDEP_VAR._LAG_DATA_1;
SET &INDEP_VAR._LAG_DATA_1;
WHERE LAG_1Q_&INDEP_VAR NE .;
RUN;


/* LAG 2Q*/
DATA &INDEP_VAR._LAG_DATA_2;
SET &INDEP_VAR._DATA;
LAG_2Q_&INDEP_VAR = LAG6(&INDEP_VAR);
RUN;


DATA &INDEP_VAR._LAG_DATA_2;
SET &INDEP_VAR._LAG_DATA_2;
WHERE LAG_2Q_&INDEP_VAR NE .;
RUN;


/* LAG 4Q*/
DATA &INDEP_VAR._LAG_DATA_4;
SET &INDEP_VAR._DATA;
LAG_4Q_&INDEP_VAR = LAG12(&INDEP_VAR);
RUN;


DATA &INDEP_VAR._LAG_DATA_4;
SET &INDEP_VAR._LAG_DATA_4;
WHERE LAG_4Q_&INDEP_VAR NE .;
RUN;

/* STEP 1 : TO RUN REGRESSION AGAINST NO_LAG INDEP_VAR VARIABLE */
 
PROC REG DATA=&INDEP_VAR._LAG_DATA_0 TABLEOUT RSQUARE ADJRSQ NOPRINT
outest=&INDEP_VAR._lag_no_lag (where=(_TYPE_ in ("PARMS","PVALUE")) drop=NPL_RATIO _IN_ _P_ _EDF_);
     &INDEP_VAR.:MODEL NPL_RATIO = &INDEP_VAR / DW VIF SPEC;
/*     OUTPUT OUT=&INDEP_VAR._LAG_DATA R=RESID;*/
RUN;
QUIT;

data FINAL_&INDEP_VAR._lag_0q (drop=_TYPE_ _RMSE_);
rename _DEPVAR_=dep_var _MODEL_=indep_var;
retain _DEPVAR_ _MODEL_ lag_count r_square adj_r_square coff_x k pvalue_x pvalue_k;
merge &INDEP_VAR._lag_no_lag (in=a where=(_TYPE_="PARMS") rename=(&INDEP_VAR.=coff_x Intercept=k _RSQ_=r_square _ADJRSQ_=adj_r_square))
            &INDEP_VAR._lag_no_lag (in=b where=(_TYPE_="PVALUE") drop=_RSQ_ _ADJRSQ_ rename=(&INDEP_VAR.=pvalue_x Intercept=pvalue_k));
by _MODEL_ _DEPVAR_ _RMSE_;
if a;
format coff_x k pvalue_x pvalue_k r_square adj_r_square 12.8;
format lag_count $10.;
lag_count = "No Lag";
run;

 
/* STEP 2 : TO RUN REGRESSION AGAINST LAG_1Q INDEP_VAR VARIABLE */
 
PROC REG DATA=&INDEP_VAR._LAG_DATA_1 tableout RSQUARE ADJRSQ noprint
outest=&INDEP_VAR._lag_1q (where=(_TYPE_ in ("PARMS","PVALUE")) drop=NPL_RATIO _IN_ _P_ _EDF_);
     &INDEP_VAR.:MODEL NPL_RATIO = LAG_1Q_&INDEP_VAR / DW VIF SPEC;
/*     OUTPUT OUT=&INDEP_VAR._LAG_DATA R=RESID_1Q;*/
RUN;
QUIT;

data FINAL_&INDEP_VAR._lag_1q (drop=_TYPE_ _RMSE_);
rename _DEPVAR_=dep_var _MODEL_=indep_var;
retain _DEPVAR_ _MODEL_ lag_count r_square adj_r_square coff_x k pvalue_x pvalue_k;
merge &INDEP_VAR._lag_1q (in=a where=(_TYPE_="PARMS") rename=(lag_1q_&INDEP_VAR.=coff_x Intercept=k _RSQ_=r_square _ADJRSQ_=adj_r_square))
            &INDEP_VAR._lag_1q (in=b where=(_TYPE_="PVALUE") drop=_RSQ_ _ADJRSQ_ rename=(lag_1q_&INDEP_VAR.=pvalue_x Intercept=pvalue_k));
by _MODEL_ _DEPVAR_ _RMSE_;
if a;
format coff_x k pvalue_x pvalue_k r_square adj_r_square 12.8;
format lag_count $10.;
lag_count = "Lag 1Q";
run;


/* STEP 3 : TO RUN REGRESSION AGAINST LAG_2Q INDEP_VAR VARIABLE */
 
PROC REG DATA=&INDEP_VAR._LAG_DATA_2 tableout RSQUARE ADJRSQ noprint
outest=&INDEP_VAR._lag_2q (where=(_TYPE_ in ("PARMS","PVALUE")) drop=NPL_RATIO _IN_ _P_ _EDF_);
     &INDEP_VAR.:MODEL NPL_RATIO = LAG_2Q_&INDEP_VAR / DW VIF SPEC;
/*     OUTPUT OUT=&INDEP_VAR._LAG_DATA R=RESID_2Q;*/
RUN;
QUIT;

data FINAL_&INDEP_VAR._lag_2q (drop=_TYPE_ _RMSE_);
rename _DEPVAR_=dep_var _MODEL_=indep_var;
retain _DEPVAR_ _MODEL_ lag_count r_square adj_r_square coff_x k pvalue_x pvalue_k;
merge &INDEP_VAR._lag_2q (in=a where=(_TYPE_="PARMS") rename=(lag_2q_&INDEP_VAR.=coff_x Intercept=k _RSQ_=r_square _ADJRSQ_=adj_r_square))
            &INDEP_VAR._lag_2q (in=b where=(_TYPE_="PVALUE") drop=_RSQ_ _ADJRSQ_ rename=(lag_2q_&INDEP_VAR.=pvalue_x Intercept=pvalue_k));
by _MODEL_ _DEPVAR_ _RMSE_;
if a;
format coff_x k pvalue_x pvalue_k r_square adj_r_square 12.8;
format lag_count $10.;
lag_count = "Lag 2Q";
run;


/* STEP 4 : TO RUN REGRESSION AGAINST LAG_4Q INDEP_VAR VARIABLE */
 
PROC REG DATA=&INDEP_VAR._LAG_DATA_4 tableout RSQUARE ADJRSQ noprint
outest=&INDEP_VAR._lag_4q (where=(_TYPE_ in ("PARMS","PVALUE")) drop=NPL_RATIO _IN_ _P_ _EDF_);
     &INDEP_VAR.:MODEL NPL_RATIO = LAG_4Q_&INDEP_VAR/ DW VIF SPEC;
/*     OUTPUT OUT=&INDEP_VAR._LAG_DATA R=RESID_4Q;*/
RUN;
QUIT;

data FINAL_&INDEP_VAR._lag_4q (drop=_TYPE_ _RMSE_);
rename _DEPVAR_=dep_var _MODEL_=indep_var;
retain _DEPVAR_ _MODEL_ lag_count r_square adj_r_square coff_x k pvalue_x pvalue_k;
merge &INDEP_VAR._lag_4q (in=a where=(_TYPE_="PARMS") rename=(lag_4q_&INDEP_VAR.=coff_x Intercept=k _RSQ_=r_square _ADJRSQ_=adj_r_square))
            &INDEP_VAR._lag_4q (in=b where=(_TYPE_="PVALUE") drop=_RSQ_ _ADJRSQ_ rename=(lag_4q_&INDEP_VAR.=pvalue_x Intercept=pvalue_k));
by _MODEL_ _DEPVAR_ _RMSE_;
if a;
format coff_x k pvalue_x pvalue_k r_square adj_r_square 12.8;
format lag_count $10.;
lag_count = "Lag 4Q";
run;

data FULL_&INDEP_VAR;
set FINAL_&INDEP_VAR._lag_:;
run;

PROC DELETE
DATA=
&INDEP_VAR._DATA
&INDEP_VAR._LAG_DATA
&INDEP_VAR._lag_no_lag
&INDEP_VAR._lag_1q
&INDEP_VAR._lag_2q
&INDEP_VAR._lag_4q
FINAL_&INDEP_VAR._lag_0q
FINAL_&INDEP_VAR._lag_1q
FINAL_&INDEP_VAR._lag_2q
FINAL_&INDEP_VAR._lag_4q
;
RUN;
 
%MEND;
%REGRESSION(FTSE_EMAS);
%REGRESSION(CONS_EDF);
%REGRESSION(BULK_CEMENT);
%REGRESSION(BULK_CEMENT_YOY);
%REGRESSION(BULK_CEMENT_MA);
%REGRESSION(CONCRETE_PROD);
%REGRESSION(CONS_GDP);
%REGRESSION(CONS_GDP_x_OL);
%REGRESSION(CONS_VAC);
%REGRESSION(CONS_VAC_MA);
%REGRESSION(STEEL_BAR);
%REGRESSION(CONS_INDEX);
%REGRESSION(WORK_DONE);
%REGRESSION(WORK_DONE_YOY);
%REGRESSION(MY_GDP);
%REGRESSION(MY_GDP_YOY);
%REGRESSION(MY_REAL_GDP);
%REGRESSION(MY_REAL_GDP_YOY);
%REGRESSION(REAL_GDP_CAP);
%REGRESSION(REAL_GDP_CAP_YOY);
%REGRESSION(MY_GDP_QOQ);
%REGRESSION(MY_GDP_MA);

/*GDP renormalised NPL*/
%MACRO GDP(INDEP_VAR);
 
DATA &INDEP_VAR._DATA (KEEP=NPL_RATIO_2017_2022 &INDEP_VAR);
SET CONS_EWS_INPUT;
RUN;


/* LAG 4Q*/
DATA &INDEP_VAR._LAG_DATA_4;
SET &INDEP_VAR._DATA;
LAG_4Q_&INDEP_VAR = LAG12(&INDEP_VAR);
RUN;


DATA &INDEP_VAR._LAG_DATA_4;
SET &INDEP_VAR._LAG_DATA_4;
WHERE LAG_4Q_&INDEP_VAR NE .;
RUN;



/* STEP 4 : TO RUN REGRESSION AGAINST LAG_4Q INDEP_VAR VARIABLE */
 
PROC REG DATA=&INDEP_VAR._LAG_DATA_4 tableout RSQUARE ADJRSQ noprint
outest=&INDEP_VAR._lag_4q (where=(_TYPE_ in ("PARMS","PVALUE")) drop=NPL_RATIO_2017_2022 _IN_ _P_ _EDF_);
     &INDEP_VAR.:MODEL NPL_RATIO_2017_2022 = LAG_4Q_&INDEP_VAR/ DW VIF SPEC;
/*     OUTPUT OUT=&INDEP_VAR._LAG_DATA R=RESID_4Q;*/
RUN;
QUIT;

data FINAL_&INDEP_VAR._lag_4q (drop=_TYPE_ _RMSE_);
rename _DEPVAR_=dep_var _MODEL_=indep_var;
retain _DEPVAR_ _MODEL_ lag_count r_square adj_r_square coff_x k pvalue_x pvalue_k;
merge &INDEP_VAR._lag_4q (in=a where=(_TYPE_="PARMS") rename=(lag_4q_&INDEP_VAR.=coff_x Intercept=k _RSQ_=r_square _ADJRSQ_=adj_r_square))
            &INDEP_VAR._lag_4q (in=b where=(_TYPE_="PVALUE") drop=_RSQ_ _ADJRSQ_ rename=(lag_4q_&INDEP_VAR.=pvalue_x Intercept=pvalue_k));
by _MODEL_ _DEPVAR_ _RMSE_;
if a;
format coff_x k pvalue_x pvalue_k r_square adj_r_square 12.8;
format lag_count $10.;
lag_count = "Lag 4Q";
run;

data FULL_&INDEP_VAR._NEW;
set FINAL_&INDEP_VAR._lag_:;
run;

 
%MEND;
%GDP(MY_GDP_YOY);
%GDP(MY_REAL_GDP_YOY);


data CONS_EWS_OUTPUT_2022_0;
format dep_var $30.;
set FULL_FTSE_EMAS 
	FULL_CONS_EDF
	FULL_BULK_CEMENT
	FULL_BULK_CEMENT_YOY
	FULL_BULK_CEMENT_MA
	FULL_CONCRETE_PROD
	FULL_CONS_GDP
	FULL_CONS_GDP_x_OL
	FULL_CONS_VAC
	FULL_CONS_VAC_MA
	FULL_STEEL_BAR
	FULL_CONS_INDEX
	FULL_WORK_DONE
	FULL_WORK_DONE_YOY
	FULL_MY_GDP
	FULL_MY_GDP_YOY
	FULL_MY_REAL_GDP
	FULL_MY_REAL_GDP_YOY
	FULL_REAL_GDP_CAP
	FULL_REAL_GDP_CAP_YOY
	FULL_MY_GDP_QOQ
	FULL_MY_GDP_MA
	FULL_MY_GDP_YOY_NEW
	FULL_MY_REAL_GDP_YOY_NEW


	;run;

proc sql;
create table CRRPT.CONS_EWS_OUTPUT_2022 as select
dep_var,
cats(indep_var,'_',lag_count) as mapping,
indep_var,
lag_count,
r_square,
adj_r_square,
coff_x,
k,
pvalue_x,
pvalue_k



from CONS_EWS_OUTPUT_2022_0
;quit;

****************************************************************************************************************


PROC REG FUNCTION
_RSME_ - Root mean squared error
_IN_ - number of regressors in model
_P_ - number of parameters in model
_EDF_ - error degrees of freedom 
_RSQ_ - R-squared
_ADJRSQ_ - Adjusted R-squared
