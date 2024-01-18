

proc sort data= WORK.WORKING_MM_0622
          out= group1 nodupkey dupout=dups ;by CUSTOMER_NAME REMARK AMOUNT; run;



proc transpose data=final_chk1 out= final_chk1 (drop = _NAME_) ;
by ADJ_DIVISION_CD;
ID CUST_ID_map;
var count;
run;


PROC REG DATA=&INDEP_VAR._LAG_DATA TABLEOUT RSQUARE ADJRSQ 
outest=&INDEP_VAR._lag_no_lag (where=(_TYPE_ in ("PARMS","PVALUE")) drop=NPL_RATIO _IN_ _P_ _EDF_);
     &INDEP_VAR.:MODEL NPL_RATIO = &INDEP_VAR / DW VIF SPEC;

RUN;
QUIT;


data FINAL_&INDEP_VAR._lag_no_lag (drop=_TYPE_ _RMSE_);
rename _DEPVAR_=dep_var _MODEL_=indep_var;
retain _DEPVAR_ _MODEL_ lag_count r_square adj_r_square coff_x k pvalue_x pvalue_k;
merge &INDEP_VAR._lag_no_lag (in=a where=(_TYPE_="PARMS") rename=(&INDEP_VAR.=coff_x Intercept=k _RSQ_=r_square _ADJRSQ_=adj_r_square))
            &INDEP_VAR._lag_no_lag (in=b where=(_TYPE_="PVALUE") drop=_RSQ_ _ADJRSQ_ rename=(&INDEP_VAR.=pvalue_x Intercept=pvalue_k));
by _MODEL_ _DEPVAR_ _RMSE_;
if a;
format coff_x k pvalue_x pvalue_k r_square adj_r_square 12.8;
format lag_count $10.;
lag_count = "No Lag";
