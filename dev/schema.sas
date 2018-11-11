/* Note that value of libname is UPPERCASE */
proc sql;
	create table columns as
	select name as variable
	,memname as table_name
	from dictionary.columns
	where libname = 'WORK'
	;
quit;

/* SAS 9.4 or later */
ods excel file="c:\temp\variables.xlsx" style=minimal;
	proc print data=columns;
	run;
ods excel close;

/* earlier versions, using SAS/ACCESS to PC Files */
PROC EXPORT data = columns
	OUTFILE = 'variables.xls' 
	DBMS = EXCEL REPLACE;
	SHEET='VARLIST'; 
RUN;