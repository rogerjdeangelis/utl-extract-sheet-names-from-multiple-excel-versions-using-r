%let pgm=utl-extract-sheet-names-from-mutiple-excel-versions-using-r;

github
https://tinyurl.com/ymr7bwkm
https://github.com/rogerjdeangelis/utl-extract-sheet-names-from-multiple-excel-versions-using-r

Extract sheet names from mutiple excel versions using r

  Solutions

      1, R xlconnect
      2, WPS libname engines

You don't need SAS access to excel for this.
Excel libname and passthru part of base WPS

Ops Question
I'd like to be able to read the sheet names from earlier and later versions of Excel Workbooks.
https://listserv.uga.edu/scripts/wa-UGA.exe?A2=SAS-L;68540271.2306D&S=


related
https://github.com/rogerjdeangelis/Export-exel-sheet-names-to-sas-dataset

/*                   _
(_)_ __  _ __  _   _| |_
| | `_ \| `_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
*/

/*----  Create legacy and current excel workbooks                        ----*/

%utl_submit_r64('
library(xlsx);
data(iris);
write.xlsx(x = iris,file = "d:/xls/iris.xls", sheetName = "type-xls" );
write.xlsx(x = iris,file = "d:/xls/iris.xlsx",sheetName = "type-xlsx");
');

/**************************************************************************************************************************/
/*                                                                                                                        */
/* FYI: My boot drive is only 128gb (esy to image)                                                                        */
/*                                                                                                                        */
/* volume data (2tb)                                                                                                      */
/* Volume serial number is 00000012 D839:A195                                                                             */
/*                                                                                                                        */
/* Legacy excel format xls                                                                                                */
/*                                                                                                                        */
/* Folder PATH listing for                                                                                                */
/*                                                                                                                        */
/* D:\XLS                                                                                                                 */
/*     iris.xls   8kb  sheet name "type_xls"                                                                              */
/*     iris.xlsx 23kb  sheet name "type_xlsx"                                                                             */
/*                                                                                                                        */
/* OUTPUT                                                                                                                 */
/* ======                                                                                                                 */
/*                                                                                                                        */
/* %put The legacy excel(xls) and the current(xlsx) sheet names are &sheets respectively;                                 */
/*                                                                                                                        */
/* The legacy excel(xls) and the current(xlsx) sheet names are TYPE-XLS TYPE-XLSX respectively                            */
/*                                                                                                                        */
/**************************************************************************************************************************/

/*
/ |    _ __
| |   | `__|
| |_  | |
|_(_) |_|

*/
%symdel sheets_xlsx sheets_xls / nowarn;
%utl_submit_r64('
   library("XLConnect");
   wb <- loadWorkbook("d:/xls/iris.xlsx");
   sheets_xlsx<-getSheets(wb);
   wb <- loadWorkbook("d:/xls/iris.xls");
   sheets_xls<-getSheets(wb);
   sheets_xlsx;
   sheets_xls;
   writeClipboard(as.character(paste(sheets_xls,sheets_xlsx, collapse = " ")));
',return=sheets);

%put The legacy excel(xls) and the current(xlsx) sheet names are &sheets respectively;

 /*          _               _
  ___  _   _| |_ _  __  _   _| |_
 / _ \| | | | __| `_ \| | | | __|
| (_) | |_| | |_| |_) | |_| | |_
 \___/ \__,_|\__| .__/ \__,_|\__|
                |_|
*/

LOG

>
library("XLConnect");
wb <- loadWorkbook("d:/xls/iris.xlsx");
sheets_xlsx<-getSheets(wb);
wb <- loadWorkbook("d:/xls/iris.xls");
sheets_xls<-getSheets(wb);
sheets_xlsx;
sheets_xls;
writeClipboard(as.character(paste(sheets_xls,sheets_xlsx, collapse = " ")));

[1] "type-xlsx"
[1] "type-xls"
>

/**************************************************************************************************************************/
/*                                                                                                                        */
/* %put The legacy excel(xls) and the current(xlsx) sheet names are &sheets respectively;                                 */
/*                                                                                                                        */
/* The legacy excel(xls) and the current(xlsx) sheet names are type-xls type-xlsx respectively                            */
/*                                                                                                                        */
/**************************************************************************************************************************/

/*___                           _ _ _
|___ \    __      ___ __  ___  | (_) |__  _ __   __ _ _ __ ___   ___
  __) |   \ \ /\ / / `_ \/ __| | | | `_ \| `_ \ / _` | `_ ` _ \ / _ \
 / __/ _   \ V  V /| |_) \__ \ | | | |_) | | | | (_| | | | | | |  __/
|_____(_)   \_/\_/ | .__/|___/ |_|_|_.__/|_| |_|\__,_|_| |_| |_|\___|
                   |_|
*/
%utl_submit_wps64('
 libname xlsx excel "d:/xls/iris.xlsx";
 ods exclude all;
 ods output members=type_xlsx;
 proc contents data=xlsx._all_;
 run;quit;
 ods output members=type_xls;
 libname xls excel "d:/xls/iris.xls";
 proc contents data=xls._all_;
 run;quit;
 ods select all;
 data want;
   set type_xls type_xlsx;
   put name=;
 run;quit;
 proc print data=want;
 run;quit;
');

 /*          _               _
  ___  _   _| |_ _  __  _   _| |_
 / _ \| | | | __| `_ \| | | | __|
| (_) | |_| | |_| |_) | |_| | |_
 \___/ \__,_|\__| .__/ \__,_|\__|
                |_|
*/

/**************************************************************************************************************************/
/*                                                                                                                        */
/* The WPS System                                                                                                         */
/*                                                                                                                        */
/* Obs    NUMBER       NAME        MEMTYPE                                                                                */
/*                                                                                                                        */
/*  1        1      'TYPE-XLS$'     DATA                                                                                  */
/*  2        1      'TYPE-XLSX$     DATA                                                                                  */
/*                                                                                                                        */
/**************************************************************************************************************************/

/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/

