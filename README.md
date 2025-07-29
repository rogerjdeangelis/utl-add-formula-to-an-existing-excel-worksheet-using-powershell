# utl-add-formula-to-an-existing-excel-worksheet-using-powershell
Add formula to an existing excel worksheet using powershell
    %let pgm=utl-add-formula-to-an-existing-excel-worksheet-using-powershell;

    Add formula to worksheet using sas powershell r and python

    PROBLEM
      Add formula to an existing excel worksheet using powershell

    %stop_submission;

    SOLUTIONS

        Powershell add BMI column to existing workbook
        (This is much harder? Then adding column when creating workbook)

        Related repos on end

    Python and r see links below

    SOAPBOX ON
      There are a lot of options other than DDE,
    SOAPBOX OFF

    github
    https://tinyurl.com/yjvrtk2y
    https://github.com/rogerjdeangelis/utl-add-formula-to-an-existing-excel-worksheet-using-powershell

    communities.sas
    https://tinyurl.com/3zptrs2z
    https://communities.sas.com/t5/SAS-Programming/Using-DDE-with-excel/m-p/754607#M237998


    /****************************************************************************************************************************/
    /* INPUT                       | PROCESS                                                   |OUTPUT                          */
    /* ====                        | =======                                                   |======                          */
    /* EXISTING WORKBOOK           | %utlfkil(d:\xls\formulasps.xlsx);                         | ------------------+            */
    /*                             | ods excel file='d:\xls\formulasps.xlsx' style=excel;      | |A1|fx      | NAME|            */
    /* d:\xls\formulasps.xlsx      | ods excel options(sheet_name="Sheet1")  ;                 | ------------------------------+*/
    /*                             | proc report data=sd1.have ;                               | [_]|   A    |  B  |  C  |  D  |*/
    /* ----------------------+     | column name hgt wgt;                                      | ------------------------------|*/
    /* | A1| fx      | NAME  |     | run;                                                      |  1 |NAME    | HGT | WGT | BMI |*/
    /* ---------------  ---------+ | ods excel close;                                          |  --|--------+-----+-----+-----|*/
    /* [_] |    A    |   B |   C | |                                                           |  2 | Alfred | 69  | 69  |16.61|*/
    /* --------------------------| | %utl_psbegin;                                             |  --|--------+-----+-----+-----|*/
    /*  1  | NAME    |HGT  |WGT  | | parmcards4;                                               |  3 | Alice  | 56.5| 56.5|18.49|*/
    /*  -- |---------+-----+-----| | # Open Excel and the workbook                             |  --|--------+-----+-----+-----|*/
    /*  2  |  Alfred |69   |69   | | $excel = New-Object -ComObject Excel.Application          |  4 | Barbara| 65.3| 65.3|16.15|*/
    /*  -- |---------+-----+-----| | $excel.Visible = $false                                   |  --|--------+-----+-----+-----|*/
    /*  3  |  Alice  |56.5 |56.5 | | $workbook=$excel.Workbooks.Open("d:\xls\formulasps.xlsx") |  5 | Carol  | 62.8| 62.8|18.27|*/
    /*  -- |---------+-----+-----| | $worksheet = $workbook.Sheets.Item("Sheet1")              |  --|--------+-----+-----+-----|*/
    /*  4  |  Barbara|65.3 |65.3 | |                                                           |  6 | Henry  | 63.5| 63.5|17.87|*/
    /*  -- |---------+-----+-----| | $worksheet.Cells.Item(1, 4) = "BMI"                       |  --|--------+-----+-----+-----|*/
    /*  5  |  Carol  |62.8 |62.8 | |                                                           |  7 | James  | 57.3| 57.3|17.77|*/
    /*  -- |---------+-----+-----| | $lastRow = $worksheet.UsedRange.Rows.Count                |  --|--------+-----+-----+-----+*/
    /*  6  |  Henry  |63.5 |63.5 | |                                                           | [Sheet1]                       */
    /*  -- |---------+-----+-----| | for ($row = 2; $row -le $lastRow; $row++) {               |                                */
    /*  7  |  James  |57.3 |57.3 | |     $cell = $worksheet.Cells.Item($row, 4)                |                                */
    /*  -- |---------+-----+-----| |     $cell.Formula = "=C$row/(B$row^2)*703"                |                                */
    /* [Sheet1]                    |  }                                                        |                                */
    /*                             |                                                           |                                */
    /* options validvarname=upcase;| $workbook.Save()                                          |                                */
    /* libname sd1 "d:/sd1";       | $workbook.Close()                                         |                                */
    /* data sd1.have;              | $excel.Quit()                                             |                                */
    /*   input                     |                                                           |                                */
    /*     name$ hgt wgt;          | # Release COM objects                                     |                                */
    /* cards4;                     | [System.Runtime.Interopservices.Marshal]::`               |                                */
    /* Alfred  69.0 112.5          | ReleaseComObject($worksheet) | Out-Null                   |                                */
    /* Alice   56.5  84.0          | [System.Runtime.Interopservices.Marshal]::`               |                                */
    /* Barbara 65.3  98.0          | ReleaseComObject($workbook) | Out-Null                    |                                */
    /* Carol   62.8 102.5          | [System.Runtime.Interopservices.Marshal]::`               |                                */
    /* Henry   63.5 102.5          | ReleaseComObject($excel) | Out-Null                       |                                */
    /* James   57.3  83.0          | [System.GC]::Collect()                                    |                                */
    /* ;;;;                        | [System.GC]::WaitForPendingFinalizers()                   |                                */
    /* run;quit;                   | ;;;;                                                      |                                */
    /*                             | %utl_psend;                                               |                                */
    /* %utlfkil(                   |                                                           |                                */
    /* d:\xls\formulasps.xlsx);    |                                                           |                                */
    /* ods excel                   |                                                           |                                */
    /* file=                       |                                                           |                                */
    /* 'd:\xls\formulasps.xlsx';   |                                                           |                                */
    /* ods excel options(          |                                                           |                                */
    /* sheet_name="Sheet1")        |                                                           |                                */
    /*  style=excel;               |                                                           |                                */
    /* proc report data=sd1.have;  |                                                           |                                */
    /* column name hgt wgt;        |                                                           |                                */
    /* run;                        |                                                           |                                */
    /* ods excel close;            |                                                           |                                */
    /*                             |                                                           |                                */
    /*                             |                                                           |                                */
    /****************************************************************************************************************************/

    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */

    options validvarname=upcase;
    libname sd1 "d:/sd1";
    data sd1.have;
      input
        name$ hgt wgt;
    cards4;
    Alfred  69.0 112.5
    Alice   56.5  84.0
    Barbara 65.3  98.0
    Carol   62.8 102.5
    Henry   63.5 102.5
    James   57.3  83.0
    ;;;;
    run;quit;

    %utlfkil(
    d:\xls\formulasps.xlsx);
    ods excel
    file=
    'd:\xls\formulasps.xlsx' style=excel;
    ods excel options(
    sheet_name="Sheet1")  ;
    proc report data=sd1.have;
    column name hgt wgt;
    run;
    ods excel close;

    /****************************************************************************************************************************/
    /* EXISTING WORKBOOK                                                                                                        */
    /*                                                                                                                          */
    /* d:\xls\formulasps.xlsx                                                                                                   */
    /*                                                                                                                          */
    /* ----------------------+                                                                                                  */
    /* | A1| fx      | NAME  |                                                                                                  */
    /* ---------------  ---------+                                                                                              */
    /* [_] |    A    |   B |   C |                                                                                              */
    /* --------------------------|                                                                                              */
    /*  1  | NAME    |HGT  |WGT  |                                                                                              */
    /*  -- |---------+-----+-----|                                                                                              */
    /*  2  |  Alfred |69   |69   |                                                                                              */
    /*  -- |---------+-----+-----|                                                                                              */
    /*  3  |  Alice  |56.5 |56.5 |                                                                                              */
    /*  -- |---------+-----+-----|                                                                                              */
    /*  4  |  Barbara|65.3 |65.3 |                                                                                              */
    /*  -- |---------+-----+-----|                                                                                              */
    /*  5  |  Carol  |62.8 |62.8 |                                                                                              */
    /*  -- |---------+-----+-----|                                                                                              */
    /*  6  |  Henry  |63.5 |63.5 |                                                                                              */
    /*  -- |---------+-----+-----|                                                                                              */
    /*  7  |  James  |57.3 |57.3 |                                                                                              */
    /*  -- |---------+-----+-----|                                                                                              */
    /* [Sheet1]                                                                                                                 */
    /****************************************************************************************************************************/

    /*
     _ __  _ __ ___   ___ ___  ___ ___
    | `_ \| `__/ _ \ / __/ _ \/ __/ __|
    | |_) | | | (_) | (_|  __/\__ \__ \
    | .__/|_|  \___/ \___\___||___/___/
    |_|
    */

    %utl_psbegin;
    parmcards4;
    # Open Excel and the workbook
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook=$excel.Workbooks.Open("d:\xls\formulasps.xlsx")
    $worksheet = $workbook.Sheets.Item("Sheet1")

    $worksheet.Cells.Item(1, 4) = "BMI"

    $lastRow = $worksheet.UsedRange.Rows.Count

    for ($row = 2; $row -le $lastRow; $row++) {
        $cell = $worksheet.Cells.Item($row, 4)
        $cell.Formula = "=C$row/(B$row^2)*703"
     }

    $workbook.Save()
    $workbook.Close()
    $excel.Quit()

    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::`
    ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::`
    ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::`
    ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    ;;;;
    %utl_psend;


    /****************************************************************************************************************************/
    /* d:\xls\formulasps.xlsx                                                                                                   */
    /*                                                                                                                          */
    /* ------------------+                                                                                                      */
    /* |A1|fx      | NAME|                                                                                                      */
    /* ------------------------------+                                                                                          */
    /* [_]|   A    |  B  |  C  |  D  |                                                                                          */
    /* ------------------------------|                                                                                          */
    /*  1 |NAME    | HGT | WGT | BMI |                                                                                          */
    /*  --|--------+-----+-----+-----|                                                                                          */
    /*  2 | Alfred | 69  | 69  |16.61|                                                                                          */
    /*  --|--------+-----+-----+-----|                                                                                          */
    /*  3 | Alice  | 56.5| 56.5|18.49|                                                                                          */
    /*  --|--------+-----+-----+-----|                                                                                          */
    /*  4 | Barbara| 65.3| 65.3|16.15|                                                                                          */
    /*  --|--------+-----+-----+-----|                                                                                          */
    /*  5 | Carol  | 62.8| 62.8|18.27|                                                                                          */
    /*  --|--------+-----+-----+-----|                                                                                          */
    /*  6 | Henry  | 63.5| 63.5|17.87|                                                                                          */
    /*  --|--------+-----+-----+-----|                                                                                          */
    /*  7 | James  | 57.3| 57.3|17.77|                                                                                          */
    /*  --|--------+-----+-----+-----+                                                                                          */
    /* [Sheet1]                                                                                                                 */
    /****************************************************************************************************************************/


     _ __ ___ _ __   ___  ___
    | `__/ _ \ `_ \ / _ \/ __|
    | | |  __/ |_) | (_) \__ \
    |_|  \___| .__/ \___/|___/
             |_|

    https://github.com/rogerjdeangelis/utl_excel_add_formulas
    https://github.com/rogerjdeangelis/utl-sending-a-formula-to-excel-to-reference-a-cell-in-another-sheet
    https://github.com/rogerjdeangelis/utl-using-only-r-openxlsx-to-add-excel-formulas-to-an-existing-sheet
    https://github.com/rogerjdeangelis/utl-using-sql-instead-the-excel-formula-language-for-solving-excel-problems-pyodbc
    https://github.com/rogerjdeangelis/utl_excel_add_formula_inplace

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
 
*/
