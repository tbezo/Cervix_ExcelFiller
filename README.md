# Cervix_ExcelFiller
ESAPI Binary Plugin (visual studio solution) for filling an Excel Sheet.

C# Code that showcases how to use the Open XML SDK to fill cells in an Excel sheet with dose information from Eclipse.

A template file is copied from the path provided in the "templateFile" variable to the location in "outFileExcel" (with an added patient last name and 
current year as subfolder).

To use the Plugin with the original Excel Sheet provided by GEC ESTRO (https://www.estro.org/ESTRO/media/ESTRO/About/hdr-gyn_biol-physik-formular_2017.xlsx)
you have to customize the row entries in the FractionData Class first. Also the regular expressions inside the ExtractFractionData() function have to be adjusted to match your clinic's naming convention.

Open XML SDK can be installed via nuget. The ESAPI reference might have to get corrected.
