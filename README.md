#Radon-Data-Analyzer
========
Radon, Temperature, Air Pressure, and Humidity data analyzer; Analyzes daily, monthly and weekly average, min, max and average error for the 4850 Davis Campus radon detector data.


Look how easy it is to use:

    1) Download 'reader_version_02(advanced_analyzer).py' and copy "SURF-RadonTrends.xlsx" to the same director as the py script.
    2) Change the value of the "sheetName" variable to the desired sheet name specified in the excel document and save the python script file.
    3) Compile and run 'reader_version_02(advanced_analyzer).py' and wait for the script to excute (FYI, it needs about 600 MB of RAM to load the excel doc so it will take some time to execute based on your computer (also it is programmed to run single threadedly)).
    

Features
--------

* reader_version_02(advanced_analyzer).py: takes in an excel file, converts it consolidated data array and returns analyzed daily, monthly and weekly average, min, max and average error for the 4850 Davis Campus radon detector data of the specified location.
* Daily_analyer.py: takes in an excel file, converts it consolidated data array and returns analyzed daily,  average, min, max and average error for the 4850 Davis Campus radon detector data of the specified location.
* Weekly_analyer.py: takes in an excel file, converts it consolidated data array and returns analyzed weekly average, min, max and average error for the 4850 Davis Campus radon detector data of the specified location.
* Monthly_analyzer.py: takes in an excel file, converts it consolidated data array and returns analyzed monthly average, min, max and average error for the 4850 Davis Campus radon detector data of the specified location.



Installation
------------

Download 'reader_version_02(advanced_analyzer).py', 'Daily_analyer.py:', 'Weekly_analyer.py' or 'Monthly_analyzer.py', specify "SURF-RadonTrends.xlsx" and execute the desired python script!

Contribute
----------

- Issue Tracker: github.com/Josephss/Radon-Data-Analyzer/issues
- Source Code: github.com/Josephss/Radon-Data-Analyzer

Support
-------

If you are having issues, please let us know.
We have a mailing list located at: info@marvelouscode.com

License
-------

The project is licensed under the BSD license.
