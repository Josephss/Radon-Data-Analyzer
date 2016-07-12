#Radon-Data-Analyzer
========
Radon, Temperature, Air Pressure, and Humidity data analyzer; Analyzes daily, monthly and yearly average, min, max and average error for the 4850 Davis Campus radon detector data.
*Takes in an excel file, converts it consolidated data array and returns weeknumber, average, min value, max value, "error" and end date of the week*

Look how easy it is to use:

    1) Download 'reader_version_02(advanced_analyzer).py' and copy "SURF-RadonTrends.xlsx" to the same director as the py script.
    2) Change the value of the "sheetName" variable to the desired sheet name specified in the excel document and save the python script file.
    3) Compile and run 'reader_version_02(advanced_analyzer).py' and wait for the script to excute (FYI, it needs about 600 MB of RAM to load the excel doc so it will take some time to execute based on your computer (also it is programmed to run single threadedly)).
    

Features
--------

* Takes in an excel file, converts it consolidated data array and returns weeknumber, average, min value, max value, "error" and end date of the week radon, temperature, air pressure and humidity of the specified location.


Installation
------------

Download 'reader_version_02(advanced_analyzer).py', specify "SURF-RadonTrends.xlsx" and execute the python script!

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
