# luca_util
Various utilities for LUCA (Decennial Census project). Info can be found at 
https://www.census.gov/programs-surveys/decennial-census/about/luca.html

## addr.gen 
Generate suite/unit addresses for huge blocks of apartment buildings.

Address | Units
--- | ---
123 Fake St | 3
1122 Boogie Woogie Ave | 2

Address | Units
--- | ---
123 Fake St | Unit A
123 Fake St | Unit B
123 Fake St | Unit C
1122 Boogie Woogie Ave | Unit A
1122 Boogie Woogie Ave | Unit B

## cleanup
Just a way to convert address strings from internal geodatabase to LUCA entry format. Gets address number and st name from same column using simple regex.

## lams_upload
Used to enter data into custom online assessing forms. No bulk upload existed so I automated the process. Input is a spreadsheet as formatted below.

St number | St name | Unit type | Unit description | City letter | ZIP | Num. owners | Num. results | Which result
--------- | ------- | --------- | ---------------- | ----------- | --- | ----------- | ------------ | ------------
Address number | Name of street | In this case, 'u' for 'unit' | Name of unit | 'L' for 'London' | ZIP | 1 Owner | 2 hits for address search | Second result of 2 
21 | Baker St | u | B | L | 123456 | 1 | 2 | 2

*Note: Must determine last 3 values by using system*


## mouseNow.py
Shamlessly copied from Al Swiegert's https://github.com/asweigart/automatetheboringstuffwithpythondotcom
