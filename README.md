# COACalculator
Excel VBA macro that calculates expected COA based on student's enrollment, dependency status, housing status, and residency status.

## How to use
1) Download the CoaCalculator.xlsm file.
2) Open the file and enable macros.
3) In the "view" tab, click "macros" on the upper right. 
4) Select CoaCalculator and click "run".
5) Enter the following information:

    - Column to print the cost of attendance calculation (J in the sample).
    - Column containing dependency info (B in the sample).
    - Column containing housing info (C in the sample).
    - Column containing residency info (D in the sample).
    - Column containing enrollment info (E in the sample).
    - Number of semesters to calculate (3 in the sample).

![alt text](https://raw.githubusercontent.com/yerolaz/COACalculator/main/MacroGUI.PNG)

## Additional notes
The macro works by matching strings found in student information columns.
- Dependency column must only include "D" or "I".
- Housing column must only include "OFF_CAMPUS", "WITH_PARENT", or "ON_CAMPUS".
- Residency column must only include "IN" or "OUT".
- Enrollment columns must be integer or decimal numbers.

This can all be change by updating the code manually in developer mode.
