Purpose of this code is to automate the mobily and mobilelinkusa report generation. Currently data is received in an excel format in multiple sheets, it needs to be calculated and the resultant report needs to be generated.

Abbreviations:
TH - Total Hours
MO - Mobily
ML - Mobilelink


Sheets in Source file:
EDFT
TH Summary
TH Hourly
ML Store
Emp List
Apr Unapr OT
Check Pivot
Current Week OT
Sheet10
SCH
SCH for MO


OT Flag generally works that after regular 40 hours have exceeded, the earning desc on the same day would change and have OT in it with the hours.

For the OT, just take the values of OT and sum them
----------

Data is taken from TH hourly, joined with ML store. Then by having the OBT joined, the data for OT is extracted and pivot is used to show. 260422, was able to perform the join.

------
# Report
Sheets made from the report:

Pay Period Wise OT Summary
Current Week OT
E.wise OT (Week) Home Location
E.wise OT (PP) Home Location
TODO: Emp Hrs (current PP)