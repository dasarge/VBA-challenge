# VBA-challenge
broke this projects down into small components

worked with the message box to go through each sheet with a message box naming the sheet 
calculated the rows from each sheet 
this work is in module 1 in file alphabetical listing 1 test 
The next step used the code from the credit_charges-3 macro see credit card module1 and  module 2 (edited)
worked on a calculator for finding the last row and number of rows from 2 to the lastrow
this was edited to fit this file - this was used to consolidate the symbols and to sum the stock volume
worked to convert the yyyymmdd year format into mm/dd/yyyy format it turns out this was not really necessary
created the header text for the summary table
used formatting to auto fit the columns when the summary table is generated
used formulas and added values to the 
during this time  started to organize the sections of the code into sections this allowed for easier editing and locating different code groups
added a description area to the top of the code
added numbered section throughout the code to make the sections more understandable and to eliminate comment from the code area
created separate working files to test shorter versions of the test file this allowed the test file to remain intact incase i screwed something up
started to run the complete code together and troubleshoot the operation
checked the calculations found that the yearly change was not calculating correctly
the problem with this calculation is that the same i value is being used for both the yearly open and yearly closed
so the issue is that both values are using the i value for the end of each section or the closing i value
it looks like the yearly opening price will have to be calculated before the next iteration is complete
 there are two steps to this first the  first opening price will have to be calculated with an i value of i=2
