WF
==
position_parser.py
Written by: Laura Wacker 02/10/2013
Most Recent Update: 02/05/2014
This file reads a pdf, writes it to a text file titled "text_output.txt", and then reads and parses the text file to an excel document, titled "open_positions.xls" 
it requires python 2.7, pdfminer, xlrd, and xlwr libraries. 
 it runs in  O(n) time, where n is the number of lines in the text document (I can't speak for the runtime of the pdf->text method)

 Updates: 
 1. fixed bug with extra account 
 2. made contracts update dynamically with the changing months 
 3. automatically reads in the most recent pdf (based on a searching loop that iterates through dates backwards. You can also manually specify a pdf
   	and override the date searching functionality

 02/28/2014 Updates: 
 1. added a two tab layout to the excel sheet
 2. pos parser now archives each output in ./Archives folder 
 3. warnings for when gold is in silver account, when silver is in gold account and when there is anything in P0 account
 4. extended dates out farther 

03/04/2014 Updates:  
 1. for robustness - delete the open_positions.xls file before writing the data to a new fresh file. 
 2. for readability - put a parameter logging_flag into the get_recent_pdf() function - just used to deterimine whether or not to warn. 
	# Should fix this later - you should only need to calculate this once.... 
 

04/14/2014 Updates: 
 1. changed the xlswrite format from .xls to .xlsx - may fix Dad's copy and paste bug. 
 
 TO DO: 
	1. make a util file. 