Scripts to identify percent-ratio and ratio-confidence interval patterns were written in Visual Basic 6 (VB6). 
This directory contains all project files to run the scripts.

The steps necessary to run them on MEDLINE abstracts are as follows:
1) Obtain MEDLINE abstracts in XML format from NCBI (https://www.nlm.nih.gov/databases/download/pubmed_medline.html)
2) Start the VB6 project: ASEC.vbp
3) Run the Convert MEDLINE XML pre-processing step, changing the directories to where you put the XML files and where you want the processed files
4) Change the file directories in the percent-ratio and ratio-CI extraction routine to point to the processed file directory
5) Change the file increment numbers to go from 1 to whatever the last file number was in the downloaded MEDLINE files
6) Run either or both routines

    The output will be a tab-delimited file with the extracted and re-calculated values, along with % differences between the two. Sentence context
is also included for visual inspection of the results.