# Winter Limnology Network Form 2 Apps Script Custom Functions

These are custom functions created in Apps Script to be used in Google Sheets as a part of the Winter Limnology Network [project](https://winter-ice.github.io/winter-ice/). ChatGPT-4 was used as a jumping off point and debugger as I had no knowledge of JavaScript when starting this project

## Functions 

### generateSequence2()

This function takes an array of information (PI Name, Lake Code, and YYYY-MM of Sampling) and creates a general 'sampling event' ID that represents a sampling effort, as well as specific sample IDs for each sample taken (e.g., one of Total Phosphorous, one for Chlorophyll-A, etc.). This function can create IDs for up to 11 sampling events.

### storeCustomIds2()

This function takes a list of sampling event IDs (and their numerical versions) and stores them in a separate google sheet (Form 3). It checks for duplicates in the parameter database and will return an error if there are duplicates found. It also updates a hidden sheet in Form 2 that marks the sampling event as 'stored' so that it will not appear when that PI goes to store a new set of sampling IDs.

### generateLabels2()

*__DEFUNCT__* This function generates labels to be printed out for sample bottles, while also keeping formatting the same in the output sheet. This function was depreciated as the google sheet could not be formatted to match physical label sheets.

### stringtoHash()

This function takes the sampling event ID created by generateSequence2 and converts it into a hash code. 
