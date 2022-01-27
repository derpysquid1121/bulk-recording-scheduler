
# Google Apps script for bulk scheduling NCast AV recordings. 

This script is made for Google Apps Scripts. It works with a google sheet
in order to pull classes and their info and then create google calendar events 
that the recording hardware in the classrooms use to schedule their recordings
It keeps track of which classes were successfully scheduled and which ones were not.
If it was unsuccessful it highlights the specific data that it was unable to parse and checks
a certain box on that row to indicate what problem it ran into. 

It also creates and organizes a spreadsheet to keep track of the scheduled recordings.
There are other scripts that my manager made that are then used on that spreadsheet
to implement a process through which students who miss class can automatically be students
the recording of the class that they missed. 

## Authors

- [@davidblackburn](https://derpysquid1121.github.io/website/)

## Screenshots

![Bulk data spreadsheet](/readme_src/Capture1.PNG)
![Bulk data spreadsheet with successes and errors](/readme_src/Capture2.PNG)
![Scheduled Recording sheet](/readme_src/Capture3.PNG)