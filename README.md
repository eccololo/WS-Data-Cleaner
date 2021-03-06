# WS Data Cleaner
## Description
This program cleans output files from [WebScraper](https://webscraper.io/) by deleting junk columns in output Excel file. It gather together all columns with our data from scraping and it creates main output files with that data.

For example we can have 10 Web Scraper files in one category folder and 10 WebScraper files in second category folder.

This program will clean that files from jung columns and it will create 2 main Excel files with all data per category column and additionaly in root folder it will create one main root Excel file with all data.

This program was written in Python with Openpyxl.
## Usage

You can use this program in 2 ways.
### Auto Mode

By using auto mode you let the program to handle for you everything like which columns to delete and how to name column headers in main root Excel file.

You just have to specify at the end of the program that you want to use auto mode.

### User Mode

In this mode you specify how many columns you want to delete in each Excel file with data starting from 1st until specified.

You also name header columns in main root Excel file.

This mode is for those who want to keep some junk columns and wants to name the output header columns by themselves.

### Additional Info

"Lorem Ipsum" Web Scraper data are in "Temp" folder. 

All files with data must be in "Data" folder in structure like "Data folder" -> "Category Folders" -> "Files with data" otherwise the program will not work.

## Demo

In this repository in "Data" folder there are already 2 Category folders with data already processed.\
In "copy" folder there are files copies before processing data, so they are unprocessed.
In "Done" folder there are files already processed.
In "Main" folder there is one main file with all data from files from this category folder.

If you want to check how this program works copy it to your localhost. Delete all folders from "Data" folder and copy all folders from "Temp" folder to "Data" folder.
Then you can run program and set mode to auto or user.

The output files should be the same as in this repository.


## Author

The author of this program is Mateusz Hyla. You can contact me through my blog [mstem.net](https://mstem.net) or social media.