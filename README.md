# Continente Scraper
## Python scraper for the latest bargains offered by Continente, the portuguese retail chain


[Continente](https://en.wikipedia.org/wiki/Continente_(supermarket)) is one of the top retail chains in Portugal.

As such, and given the rampant inflation currently hitting the Eurozone, making the most of every deal available doesn’t seem like a bad idea.

I therefore came up with this handy Python web scraper that scours Continente’s website for their current price discounts on specific predefined categories such as fruit, tea, body care products, etc, but also allows searches for specific brand names.
To add or remove categories, please edit the Categories.txt file. You can find each category's URL in Continente's [website](https://www.continente.pt).

The program will output an Excel file, where each category/brand name will have its own sheet and respective available deals. Each sheet will also come sorted by €/(Unit|Kg|L) as well as feature a unique tab and header color. Pretty neat, huh? In case you want to run the program from a BAT file, I threw in an old Continente logo for you to use as the Shortcut’s icon.

## Before running

Spreadsheet editing software like Excel will be required. Free alternatives like [LibreOffice](https://www.libreoffice.org/) will also get the job done.
Please fill out the Variables.txt with the required paths. 
Please check the requirements file to make sure you meet the necessary dependencies to run this program. Otherwise, open command prompt and: pip install -r "C:\path\to\requirements.txt"
