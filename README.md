# Football BTTS system HTML scraper #


## Summary ##

Two scripts are included that work together and in this order:

* match_scraper.py
* result_scraper.py 


***match_scraper.py*** grabs the football matches, that score over 20 Points, from [this](http://www.over25tips.com/both-teams-to-score-tips) webpage and writes them in an .xlsx file, if they are not *already* written. If the .xlsx file does not exist, it creates it.

***result_scraper.py*** checks the football matches in the .xlsx file, grabs their betting odds and game results from [this](http://www.oddsportal.com/) site and appends them in the file as well.


## Changelog ##

### Version 1.0 ###

* Use UTC timezone on all websites
* Loads JavaScript content on oddsportal.com with PhantomJS
* Visually marks postponed games for deletion inside the spreadsheet file
* Visually marks non-found games inside the spreadsheet file. Usually an inaccurate game date by over25tips.com is the culprit
* Converts misloaded GB odds (7/30) to EU fractual odds (1.23)
* Calculates missing odds in oddsportal.com by averaging the rest of the odds (which is the mathematically correct way)
* Keeps backup spreadsheet (data_bak.xlsx)

## Dependencies ##

Except from the modules you'll need to install, like `BeautifoulSoup` and `openpyxl`, you'll need to download `PhantomJS`.
Add `PhantomJS.exe` to your PATH or copy it in the same folder as the *result_scraper.py* script.
