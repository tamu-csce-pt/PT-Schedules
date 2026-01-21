## To use:

1. Manually edit any page besides `index.html`, such as weekly reviews, awards, gallery, and previous PTs (for this one just rename the index.html file and create an updated one later)
2. Download the office hours sheet as a .xlsx file and run `python3 .\office_hours_automate.py`
3. Add 2nd column of resulting file to PT Website Info sheet under Office Hours - Script
4. Download Website Information tab in the google sheet as tsv, then add that tsv to this repo (it should be gitignored)
5. Make sure all photos in drive are accessible to anyone with the link
6. Turn skipPics variable to False in `updateWebsite.py` if first time running
7. Run `python3 ./updateWebsite.py "<filename.tsv>"`
8. Commit updated index.html file to github
9. Profit
