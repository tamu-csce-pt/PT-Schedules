## To use:

1. Manually edit any page besides `index.html`, such as weekly reviews, awards, gallery, and previous PTs (for this one just rename the index.html file and create an updated one later)
2. Download the office hours sheet as a .xlsx file and run `python3 .\office_hours_automate.py`
3. Add of previous command result to PT sheet
4. Download PT website tab in the google sheet as tsv
5. Make sure all photos in drive are accessible to anyone with the link
6. Turn skipPics to False in `updateWebsite.py` if first time running
7. Run `python3 ./updateWebsite.py "<filename.tsv>"`
8. Commit updated index.html file to github
9. Profit
