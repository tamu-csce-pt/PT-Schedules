# Main page:
1. Download the office hours sheet as a .xlsx file and run `python3 .\office_hours_automate.py`
2. Add 2nd column of resulting file to PT Website Information sheet under Office Hours - Script
3. Download Website Information tab in the google sheet as tsv, then add that tsv to this repo
4. Make sure all photos in drive are accessible to anyone with the link
5. Turn variables skipPics and isOldSemester to False in `updateWebsite.py` if first time running
6. Run `python3 ./updateWebsite.py "<filename.tsv>"`

# Side pages:

1. Manually edit any page besides `index.html`, such as weekly reviews, awards, and gallery
2. To update Previous PTs page, rerun the updateWebsite script with `isOldSemester` as True

### *Notes*

`skipPics` when true doesn't re-download pics from drive, saves time. Only needs to be false the first time you run the script for the isOldSemester
`isOldSemester` when true will generate the page with only names and pictures, no labs or office hours.
Push updated index.html file to github to rebuild website