# Main page:
1. Download the office hours sheet as a .xlsx file and run `python3 .\office_hours_automate.py`
2. Download Website Information tab from the PT google sheet as tsv, then add that tsv to this repo
3. Add 2nd column (times) of resulting file to PT Website Information sheet (tsv) under Office Hours - Script
4. Ensure all photos in drive are accessible to anyone with the link and make sure there's an 'images' folder. If the current images folder is for a previous semester, rename 'images' folder to 'imagesSeasonYear'
5. Ensure photo links in the tsv are in the format 'https://drive.google.com/open?id=...'
6. Turn variables skipPics and isOldSemester to False in `updateWebsite.py` if first time running
7. Run `python3 ./updateWebsite.py "<filename.tsv>"`

# Side pages:

1. Manually edit any page besides `index.html`, such as weekly reviews, awards, and gallery
2. To update Previous PTs page, rerun the updateWebsite script with `isOldSemester` as True

### *Notes*

`skipPics` when true doesn't re-download pics from drive, saves time. Only needs to be false the first time you run the script for the isOldSemester

`isOldSemester` when true will generate the page with only names and pictures, no labs or office hours.
Push updated index.html file to github to rebuild website

For the beginning of a semester, ensure you change the main + week in review pages and add the group photo into the gallery