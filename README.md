# Main page:
Contains 2 things: Office hours + lab hours
1. Download Website Information tab from the PT google sheet as tsv, then add that tsv to this repo.
2. Download the office hours sheet as a .xlsx file and run `python3 .\office_hours_automate.py` to produce a csv
3. In excel, add the 2nd column (the times) of the csv to PT Website Information sheet (tsv) under Office Hours - Script
4. Ensure all photos in drive are accessible to anyone with the link and make sure there's an 'images' folder
5. Ensure photo links in the tsv are in the format 'https://drive.google.com/open?id=...' and to remove any erroroneous rows (which can be done easily in a text editor)
6. Turn variables skipPics and isOldSemester to False in `updateWebsite.py` if first time running
7. Run `python3 ./updateWebsite.py "<filename.tsv>"`
8. Rename 'images' folder to 'imagesSeasonYear'
9. Push to github

# Side pages:

1. Manually edit any page besides `index.html`, such as weekly reviews, awards, and gallery
2. To update Previous PTs page, rerun the updateWebsite script with `isOldSemester` as True

### *Notes*

`skipPics` when true doesn't re-download pics from drive, saves time. Only needs to be false the first time you run the script for the isOldSemester

`isOldSemester` when true will generate the page with only names and pictures, no labs or office hours.
Push updated index.html file to github to rebuild website

For the beginning of a semester, ensure you change the main + week in review pages and add the group photo into the gallery

Whenever you might need to update the individual PT images, update the drive link in the .tsv and make sure that the folder is called 'images' before running updateWebsite.py. Make sure to rename 'images' to 'imagesSeasonYear' after running the script.  
You can also just convert an image into a .webp and replace its original instead of using the script.