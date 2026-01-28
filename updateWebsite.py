import requests
import argparse
import pandas as pd

skipPics = True # CHANGE THIS WHEN NEEDED
isOldSemester = True # Leave False unless you're adding to Previous Peer Teachers

parser = argparse.ArgumentParser()
parser.add_argument("tsvfile", type=str, metavar='str')

def getCourseString(pt):
    oneHundred = f'{pt['100s']}'
    twoHundred = f'{pt['200s']}'
    threeHundred = f'{pt['300s']}'
    l = []
    if oneHundred != '' and oneHundred is not None and oneHundred != 'nan':
        l.append(f'{oneHundred}')
    if twoHundred != '' and twoHundred is not None and twoHundred != 'nan':
        l.append(f'{twoHundred}')
    if threeHundred != '' and threeHundred is not None and threeHundred != 'nan':
        l.append(f'{threeHundred}')
    # print(l)
    if len(l) == 0:
        return "TBA"
    return f"{', '.join(l)}"

def get313CourseString(pt):
    three12 = f'{pt['312']}'
    three13 = f'{pt['313']}'
    three14 = f'{pt['314']}'
    three31 = f'{pt['331']}'
    l = []
    if three12 != '' and three12 is not None and three12 != 'nan':
        l.append(f'{three12}')
    if three13 != '' and three13 is not None and three13 != 'nan':
        l.append(f'{three13}')
    if three14 != '' and three14 is not None and three14 != 'nan':
        l.append(f'{three14}')
    if three31 != '' and three31 is not None and three31 != 'nan':
        l.append(f'{three31}')
    # print(l)
    if len(l) == 0:
        return ""
    return f"<li>{', '.join(l)}</li>"

def getLabString(labs: str):
    retString = ''
    if isinstance(labs, str):
        for lab in labs.split(','):
            retString += f'<li>CSCE {lab.strip()}</li>'
    if retString == '':
        retString = "TBA"
    return retString

def getOfficeHourString(office_hours: str):
    # print(office_hours)
    retString = ''
    if isinstance(office_hours, str):
        for oh in office_hours.split(','):
            retString += f'<li>{oh.strip()}</li>'
    if retString == '':
        retString = "TBA"
    return retString

def main(skipPics):
    args = parser.parse_args()
    websiteFile = args.tsvfile
    if not websiteFile:
        print("Usage:")
        print("python3 ./updateWebsite.py <path_to_tsvfile>")
        exit(0)
    
    websiteInformation = pd.read_csv(websiteFile, sep='\t')
    
    websiteDict = websiteInformation.to_dict(orient='records')
    
    # print(websiteDict)
    
    imageLinks = {}
    
    htmlString = ""
    
    for pt in websiteDict:
        # print(pt['Image'])
        name = pt['Name']
        if not skipPics:
            driveLink = pt['Pictures']
            id = driveLink.split(r'/')[-1].split(r'?')[-1]
            # print(id)
            response = requests.get(f'https://drive.google.com/uc?{id}')
            if response.status_code == 200:
                print(pt['Name'])
                imageLinks[pt['Name']] = f'./images/{pt["Name"]}.webp'
                with open(imageLinks[pt['Name']], 'wb') as f:
                    for chunk in response.iter_content(1024):
                        f.write(chunk)
        else:
            imageLinks[pt['Name']] = f'./images/{pt["Name"]}.webp'
        # '''
        #     <h3><img alt="William Aalund" class="float-right" height="125" src="../../_files/_images/_content-images/CSCE-peer-William-Aalund.jpg" width="85"/></h3>
        #     <h3>William Aalund</h3>
        #     <p><strong>Email:</strong> williamraalund@tamu.edu</p>
        #     <ul>
        #     <li>100, 200, 300 Level Courses</li>
        #     <li>CSCE 313, 314, 331</li>
        #     <li>Labs 331-908, 331-909</li>
        #     </ul>
        #     <p><strong>Office hours:</strong></p>
        #     <ul>
        #     <li>Tuesday 10:30-11:30 a.m.</li>
        #     <li>Thursday 9:30-10:30 a.m.</li>
        #     <li>Friday 9:00-11:00 a.m., 1:00-2:00 p.m.</li>
        #     </ul>
        # '''
        
        if isOldSemester:
            html = f'''
                <div class="peerTeacher">
                <h3><img alt="{name}" class="float-right" src="{imageLinks[name]}" width="auto" height="150"/></h3>
                <h3>{name}</h3>
                <br>
                <br>
                <br>
                <br>
                </div>
            '''
        else:
            html = f'''
                <div class="peerTeacher">
                <h3><img alt="{name}" class="float-right" src="{imageLinks[name]}" width="auto" height="150"/></h3>
                <h3>{name}</h3>
                <p class = "ContentP"><strong>Email:</strong> {pt["Email"]}</p>
                <ul>
                <li>{getCourseString(pt)}</li>
                {get313CourseString(pt)}
                </ul>
                <p class = "ContentP"><strong>Labs:</strong></p>
                <ul>
                {getLabString(pt['Labs'])}
                </ul>
                <p class = "ContentP"><strong>Office Hours:</strong></p>
                <ul>
                {getOfficeHourString(pt['Office Hours - Script'])}
                </ul>
                </div>
            '''
        htmlString += html

    with open('index.html', 'w') as indexFile:
        indexFile.truncate(0)

        with open('basehtml.html', 'r') as base:
            text = base.read()
            indexFile.write(text.split("<!--CODE-->")[0])
            indexFile.write(htmlString)
            indexFile.write(text.split("<!--CODE-->")[1])

if __name__ == "__main__":
    main(skipPics)