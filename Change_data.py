jsonDataSource = "source.json"
nameExcel = 'new_events.xlsx'


def ChangeDateFormat(stringDate):
    #ex. 2023-11-06T21:02:08.438Z
    #from YYYY-MM-DDThh:mm:ssZ
    #to
    #date MM/DD/YYYY
    #time hh:mm
    import re
    pattern = "(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2})T(?P<time>\d{2}:\d{2})"

    matchesDate = re.finditer(pattern, stringDate)
    for match in matchesDate:
        dict = match.groupdict()
        date = dict["month"]+"/"+dict["day"]+"/"+dict["year"]
        #print(date)
        #print(dict['time'])
        return (date,dict['time'])
    print(stringDate)


#readsfile with location from google and returns array of data
def ReadFileLocationGoogle():
    import json
    data = []
    with open(jsonDataSource, "r") as read:
        dict = json.load(read)
        dict = dict["timelineObjects"]
        for point in dict:
            dataPoint = []
            if "placeVisit" in point.keys():
                location = point["placeVisit"]["location"]
                duration = point["placeVisit"]["duration"]
                if "name" in location.keys():
                    dataPoint.append(location["name"])
                else:
                    continue
                    dataPoint.append("noname")
                #ChangeDateFormat
                date, time = ChangeDateFormat(duration["startTimestamp"])
                dataPoint.append(date)
                dataPoint.append(time)


                dataPoint.append(location["address"])
                dataPoint.append(float(location["latitudeE7"])/10**7)
                dataPoint.append(float(location["longitudeE7"])/10**7)
                #print( point["placeVisit"]["location"]["name"])
                data.append(dataPoint)
    return data


def WriteToExcel(data):
    from openpyxl.workbook import Workbook
    workbook_name = nameExcel
    wb = Workbook()
    page = wb.active
    page.title = 'events'
    workbook = Workbook()
    headers = ['description','date','time','location','latitude','longitude']
    page.append(headers) # write the headers to the first line
    #data = [['description1','date1','time1','location1','latitude1','longitude1']]
    for info in data:
        page.append(info)
    wb.save(filename = workbook_name)

data = ReadFileLocationGoogle()
WriteToExcel(data)