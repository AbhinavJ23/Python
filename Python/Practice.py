import urllib.request
import json


def print_results(data):
    jsonObj = json.loads(data)
    
    if "title" in jsonObj["metadata"]:
        print(jsonObj["metadata"]["title"])
    
    if "count" in jsonObj["metadata"]:
        count = jsonObj["metadata"]["count"]
        print ("Number of Events " + str(count))
    
    for i in jsonObj["features"]:
        print(i["properties"]["place"])
    print("------------------\n")
    
    for i in jsonObj["features"]:
        if i["properties"]["mag"] >= 4.0:
            print("%2.1f" % i["properties"]["mag"], i["properties"]["place"]) 
    print("------------------\n")

def main():
    urldata = "http://earthquake.usgs.gov/earthquakes/feed/v1.0/summary/2.5_day.geojson"
    
    url = urllib.request.urlopen(urldata)
    print("Result Code: ", url.getcode())
    if(url.getcode() == 200):
        data = url.read()
        print_results(data)
    else:
        print("Error accessing url " + str(url.getcode()))

if __name__ =="__main__":
    main()
    