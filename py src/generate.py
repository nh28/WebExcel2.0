from DataWebsite import DataWebsite
from  AnalyticsWebsite import AnalyticsWebsite

data_en = "1991-2020 Normals - Data Plan - DataT_EN.xlsx"
data_fr = "1991-2020 Normals - Data Plan - DataT_FR.xlsx"

an_en = "1991-2020 Normals - Data Plan - AnalyticsT_EN.xlsx"
an_fr = "1991-2020 Normals - Data Plan - AnalyticsT_FR.xlsx"

sheetDataEN = "Data- Archive-Web Tables"
sheetDataFR = "Data Archive Table-Web Table"
sheetAnalytics =	"Analytic- Archive-Web Tables"

webEN = "https://climate.weather.gc.ca/climate_normals/results_1991_2020_e.html?searchType=stnName_1991&txtStationName_1991=bagotville&searchMethod=contains&txtCentralLatMin=0&txtCentralLatSec=0&txtCentralLongMin=0&txtCentralLongSec=0&stnID=83000000&dispBack=1"
webFR = "https://climat.meteo.gc.ca/climate_normals/results_1991_2020_f.html?searchType=stnName_1991&txtStationName_1991=bagotville&searchMethod=contains&txtCentralLatMin=0&txtCentralLatSec=0&txtCentralLongMin=0&txtCentralLongSec=0&stnID=83000000&dispBack=1"

station_name = "BAGOTVILLE"
station_province = "QC"

try:
    print("Getting Data...Please Wait")

    dataEN = DataWebsite(data_en, sheetDataEN)
    dataEN.parse_website(webEN, dataEN.handler)
    print("1/4 complete")
    dataEN.handler.close_workbook()

    """dataFR = DataWebsite(data_fr, sheetDataFR)
    dataFR.parse_website(webFR, dataFR.handler)
    print("2/4 complete")
    dataFR.handler.close_workbook()

    analyticsEN = AnalyticsWebsite(an_en, sheetAnalytics)
    analyticsEN.parse_website(webEN, analyticsEN.handler)
    print("3/4 complete")

    analyticsFR = AnalyticsWebsite(an_fr, sheetAnalytics)
    analyticsFR.parse_website(webFR, analyticsFR.handler)
    print("4/4 complete")"""

except Exception as e:
    print("An error occurred:", e)

print("Done")		