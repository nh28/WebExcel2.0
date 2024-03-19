from CSV import CsvProcessor
from DataWebsite import DataWebsite
from  AnalyticsWebsite import AnalyticsWebsite

data_en = "1991-2020 Normals - Data Plan - DataT_EN.xlsx"
data_fr = "1991-2020 Normals - Data Plan - DataT_FR.xlsx"

an_en = "1991-2020 Normals - Data Plan - AnalyticsT_EN.xlsx"
an_fr = "1991-2020 Normals - Data Plan - AnalyticsT_FR.xlsx"

sheetDataEN = "Data- Archive-Web Tables"
sheetDataFR = "Data Archive Table-Web Table"
sheetAnalytics =	"Analytic- Archive-Web Tables"

webEN = "https://climate-dev.cmc.ec.gc.ca/climate_normals/results_1991_2020_e.html?searchType=stnName_1991&txtStationName_1991=bagotville&searchMethod=contains&txtCentralLatMin=0&txtCentralLatSec=0&txtCentralLongMin=0&txtCentralLongSec=0&stnID=83000000&dispBack=1"
webFR = "https://climat-dev.cmc.ec.gc.ca/climate_normals/results_1991_2020_f.html?searchType=stnName_1991&txtStationName_1991=bagotville&searchMethod=contains&txtCentralLatMin=0&txtCentralLatSec=0&txtCentralLongMin=0&txtCentralLongSec=0&stnID=83000000&dispBack=1"

station_name = "BAGOTVILLE"
station_province = "QC"

try:
    print("Getting Data...Please Wait")

    dataEN = DataWebsite(data_en, sheetDataEN)
    dataEN.parse_website(webEN, dataEN.handler)
    print("1/6 complete")
    dataEN.handler.close_workbook()
    
    dataFR = DataWebsite(data_fr, sheetDataFR)
    dataFR.parse_website(webFR, dataFR.handler)
    print("2/6 complete")
    dataFR.handler.close_workbook()
    
    analyticsEN = AnalyticsWebsite(an_en, sheetAnalytics)
    analyticsEN.parse_website(webEN, analyticsEN.handler)
    print("3/6 complete")

    
    analyticsFR = AnalyticsWebsite(an_fr, sheetAnalytics)
    analyticsFR.parse_website(webFR, analyticsFR.handler)
    print("4/6 complete")
    
    csv = CsvProcessor()
    csv.process_csv_to_excel("en_1991-2020_Normals_" + station_province + "_" + station_name + ".csv", "1991-2020 Normals - Data Plan - DataCSV_EN.xlsx", sheetDataEN)
    print("5/6 complete")

    csv2 = CsvProcessor()
    csv2.process_csv_to_excel("fr_1991-2020_Normales_" + station_province + "_" + station_name + ".csv", "1991-2020 Normals - Data Plan - DataCSV_FR.xlsx", sheetDataEN)
    print("6/6 complete")


except Exception as e:
    print("An error occurred:", e)

print("Done")		