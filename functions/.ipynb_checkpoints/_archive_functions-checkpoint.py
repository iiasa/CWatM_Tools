import pandas as pd
# xlrd==1.2.0
import xlrd, datetime, os
xlrd.xlsx.ensure_elementtree_imported(False, None)
xlrd.xlsx.Element_has_iter = True

def read_observations_excel(observations_folder):

    # Returns three arrays:
    #   Observations_namesLocations_array, [Stations[i], Latitudes[i], Longitudes[i]]_i
    #   DATES_observed, [Observation dates for station i]_i
    #   FLOWS_observed, [Observation values for station i]_i

    # observations_folder holds one excel (observation_locations.xlsx) with names and locations, and
    # a folder (Observations) with an Excel file for each station. The Excel files are named after the station names.
    #
    # observations_locations.xlsx has three columns: Station name, Latitude, Longitude
    #
    # FOLDER STRUCTURE
    # observations_folder
    # ↵ observations_locations.xlsx
    # ↵ Observations
    #   ↵ StationName1.xlsx #first two rows not read
    #   ↵ StationName2.xlsx #first two rows not read
    #   ↵ ...

    # Read list of observation stations

    #%python

    import pandas
    
    Observations_namesLocations = pd.read_excel(observations_folder + '/observations_locations.xlsx', engine='openpyxl')

    Stations = Observations_namesLocations['Station'].tolist()
    Latitudes = Observations_namesLocations['Latitude'].tolist()
    Longitudes = Observations_namesLocations['Longitude'].tolist()

    Observations_namesLocations_array = []

    for i in range(len(Stations)):
        Observations_namesLocations_array.append([Stations[i], Latitudes[i], Longitudes[i]])

    # Read observation time series

    DATES_observed = []
    FLOWS_observed = []

    observed_discharge_folder = observations_folder + '/Observations'
    observed_folder_list = os.listdir(observed_discharge_folder)

    print('Loading observations')

    for discharge_location in Observations_namesLocations_array:

        if discharge_location[0] + '.xlsx' in observed_folder_list:
            book = xlrd.open_workbook(observed_discharge_folder + '/' + discharge_location[0] + '.xlsx')
            sheet = book.sheet_by_index(0)
            num_rows = sheet.nrows

            _Dates_observed = [xlrd.xldate_as_tuple(int(sheet.cell(row, 0).value), 0) for row in range(2, num_rows)]
            Dates_observed = [datetime.datetime(d[0], d[1], d[2]) for d in _Dates_observed]

            Flows_observed = [sheet.cell(row, 1).value for row in range(2, num_rows)]

            DATES_observed.append(Dates_observed)
            FLOWS_observed.append(Flows_observed)

        else:
            print('missing ' + discharge_location[0])
            DATES_observed.append([])
            FLOWS_observed.append([])

    return Observations_namesLocations_array, DATES_observed, FLOWS_observed

