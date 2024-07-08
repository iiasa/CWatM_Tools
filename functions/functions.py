import pandas as pd
import xlrd, datetime, os
#xlrd.xlsx.ensure_elementtree_imported(False, None)
#xlrd.xlsx.Element_has_iter = True

def read_observations_excel(observations_folder, Excel_two_first_rows_not_read=False):

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
    #   ↵ StationName1.xlsx #two columns: Date, Observation
    #   ↵ StationName2.xlsx #two columns: Date, Observation
    #   ↵ ...

    # for the following, set Excel_two_first_rows_not_read = True. 
    # In this case, the headers can be anything, as they are not read. 
    # Dates are to be in the first column, and observations in the second.
    #   ↵ StationName1.xlsx #first two rows not read
    #   ↵ StationName2.xlsx #first two rows not read
    #   ↵ ...

    # Read list of observation stations

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

            if not Excel_two_first_rows_not_read:
                # version 1 of observations: two headers Date, Observation
                sheet = pd.read_excel(observed_discharge_folder + '/' + discharge_location[0] + '.xlsx')

            else:
                # version 2 of observations: no headers, first two rows are not read. 
                # Dates in the first column and observations in the second.
                sheet = pd.read_excel(observed_discharge_folder + '/' + discharge_location[0] + '.xlsx', 
                                         header=None, skiprows=2, 
                                         names=['Date', 'Observation'])

            sheet['Date'] = pd.to_datetime(sheet['Date']).dt.date
            Dates_observed = sheet['Date'].tolist()
            Flows_observed = sheet['Observation'].tolist()

            DATES_observed.append(Dates_observed)
            FLOWS_observed.append(Flows_observed)

        else:
            print('missing ' + discharge_location[0])
            DATES_observed.append([])
            FLOWS_observed.append([])

    return Observations_namesLocations_array, DATES_observed, FLOWS_observed