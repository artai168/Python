'''
    ---- Distance and Duration ----
    Python version 3.10.10

    Description:
    This Python program is based on Google map API to
    1) Search map coordinate of two addresses
    2) Find the distance of the two addresses
    3) The traffic time of the two addresses

    Setup:
    pip install googlemaps
'''

import googlemaps

def get_locations(api_key, addresses):
    gmaps = googlemaps.Client(key=api_key)
    locations = []

    for address in addresses:
        geocode_result = gmaps.geocode(address)

        if geocode_result:
            first_result = geocode_result[0]
            geometry = first_result['geometry']
            location = geometry['location']
            locations.append(location)
        else:
            locations.append(None)

    return locations

def calculate_distance(api_key, origin, destination):
    gmaps = googlemaps.Client(key=api_key)
    distance_matrix = gmaps.distance_matrix(origin, destination, mode="driving", units="metric")

    print(distance_matrix)

    if distance_matrix['status'] == 'OK':
        distance = distance_matrix['rows'][0]['elements'][0]['distance']['text']
        duration = distance_matrix['rows'][0]['elements'][0]['duration']['text']
        return distance, duration
        
    else:
        return None

#-------------------------------------------------------------------------
#---------------------- Addresses and Google API key ---------------------
#-------------------------------------------------------------------------
GOOGLE_PLACES_API_KEY= "API KEY"
addresses_to_search = ["ADDRESS 1","ADDRESS 2"]
#-------------------------------------------------------------------------

locations = get_locations(GOOGLE_PLACES_API_KEY, addresses_to_search)

#Caculate Distance
origin_coordinates = (locations[0]['lat'],locations[0]['lng'])
destination_coordinates = (locations[1]['lat'],locations[1]['lng'])

result = calculate_distance(GOOGLE_PLACES_API_KEY, origin_coordinates, destination_coordinates)
if result:
    distance, duration = result
    print(f"The distance between the locations is: {distance}")
    print(f"The travel time by driving is: {duration}")
else:
    print("Unable to calculate the distance and travel time.")
