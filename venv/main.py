from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
import datetime
import smtplib
from openpyxl import Workbook

browser = webdriver.Chrome(executable_path="/Users/gabrieldesseresusini/Downloads/chromedriver")

# -------------------- VARIABLES DECLARATIONS -------------------- #

# Select the required field we're going to search out of flights/hotel/restaurants/activities
flight_selection = '//button[@id="tab-flight-tab-hp"]'
# Select the type of flight tickets we want
return_ticket = '//label[@id="flight-type-roundtrip-label-hp-flight"]'
one_way_ticket = '//label[@id="flight-type-one-way-label-hp-flight"]'
round_trip = '//label[@id="flight-type-multi-dest-label-hp-flight"]'


# Select Search button
search_btn = '//button[@class="btn-primary btn-action gcw-submit"]'

# From City --- To City
from_city = 'Bordeaux'
to_city = 'Sydney, Australie'

# Website desired
link = "https://www.expedia.fr/"

depart_date = []



def choose_field():
    field = browser.find_element_by_xpath(flight_selection)
    field.click()
    # time.sleep(1)


def choose_ticket(ticket_type):
    try:
        ticket = browser.find_element_by_xpath(ticket_type)
        ticket.click()
        # time.sleep(2)
    except Exception as e:
        pass


def departure_city(element):
    dep_city = browser.find_element_by_xpath('//input[@id="flight-origin-hp-flight"]')
    # time.sleep(2)
    dep_city.clear()
    # time.sleep(1)
    dep_city.send_keys(' ' + element)
    time.sleep(1.5)
    first_onlist = browser.find_element_by_xpath('//a[@id="aria-option-0"]')
    first_onlist.click()


def arrival_city(element):
    arr_city = browser.find_element_by_xpath('//input[@id="flight-destination-hp-flight"]')
    # time.sleep(2)
    arr_city.clear()
    # time.sleep(1)
    arr_city.send_keys(' ' + element)
    time.sleep(1.5)
    first_onlist = browser.find_element_by_xpath('//a[@id="aria-option-0"]')
    first_onlist.click()


def departure_date(element):
    dep_date = browser.find_element_by_xpath('//input[@id="flight-departing-hp-flight"]')
    time.sleep(1)
    dep_date.clear()
    # time.sleep(1)
    dep_date.send_keys(element)
    time.sleep(1.5)


def return_date(day, month, year):
    retrn_date = browser.find_element_by_xpath('//input[@id="flight-returning-hp-flight"]')
    time.sleep(1)
    for i in range(11):
        retrn_date.send_keys(Keys.BACKSPACE)
    time.sleep(1)
    retrn_date.send_keys(day + '/' + month + '/' + year)
    time.sleep(1.5)


def search_flight():
    search = browser.find_element_by_xpath(search_btn)
    time.sleep(1)
    search.click()
    time.sleep(10)
    print("The flight results are ready")


# Creations of the dataFrame
dataFrame = pd.DataFrame()


def compile_data():
    global dataFrame
    global dep_time_list
    # global depart_date_list
    global arr_time_list
    # global return_date_list
    global airline_list
    global flight_duration_list
    global flight_stops_list
    global flight_price_list
    global layover_list

    # departure times
    dep_time = browser.find_elements_by_xpath('//span[@data-test-id="departure-time"]')
    dep_time_list = [value.text for value in dep_time]

    # arrival time
    arr_time = browser.find_elements_by_xpath('//span[@data-test-id="arrival-time"]')
    arr_time_list = [value.text for value in arr_time]

    # flight Company
    airline = browser.find_elements_by_xpath('//span[@data-test-id="airline-name"]')
    airline_list = [value.text for value in airline]

    # flight duration
    flight_duration = browser.find_elements_by_xpath('//span[@data-test-id="duration"]')
    flight_duration_list = [value.text for value in flight_duration]

    # flight stops
    flight_stops = browser.find_elements_by_xpath('//span[@class="number-stops"]')
    flight_stops_list = [value.text for value in flight_stops]

    # flight_price
    flight_price = browser.find_elements_by_xpath('//span[@data-test-id="listing-price-dollars"]')
    flight_price_list = [value.text for value in flight_price]

    # layovers
    layover = browser.find_elements_by_xpath("//span[@data-test-id='layover-airport-stops']")
    layover_list = [value.text for value in layover]


    # depart_date_list  = [value for value in depart_date]

    now = datetime.datetime.now()
    current_date = (str(now.day) + '-' + str(now.month) + '-' + str(now.year))
    current_time = (str(now.hour) + ':' + str(now.minute))
    current_price = 'price as of ' + '(' + current_date + '---' + current_time + ') for ' + departure_day

    for i in range(len(dep_time_list)):
        try:
            dataFrame.loc[i, 'dep_time'] = dep_time_list[i]
        except Exception as e:
            pass
        try:
            dataFrame.loc[i, 'arr_time'] = arr_time_list[i]
        except Exception as e:
            pass
        try:
            dataFrame.loc[i, 'airline'] = airline_list[i]
        except Exception as e:
            pass
        try:
            dataFrame.loc[i, 'flight_duration'] = flight_duration_list[i]
        except Exception as e:
            pass
        try:
            dataFrame.loc[i, 'flight_stops'] = flight_stops_list[i]
        except Exception as e:
            pass
        try:
            dataFrame.loc[i, 'layover'] = layover_list[i]
        except Exception as e:
            pass
        try:
            dataFrame.loc[i, str(current_price)] = flight_price_list[i]
        except Exception as e:
            pass


date = 1


for i in range(6):
    departure_day = str(date) + '/04/2020'
    # open up the browser page
    browser.get(link)

    time.sleep(5)

    # choose flight / hotel / restaurant
    choose_field()
    time.sleep(2)
    # choose ticket type
    choose_ticket(return_ticket)

    # choose departure city
    departure_city(from_city)

    # choose arrival city
    arrival_city(to_city)

    # depart_date.append(departure_day)
    # Set departure date
    departure_date(departure_day)
    date += 1

    # Set return date
    return_date('20', '07', '2020')



    search_flight()

    compile_data()

    dataFrame.to_excel('flights_prices.xlsx')

    print('Excel sheet created')
    time.sleep(5)
