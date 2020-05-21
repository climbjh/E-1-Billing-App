# E-1-Billing-App
Application to take raw *.csv data and convert into usable billing table as *.xlsx

This app will recieve two auto-genrated *.csv files pulled from Tsheets and Foundation programs and use the Pandas framework in Python to parse them into one *.xlsx file, which will then be populated with "rate" information using an SQLite3 query and tables for Employee Pay Rate and Cost Code Rate.  This is designed to simplify the billing process for the Energy-1 team.

Ideally I would like to create an interface for the billing team to interact easily with the app - ie. "click here to select *.csv files"

The app will generate new *.csv files with the date of generation in the name, as well as a *.xlsx file with the date.  These will all be put into a newly generated folder with the date of creation as the name as a subfolder of "Billing"
