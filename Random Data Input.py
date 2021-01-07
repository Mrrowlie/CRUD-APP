
from faker import Faker
import pymongo
import random
from dateutil.relativedelta import relativedelta
import datetime
fake = Faker('en_GB')




client = pymongo.MongoClient("mongodb+srv://Admin:Password@cluster0.xhsaj.mongodb.net/<dbname>?retryWrites=true&w=majority")
db = client.CrudApp
client_collection = db["CrudAppCollection"]




fake_units = [
    {"Unit_Court": "Mayfair Court",
     "Unit_Postcode": "NW3 5PD",
     "Unit_Road": "Mayfair Road",
     "Unit_City": "London",
     "Unit_House": ["Paris House","Milan House","New York House","Alicante House"],
     "RTM_Group": 1
     },
    {"Unit_Court": "Park Lane Court",
     "Unit_Postcode": "SW2 8HG",
     "Unit_Road": "Park Lane Road",
     "Unit_City": "London",
     "Unit_House": ["Wales House","England House","Scotland House"],
     "RTM_Group": 1
     },
    {"Unit_Court": "Westminster Court",
     "Unit_Postcode": "NE4 9JN",
     "Unit_Road": "Westminster Road",
     "Unit_City": "London",
     "Unit_House": ["Sweden House","Norway House","Finland House","Denmark House"],
     "RTM_Group": 2
     },
    {"Unit_Court": "Wimbledon Court",
     "Unit_Postcode": "SE5 5ID",
     "Unit_Road": "Wimbledon Road",
     "Unit_City": "London",
     "Unit_House": ["Dubai House","Beijing House","Texas House"],
     "RTM_Group": 2
     },
    {"Unit_Court": "Mayors Court",
     "Unit_Postcode": "NE1 9IJ",
     "Unit_Road": "Mayors Road",
     "Unit_City": "London",
     "Unit_House": ["Tokyo House","China House","Vietnam House","Cape Town House"],
     "RTM_Group": 3
     },
    {"Unit_Court": "Southbank Court",
     "Unit_Postcode": "NW7 5RD",
     "Unit_Road": "Southbank Road",
     "Unit_City": "London",
     "Unit_House": ["Jordan House","Turkey House","Africa House"],
     "RTM_Group": 3
     }]





for i in range(1000):
    unit_ran = random.randrange(0,6)
    service_date = fake.date_between(start_date='-2y', end_date='today')

    document =  {"RTM_Group": fake_units[unit_ran].get("RTM_Group"),
        "Unit_Full" : str(i + 1) + " " + fake_units[unit_ran].get("Unit_Court"),
        "Unit": i + 1,
        "Unit_Court": fake_units[unit_ran].get("Unit_Court"),
        "Unit_House": random.choice(fake_units[unit_ran].get("Unit_House")),
        "Unit_Road": fake_units[unit_ran].get("Unit_Road"),
        "Unit_City": fake_units[unit_ran].get("Unit_City"),
        "Unit_Postcode": fake_units[unit_ran].get("Unit_Postcode"),
        "Ownership": "Leasehold",
        "Name": fake.name(),
        "Address_Line_1": str(random.randrange(1,500)) + " " + fake.street_name() + " " + fake.street_suffix(),
        "Address_Line_2": "",
        "Address_Line_3": "",
        "City": fake.city() + " " + fake.city_suffix(),
        "County": fake.city(),
        "Postcode": fake.postcode(),
        "Country":"United Kingdom",
        "Phone": fake.phone_number(),
        "Email_1": fake.ascii_email(),
        "Email_2": fake.ascii_email(),
        "Director": random.choice(["Yes","No"]),
        "Notes": "Random Test Data.",
        "Last_Serviced_ISO": datetime.datetime.combine(service_date,  datetime.time.min),
        "Service_Due_ISO": datetime.datetime.combine(service_date + relativedelta(years=1),  datetime.time.min),
        "Service_Due_Month": (service_date +  relativedelta(years=1)).month,
        "Service_Due_Year": (service_date +  relativedelta(years=1)).year,
        "RTM_Email": fake.ascii_email(),
        "RTM_Member": random.choice(["Yes","No"])}

    client_collection.insert_one(document)
    print(document)