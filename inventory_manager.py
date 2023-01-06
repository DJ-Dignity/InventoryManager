import re
import decimal
import csv
import os
import configparser
import json

import pdfreader
from lxml import etree
from pymongo import MongoClient

# Global Variables
# Settings variables
csv_fields = ["Index", "Quantity", "Part Number", "Manufacturer Part Number", "Description", "Customer Reference", "Backorder", "Unit Price", "Extended Price"]
#settingsPath = "C:\\Users\\green\\PycharmProjects\\InventoryManager\\settings.xml"
ordered_parts = []
order_settings = {}  # Contains orderWD
suppliers_list = [] # Contains list of suppliers who's receipts can be parsed
mongo_settings = {} # Contains mongoURL, mongoUserName, and mongoPassword

# Part variables
# TODO: Change these to databases
# <partType>Optionals determines which indices in <partType>Parameters are optional and which must be found in the part description
enclosure_styles = {"1590BB2": "1590BB2", "1590BB": "1590BB", "1590B": "1590B", "125B": "125B", "1590A": "1590A", "1590LB": "1590LB", "1590XX": "1590XX", "1590DD": "1590DD",
                   "1032L": "1032L"}
enclosure_materials = {"Aluminum": "AL"}
enclosure_colors = {"White": "WHT", "Orange": "ORG", "Red": "RED", "Pink": "PNK", "Violet": "VIO", "Blue": "BLU", "Silver": "SLV", "Grey": "GRY", "Green": "GRN",
                   "Copper": "CPR", "Brown": "BRN", "Gold": "GLD", "Yellow": "YLW", "Champagne": "CMP", "Black": "BLK", "Cream": "CRM", "Chromium": "CHR", "Gray": "GRY"}
enclosure_finishes = {"Matte": "MAT", "Metallic": "MET", "Hammer": "HAM"}
enclosure_optionals = [False, False, False, True]
enclosure_parameters = [enclosure_styles, enclosure_materials, enclosure_colors, enclosure_finishes]

switch_types = {"Toggle": "TGL", "Momentary": "MOM", "Push Button": "PSH", "DIP": "DIP", "Rotary": "ROT", "Stomp Foot": "PSH", "Slide": "SLD", "Rocker": "RCK", "Tact": "TCT"}
switch_contacts = {"SPST": "SPST", "SPDT": "SPDT", "DPST": "DPST", "DPDT": "DPDT", "3PDT": "3PDT", "Encloder": "ENC", "1 Pole 7 Position": "1P7T", "1 Pole 8 Position": "1P8T",
                  "1 Pole 10 Position": "1P10T", "1 Pole 12 Position": "1P12T", "2 Pole 3 Position": "2P3T", "2 Pole 4 Position": "2P4T", "2 Pole 6 Position": "2P6T",
                  "3 Pole 4 Position": "3P4T", "4 Pole 3 Position": "4P3T"}
switch_options = {"Short": "SHT", "Right Angle": "RHT", "LED": "LED"}
switch_optionals = [False, True, True]
switch_parameters = [switch_types, switch_contacts, switch_options]

socket_sizes = {"6": "6", "8": "8", "14": "14", "16": "16", "18": "18", "20": "20", "24": "24", "28": "28", "32": "32", "40": "40", "42": "42"}
socket_optionals = [False]
socket_parameters = [socket_sizes]

#trimmerValue = {"100 Ohm": "100", "200 Ohm": "200", "300 Ohm": "300", "330 Ohm": "330", "470 Ohm": "470", "2K Ohm": "2K", "3K Ohm": "3K", "4.7K Ohm": "4K7", "10K Ohm": "10K", "15K Ohm": "15K", "20K Ohm": "20K", "30K Ohm": "30K", "47K Ohm": "47K", "50K Ohm": "50K", "150K Ohm": "150K", "200K Ohm": "200K", "300K Ohm": "300K", "470K Ohm": "470K", "500K Ohm": "500K", "2M Ohm": "2M"}

part_abbr = {"Enclosure": "ENC", "Switch": "SW", "IC Socket": "SOC", "Trimmer": "TRM", "Potentiometer": "POT"}
part_optionals = {"Enclosure": enclosure_optionals, "Switch": switch_optionals, "IC Socket": socket_optionals}
part_types = {"Enclosure": enclosure_parameters, "Switch": switch_parameters, "IC Socket": socket_parameters}

class OrderedPart:
    def __init__(self):
        self.part_description = ""
        self.part_number = ""
        self.supplier_part_number = ""
        self.unit_price = 0.0
        self.qty = 0
        self.backorder = 0
        self.extended_price = 0.0


def main():
    #config = configparser.ConfigParser()
    # Gather settings from settings.xml
    parse_settings_json()
    #parse_settings_xml()
    # Parse through the Order WD and create CSV's for PDF receipts as needed
    create_csvs()
    # Update MongoDB with new orders
    update_db_orders()
    print("order_wd = %s" % order_settings["order_wd"])
    print("suppliers_list:")
    print(suppliers_list)
    print("mongo_url = %s" % mongo_settings["mongo_url"])
    print("mongo_username = %s" % mongo_settings["mongo_username"])
    print("mongo_password = %s" % mongo_settings["mongo_password"])


def parse_settings_xml():
    print(os.getcwd())
    settings_path = os.getcwd() + "\\settings.xml"
    tree = etree.parse(settings_path, etree.XMLParser(ns_clean=True, recover=True, remove_blank_text=True))
    for o_settings in tree.xpath("//order_settings"):
        for setting in o_settings.getchildren():
            if setting.tag == "order_wd":
                order_settings["order_wd"] = setting.text
            elif setting.tag == "suppliers":
                for supplier in setting.getchildren():
                    if supplier.tag == "supplier":
                        suppliers_list.append(supplier.text)
    for m_settings in tree.xpath("//mongo_settings"):
        for setting in m_settings.getchildren():
            if setting.tag == "mongo_url":
                mongo_settings["mongo_url"] = setting.text
                #mongo_settings["mongo_url"] = str(html.fromstring(setting.text))
            elif setting.tag == "username":
                mongo_settings["mongo_username"] = setting.text
            elif setting.tag == "password":
                mongo_settings["mongo_password"] = setting.text


def parse_settings_json():
    settings_path = os.getcwd() + "\\settings.json"
    with open(settings_path, "r") as read_file:
        settings = json.load(read_file)
        print("settings:")
        print(settings)
        order_settings["order_wd"] = settings["order_settings"]["order_wd"]
        print("suppliers for json:")
        print(settings["order_settings"]["suppliers"])
        for supplier in settings["order_settings"]["suppliers"]:
            suppliers_list.append(supplier)
        mongo_settings["mongo_url"] = settings["mongo_settings"]["mongo_url"]
        mongo_settings["mongo_username"] = settings["mongo_settings"]["username"]
        mongo_settings["mongo_password"] = settings["mongo_settings"]["password"]
        mongo_settings["project_id"] = settings["mongo_settings"]["project_id"]


def create_csvs():
    # Walk through the order work directory
    for dir_name, dir_names, file_names in os.walk(order_settings["order_wd"]):
        for supplier in suppliers_list:
            if dir_name.endswith(supplier):
                for file_name in file_names:
                    print("file_name: %s" % file_name)
                    if file_name.endswith(".pdf"):
                        file_name_split = file_name.split(".")
                        if file_name_split[0] + ".csv" not in file_names:
                            match supplier:
                                case "Digikey":
                                    Digikey_order_parsing(os.path.join(dir_name, file_name))
                                case "Tayda":
                                    Tayda_order_csv_creator(os.path.join(dir_name, file_name))
                                case "Small Bear":
                                    Small_Bear_order_csv_creator(os.path.join(dir_name, file_name))
                                case other:
                                    print("Could not create CSV from PDF. Must be done manually. ")
                        else:
                            print("pdf found with csv conversion: %s" % os.path.join(dir_name, file_name))


def Tayda_order_csv_creator(file_name):
    # Open PDF and parse the strings in it
    page_strings = parse_pdf_strings(file_name)
    idx = 0
    #print("page_strings:")
    #print(page_strings)
    for page in page_strings:
        search_parameter = "start_of_table"
        for string in page:
            # Go through the page_strings and find each part
            match search_parameter:
                case "start_of_table":
                    # Subtotal is the last string before the 1st part's description
                    if string == "Subtotal":
                        #print("Start of Table Found")
                        search_parameter = "part_description"
                        ordered_parts.append(OrderedPart())
                case "part_description":
                    if string == "Subtotal":
                        # This indicates the end of the page, so we break, discard the new empty part, and look at the next page
                        ordered_parts.pop()
                        break
                    if string.startswith("A-"):
                        # This is the supplier part number for Tayda
                        ordered_parts[idx].supplier_part_number = string
                        search_parameter = "unit_price"
                    else:
                        # The string gathering is spotty, so this could print out with extra spaces
                        ordered_parts[idx].part_description += " " + string
                case "unit_price":
                    # Remove the $ and convert into a decimal with 2 sig figs
                    ordered_parts[idx].unit_price = decimal.Decimal(re.sub(r'[^\d.]', "", string))
                    search_parameter = "qty"
                case "qty":
                    try:
                        # Convert the quantity into an integer
                        ordered_parts[idx].qty = int(string)
                        search_parameter = "backorder"
                    except:
                        pass
                case "backorder":
                    try:
                        # Backorder = ordered quantity - shipped quantity
                        ordered_parts[idx].backorder = int(ordered_parts[idx].qty) - int(string)
                        search_parameter = "extended_price"
                    except:
                        pass
                case "extended_price":
                    # Total price of all quantity of parts ordered
                    ordered_parts[idx].extended_price = decimal.Decimal(re.sub(r'[^\d.]', "", string))
                    ordered_parts.append(OrderedPart())
                    idx += 1
                    search_parameter = "part_description"

    # Debug: print out list of part descriptions
    # TODO: Try to figure out partNumber with Description
    assign_part_number()
    print("ordered_parts:")
    print(ordered_parts)
    for part in ordered_parts:
        print("\t" + part.part_description)

    create_order_csv(file_name, ordered_parts)


def Small_Bear_order_csv_creator(file_name):
    # Open PDF and parse the strings in it
    page_strings = parse_pdf_strings(file_name)
    idx = 0
    print("page_strings:")
    print(page_strings)
    for page in page_strings:
        search_parameter = "start_of_table"
        for string in page:
            # Go through the page_strings and find each part
            match search_parameter:
                case "start_of_table":
                    # String can be broken up, so also search without the t
                    if string == "Total" or string == "otal":
                        search_parameter = "qty"
                        ordered_parts.append(OrderedPart())
                case "qty":
                    try:
                        ordered_parts[idx].qty = int(string)
                        search_parameter = "supplier_part_number"
                    except:
                        ordered_parts.pop()
                        break
                case "supplier_part_number":
                    # Sometimes the supplier p/n is split into 2 lines; trying to use length to distinguish between that and the part description
                    if len(string) <= 10:
                        if ordered_parts[idx].supplier_part_number == "":
                            ordered_parts[idx].supplier_part_number += string
                        else:
                            ordered_parts[idx].supplier_part_number += " " + string
                    else:
                        ordered_parts[idx].part_description += string
                        search_parameter = "part_description"
                case "part_description":
                    # If we reach the unit price we've got the full description
                    try:
                        # Remove the $ and convert into a decimal with 2 sig figs
                        float(string.replace("$", ""))
                        ordered_parts[idx].unit_price = decimal.Decimal(re.sub(r'[^\d.]', "", string))
                        search_parameter = "extended_price"
                    except:
                        ordered_parts[idx].part_description += " " + string
                case "extended_price":
                    # Remove the $ and convert into a decimal with 2 sig figs
                    ordered_parts[idx].extended_price = decimal.Decimal(re.sub(r'[^\d.]', "", string))
                    ordered_parts.append(OrderedPart())
                    idx += 1
                    search_parameter = "qty"

        assign_part_number()
        print("ordered_parts:")
        print(ordered_parts)
        for part in ordered_parts:
            print("\t" + part.part_description)

        create_order_csv(file_name, ordered_parts)


def Digikey_order_parsing(file_name):
    # Open PDF and parse the strings in it
    page_strings = parse_pdf_strings(file_name)
    idx = 0
    print("page_strings:")
    print(page_strings)
    for page in page_strings:
        search_parameter = "date"
        order_date = ""
        order_no = ""
        invoice_no = ""
        start_of_doc_date_found = False
        end_of_doc_date_found = False
        start_of_pack_list_found = False
        end_of_pack_list_found = False
        start_of_zip_found = False
        end_of_zip_found = False
        for string in page:
            # Go through the page string and find needed information
            match search_parameter:
                case "date":
                    if end_of_doc_date_found:
                        if string.endswith("/"):
                            order_date += string[:-1]
                            print("order_date: %s" % order_date)
                            search_parameter = "order_no"
                        else:
                            order_date += string
                    elif start_of_doc_date_found and string == ":":
                        end_of_doc_date_found = True
                    elif string == "Docu":
                        start_of_doc_date_found = True
                case "order_no":
                    if end_of_pack_list_found:
                        try:
                            int(string)
                            order_no += string
                        except:
                            print("order_no: %s" % order_no)
                            search_parameter = "invoice_no"
                    elif start_of_pack_list_found and string == ":":
                        end_of_pack_list_found = True
                    elif string == "Pack":
                        start_of_pack_list_found = True
                case "invoice_no":
                    if end_of_zip_found and string != "USA":
                        try:
                            int(string)
                            invoice_no += string
                        except:
                            print("invoice_no: %s" % invoice_no)
                            break
                    elif start_of_zip_found and string == "77":
                        end_of_zip_found = True
                    elif string == "5670":
                        start_of_zip_found = True

    return invoice_no, order_no, order_date


def parse_pdf_strings(file_name):
    # Open PDF and parse the strings in it
    fd = open(file_name, "rb")
    viewer = pdfreader.SimplePDFViewer(fd)
    page_strings = []
    for canvas in viewer:
        page_strings.append(canvas.strings)
    return page_strings


def create_order_csv(file_name, ordered_parts):
    # Create CSV for order
    file_name_split = file_name.split(".")
    csv_name = file_name_split[0] + ".csv"
    with open(csv_name, "w", newline="") as file:
        writer = csv.writer(file, delimiter=",")
        writer.writerow(csv_fields)
        idx = 1
        for part in ordered_parts:
            writer.writerow([idx, part.qty, part.supplier_part_number, "", part.part_description, "", part.backorder, part.unit_price, part.extended_price])


def assign_part_number():
    error_message = ["An issue occured when attempting to create a part number for ", ". Please enter the part number manually."] # Print this if a reference can't be found
    # Lists and dictionaries are defined at the top in the Global Variables section
    for part in ordered_parts:
        part_type_found = False
        for part_type in part_types:
            if part_type.upper() in part.part_description.upper():
                part.part_number += part_abbr[part_type]
                #print("partTypes[partType]:")
                #print(partTypes[partType])
                for idx, part_parameters in enumerate(part_types[part_type]):
                    #print("partParameters:")
                    #print(partParameters)
                    part_optional = part_optionals[part_type][idx] # Set according to the partOptionals list
                    part_found = False
                    #print("partOptional = %s" % partOptional)
                    for part_parameter in part_parameters:
                        if part_parameter.upper() in part.part_description.upper():
                            part.part_number += "-" + part_parameters[part_parameter]
                            part_found = True
                            if not part_optional:
                                break
                    if not part_optional and not part_found:
                        print("Error: part parameter not optional and not found")
                        print(error_message[0] + part.part_description + error_message[1])
            #    print("partType = %s" % partType)
                part_type_found = True
                break
        if not part_type_found:
            print(error_message[0] + part.part_description + error_message[1])
        print("Part number = %s" % part.part_number)


def update_db_orders():
    # Set up MongoDB connection
    # TODO: lxml is removing the &w in the retryWrites section of the URL. Need to figure out why and stop it from doing that.
    # TODO: Maybe change to JSON or ini?
    mongo_url = mongo_settings["mongo_url"]
    print("mongo_url = %s" % mongo_url)
    #mongo_url = "mongodb+srv://ngreenway:checksum@cluster0.7osjmnb.mongodb.net/?retryWrites=true&w=majority"
    client = MongoClient(mongo_url)
    db_names = client.list_database_names()
    print("db_names:")
    print(db_names)
    if "inventory_manager" in db_names:
        db = client.inventory_manager
    else:
        db = client["inventory_manager"]
    db_names = client.list_database_names()
    print("db_names:")
    print(db_names)
    coll_names = db.list_collection_names()
    if "orders" in coll_names:
        orders_coll = db.orders
    else:
        orders_coll = db["orders"]
    # Walk through the order work directory
    for dir_name, dir_names, file_names in os.walk(order_settings["order_wd"]):
        for supplier in suppliers_list:
            if dir_name.endswith(supplier):
                for file_name in file_names:
                    if file_name.endswith(".csv"):
                        # Find the order number in the file name
                        file_name_noext = file_name.replace(".csv", "")
                        file_name_split = file_name_noext.split(" ")
                        for word in file_name_split:
                            try:
                                int(word)
                                print("Found CSV for order number %s in %s directory." % (word, supplier))
                                #cursor = orders_coll.find({"order_number": int(word)})
                                doc_exists = orders_coll.count_documents({"order_number": int(word)})
                                if not doc_exists:
                                    orders_coll.insert_one({"order_number": int(word)})
                                break
                            except:
                                pass


if __name__ == "__main__":
    main()
