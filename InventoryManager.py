import pdfreader
import re
import decimal
import csv

csvFields = ["Index", "Quantity", "Part Number", "Manufacturer Part Number", "Description", "Customer Reference", "Backorder", "Unit Price", "Extended Price"]

def main():
    fileName = "L:\\KiCad\\~Parts Inventory\\Order History\\Tayda\\Order # 1000369536.pdf"
    TaydaOrderCSVCreator(fileName)

class OrderedPart:
    def __init__(self):
        self.partDescription = ""
        self.partNumber = ""
        self.supplierPartNumber = ""
        self.unitPrice = 0.0
        self.qty = 0
        self.backorder = 0
        self.extendedPrice = 0.0

def TaydaOrderCSVCreator(fileName):
    # Open PDF and parse the strings in it
    fd = open(fileName, "rb")
    viewer = pdfreader.SimplePDFViewer(fd)
    page_strings = []
    for canvas in viewer:
        page_strings.append(canvas.strings)
    orderedParts = []
    idx = 0
    for page in page_strings:
        searchParameter = "startOfTable"
        for string in page:
            # Go through the page_strings and find each part
            match searchParameter:
                case "startOfTable":
                    # Subtotal is the last string before the 1st part's description
                    if string == "Subtotal":
                        print("Start of Table Found")
                        searchParameter = "partDescription"
                        orderedParts.append(OrderedPart())
                case "partDescription":
                    if string == "Subtotal":
                        # This indicates the end of the page, so we break and look at the next page
                        orderedParts.pop()
                        break
                    if string.startswith("A-"):
                        # This is the supplier part number for Tayda
                        orderedParts[idx].supplierPartNumber = string
                        searchParameter = "unitPrice"
                    else:
                        # The string gathering is spotty, so this could print out with extra spaces
                        orderedParts[idx].partDescription += " " + string
                case "unitPrice":
                    # Remove the $ and convert into a decimal with 2 sig figs
                    orderedParts[idx].unitPrice = decimal.Decimal(re.sub(r'[^\d.]', "", string))
                    searchParameter = "qty"
                case "qty":
                    try:
                        # Convert the quantity into an integer
                        orderedParts[idx].qty = int(string)
                        searchParameter = "backorder"
                    except:
                        pass
                case "backorder":
                    try:
                        # Backorder = ordered quantity - shipped quantity
                        orderedParts[idx].backorder = int(orderedParts[idx].qty) - int(string)
                        searchParameter = "extendedPrice"
                    except:
                        pass
                case "extendedPrice":
                    # Total price of all quantity of parts ordered
                    orderedParts[idx].extendedPrice = decimal.Decimal(re.sub(r'[^\d.]', "", string))
                    orderedParts.append(OrderedPart())
                    idx += 1
                    searchParameter = "partDescription"

    # Debug: print out list of part descriptions
    # TODO: Try to figure out partNumber with Description
    print("orderedParts:")
    for part in orderedParts:
        print("\t" + part.partDescription)

    # Create CSV for order
    fileName_split = fileName.split(".")
    csvName = fileName_split[0] + ".csv"
    with open(csvName, "w", newline="") as file:
        writer = csv.writer(file, delimiter=",")
        writer.writerow(csvFields)
        idx = 1
        for part in orderedParts:
            writer.writerow([idx, part.qty, part.supplierPartNumber, "", part.partDescription, "", part.backorder, part.unitPrice, part.extendedPrice])


if __name__ == "__main__":
    main()
