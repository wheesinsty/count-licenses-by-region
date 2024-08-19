import pandas as pd

# read users file
while True:
    input_users_filepath = input("Please enter the filepath to the users file including the file format (.csv or .xlsx): ")
    if ".csv" in input_users_filepath:
        try:
            users = pd.read_csv(input_users_filepath)
        except:
            print("Incorrect filename or format. Please double check the filepath and filetype.")
            continue
        break
    elif ".xlsx" in input_users_filepath:
        try:
            users = pd.read_excel(input_users_filepath)
        except:
            print("Incorrect filename or format. Please double check the filepath and filetype.")
            continue
        break

# read licenses file
while True:
    input_licenses_filepath = input("Please enter the filepath to the licenses file: ")
    if ".csv" in input_licenses_filepath:
        try:
            productList = pd.read_csv(input_licenses_filepath, index_col = "Product Title")
        except:
            print("Incorrect filename or format. Please double check the filepath and filetype.")
            continue
        break
    elif ".xlsx" in input_licenses_filepath:
        try:
            productList = pd.read_excel(input_licenses_filepath, index_col = "Product Title")
        except:
            print("Incorrect filename or format. Please double check the filepath and filetype.")
            continue
        break

# delete unnecessary columns
def deleteCol():
    # clear all locations
    for col in productList.columns:
        if col not in ["Total Licenses", "Expired Licenses", "Assigned Licenses", "Status Message"]:
            del productList[col]

# calculate license counts 
def calculateLicenses(): 
    for i in range(len(users)):
        location = users.loc[i, "Usage location"]
        licenses = [license for license in str(users.loc[i, "Licenses"]).split("+")]
        
        # case 1: no license
        if licenses[0] == 'nan':
            continue
        
        # case 2: no location
        if type(location) == float:
            # create new column for No Location if it doesn't exist, then fill missing values with 0 
            if "No Location" not in productList.columns:
                productList["No Location"] = 0
            # increment the No Location count
            for license in licenses:
                if license not in productList.index.tolist():
                    productList.loc[license] = 0
                productList.loc[license, location] += 1
                    
        # case 3: both location and license exist
        else:
            # create new column for the location if it doesn't exist, then fill missing values with 0
            if location not in productList.columns:
                productList[location] = 0
            # increment the Location count
            for license in licenses:
                if license not in productList.index.tolist():
                    productList.loc[license] = 0
                productList.loc[license, location] += 1

# save the final dataframe to the licenses file 
def saveFile():
    print(productList)
    productList.to_excel(input_licenses_filepath)

# main function to call the other functions
def main():
    deleteCol()
    calculateLicenses()
    saveFile()

main()