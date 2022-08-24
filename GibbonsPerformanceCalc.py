
import pandas as pd
import streamlit as st
import datetime as dt

#Retrieves Data from Excel File
def getData(file):
    matrixFile = pd.ExcelFile(file)
    df = pd.read_excel(matrixFile, None)
    return df
    
#Allows User to Upload File
def fileUpload(id):
    #Setting the results to 0 before actual assignment because otherwise streamlit throws a temper tantrum
    df = 0
    uploadedFile = st.file_uploader("Select File...", key=id)
    if uploadedFile is not None:
        df = getData(uploadedFile)
    return df

#Allows User to Select Months
def monthSelect():
    monthsList = ["all months", "january", "february", "march", "april", "may", "june", "july", 
    "august", "september", "october", "november", "december"]
    st.header("Select Dates")
    st.subheader("Months To Select")
    chosenMonths = st.multiselect("Select Months", monthsList)
    return chosenMonths

#Allows User to Select Years
def yearSelect():
    currentDate = dt.date.today()
    currentYear = currentDate.year
    yearsList = list(range(2021, currentYear + 1))
    st.subheader("Years to Select")
    chosenYears = st.multiselect("Select Years", yearsList)
    return chosenYears

def lowerSheets(sheet):
    #placeholder to not give error
    newSheet = 0
    try:
        newSheet = dict((k.lower(), v) for k, v in sheet.items())
    except:
        pass
    return newSheet

#Front-End Streamlit Display; Returns User Selections to Calculation Function
def display():
    #Top titles
    st.title("JC Gibbons Performance Report")
    st.header("Upload Files")
    st.subheader("Upload On-Time-Late Matrix")
    #File Uploading & Validation For Sheets
    matrixDF = fileUpload("mat")
    st.subheader("Upload Monthly Quality Postings")
    matrixDF = lowerSheets(matrixDF)
    postingDF = fileUpload("post")
    #User-choice Of Date
    chosenMonths = monthSelect()
    chosenYear = yearSelect()
    #User-choice of Weights
    shippingWeight = st.number_input("Select Weight Percentage of Late Parts", min_value =0, max_value=100, value=50)
    PPMWeight = st.number_input("Select Weight Percentage of PPM", min_value=0, max_value=100, value=50)
    return matrixDF, postingDF, chosenMonths, chosenYear, shippingWeight, PPMWeight

def chooseMatrixSheets(chosenMonths, chosenYear):
    #Creating empty list of sheets
    sheets = []
    #Setting month list if all months are chosen
    try:
        if chosenMonths[0] == "all months":
            months = ["january", "february", "march", "april", "may", "june", "july", 
            "august", "september", "october", "november", "december"]
        else:
            months = chosenMonths
        #Finding number of years, looping for that number of years to create same format as sheet
        yearAmount = len(chosenYear)
        for i in range(yearAmount):
            for q in range(len(months)):
                combination = (" ".join([months[q], str(chosenYear[i])]))
                sheets.append(combination)
        return sheets
    except:
        pass

def calculations(matrixSheets, postingSheet, matrixDF, postingDF, shippingWeight, PPMWeight):
    #Going through each sheet and doing the math, storing each one in a calculation dataframe
    calculationDataFrame = pd.DataFrame({"Department":[],"Date":[],"Parts Ordered":[],"Parts Late":[],
    "Percent Late":[],"PPM":[]})#"Performance":[]})
    try:
        for i in range(len(matrixSheets)):
            try:
                #Cleaning the quantity column
                currentMatrixSheet = matrixDF[matrixSheets[i]].astype(str)
                # for testing st.dataframe(currentMatrixSheet)
                quantity = currentMatrixSheet["Quantity Ordered"]
                quantity = quantity.where(quantity != "nan").values
                quantity = pd.to_numeric(quantity)
                # for testing st.dataframe(quantity)
                currentMatrixSheet["Quantity Ordered"] = quantity
                currentMatrixSheet.dropna(inplace=True)
                #for testing st.dataframe(currentMatrixSheet)

                #Finding Late Parts
                #Removing NA
                lateNoNA = currentMatrixSheet[(currentMatrixSheet["Days Late from original"] != "na") & (currentMatrixSheet["Days Late from original"] != "nan")]
                lateClean = lateNoNA["Days Late from original"]
                #Converting Strings to Int, Finding Late Parts & Summing
                lateConvert = pd.to_numeric(lateClean)
                DAVLatePartsList = currentMatrixSheet[(lateConvert.astype(int) > 3) & (currentMatrixSheet["Department"] == "DAV")]
                CNCLatePartsList = currentMatrixSheet[(lateConvert.astype(int) > 3) & (currentMatrixSheet["Department"] == "CNC")]
                DAVlatePartsList = pd.to_numeric(DAVLatePartsList["Quantity Ordered"])
                CNClatePartsList = pd.to_numeric(CNCLatePartsList["Quantity Ordered"])
                sumDAVlate = DAVlatePartsList.sum()
                sumCNClate = CNClatePartsList.sum()

                #Summing Total In-House Parts
                DAVPartsList = currentMatrixSheet[currentMatrixSheet["Department"] == "DAV"]
                CNCPartsList = currentMatrixSheet[currentMatrixSheet["Department"] == "CNC"]
                DAVQuantity = DAVPartsList["Quantity Ordered"]
                CNCQuantity = CNCPartsList["Quantity Ordered"]
                DAVQuantity = DAVQuantity.astype(int)
                CNCQuantity = CNCQuantity.astype(int)
                DAVSum = DAVQuantity.sum()
                CNCSum = CNCQuantity.sum()

                #Calculating Percent Late
                DAVpercentageLate = (sumDAVlate / DAVSum) * 100
                CNCpercentageLate = (sumCNClate / CNCSum) * 100

                #Finding PPM
                #Loading in Data
                ppmDataFrame = postingDF[postingSheet]
                cols = ppmDataFrame.columns
                #Choosing only dates
                dates = [date for date in cols if not isinstance(date, str)]
                #Splitting DAV and CNC
                DAVppm = ppmDataFrame.iloc[:5]
                CNCppm = ppmDataFrame.iloc[5:]
                #Getting Rid of Other columns in each table
                DAVppm = DAVppm[dates]
                CNCppm = CNCppm[dates]
                #Putting dates in matching format
                datesFinal = []
                for q in range(len(dates)):
                    datesNewFormat = dt.datetime.strftime(dates[q], "%B %Y")
                    datesFinal.append(datesNewFormat.lower())
                DAVppm.columns = datesFinal
                CNCppm.columns = datesFinal
                #Selecting resulting ppm (matrixSheets is the date & iloc 3 is the row of PPM)
                DAVppmFinal = DAVppm[matrixSheets[i]].iloc[3]
                CNCppmFinal = CNCppm[matrixSheets[i]].iloc[3]
            
                #Calculting Performance Score
                DAVperformance = (((abs(DAVpercentageLate - 100))/100)*shippingWeight) + (((abs(DAVppmFinal-50000))/50000)*PPMWeight)
                CNCperformance = (((abs(CNCpercentageLate - 100))/100)*shippingWeight) + (((abs(CNCppmFinal-50000))/50000)*PPMWeight)
                DAVperformance = DAVperformance.round(1)
                CNCperformance = CNCperformance.round(1)
                #Creating temp dataframe to add onto new one
                tempFrame = pd.DataFrame({"Department":["DAV", "CNC"],"Date":[matrixSheets[i],matrixSheets[i]],
                "Parts Ordered":[DAVSum,CNCSum],"Parts Late":[sumDAVlate,sumCNClate],
                "Percent Late":[DAVpercentageLate, CNCpercentageLate],"PPM":[DAVppmFinal, CNCppmFinal], "Performance":[DAVperformance, CNCperformance]})

                #Appending temp dataframe to overall
                calculationDataFrame = calculationDataFrame.append(tempFrame)
                
                

            except:
                pass
    except:
        pass

    #Converting Date to Datetime so it sorts properly
    backToDTList = []
    for z in range(len(calculationDataFrame["Date"])):
                backToDT = dt.datetime.strptime(calculationDataFrame["Date"].iloc[z], "%B %Y")
                backToDTList.append(backToDT)
    calculationDataFrame["Date"] = backToDTList
    #Displaying Data
    st.header("Raw Data")
    st.subheader("Sorting Option")
    #Sorting Data based on user decision
    sortOption = st.selectbox("Sorting Method", calculationDataFrame.columns)
    calculationDataFrame = calculationDataFrame.sort_values(by=sortOption)
    #Formatting Nicely
    calculationDataFrame = calculationDataFrame.style.format({"Date":"{:%m.%d.%Y}","Parts Late":"{:,.0f}","PPM":"{:.0f}","Parts Ordered":"{:,.0f}",
    "Percent Late":"{:.2f}", "Performance":"{:.2f}"})
    st.dataframe(calculationDataFrame)

def main():
    #try:
    matrixDF, postingDF, chosenMonths, chosenYear, shippingWeight, PPMWeight = display()
    matrixSheets = chooseMatrixSheets(chosenMonths, chosenYear)
    postingSheet = "PPM"
    calculations(matrixSheets, postingSheet, matrixDF, postingDF, shippingWeight, PPMWeight)
    #except:
        #pass 

main()