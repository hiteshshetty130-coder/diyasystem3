import requests
import zipfile
import pandas as pd
import io
import logging

logging.basicConfig(level=logging.INFO)

def main():
    retry=3

    #loop Till the maximum attempts Reaches 
    for attempt in range(1,retry+1):
        try:
            url="https://www.thespreadsheetguru.com/wp-content/uploads/2022/12/EmployeeSampleData.zip" 
            response=requests.get(url)
            file_path="output.xlsx"

            #IO.bytesIO is used to handle binary data returned from the url
            zipe_file_path=io.BytesIO(response.content)
            with zipfile.ZipFile(zipe_file_path,"r") as z: #using zipefile module open in read mode all the files inside 
                try:
                    #fetch the excel file
                    excel_file=next(f for f in z.namelist() if f.endswith(".xlsx"))   
                    extracted_data=z.open(excel_file)
                    df=pd.read_excel(extracted_data)        #create a dataframe from the file extracted

                except Exception as e: #if file cannot be extracted then log exception
                    logging.exception("Corrupted File!!") 

                #check for supported file format 
                if not excel_file.endswith(".xlsx"):
                    logging.error("Unsupported File Format!")

            df.to_excel(file_path,index=False) #convert to excel

            if not file_path.endswith(".xlsx"):
                logging.error("Incorrect File Type!!")
            
            logging.info("\n%s",df)

            return df,"output.xlsx"
            break


        #exception handling for invalid ZIP File,Connection problem,and any unexcepted error handling
        except zipfile.BadZipFile:
            logging.exception("Sorry Inavlid Zip file!!")

        except requests.ConnectionError:
            logging.exception("sorry no Internet")

        except Exception as e:
            logging.exception(f"Unexpected Error Occurred:{e}")

        if attempt==retry:
            logging.warning("MAXIMUM ATTEMPT REACHED")
        else:
            logging.info("RETRYING")


if __name__=="__main__":
    main()
