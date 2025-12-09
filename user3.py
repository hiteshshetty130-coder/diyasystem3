import requests
import zipfile
import pandas as pd
import io
import logging

logging.basicConfig(level=logging.INFO)

def main():
    retry=3

    for attempt in range(1,retry+1):
        try:
            url="https://www.thespreadsheetguru.com/wp-content/uploads/2022/12/EmployeeSampleData.zip"
            response=requests.get(url)
            file_path="output.xlsx"

            zipe_file_path=io.BytesIO(response.content)
            with zipfile.ZipFile(zipe_file_path,"r") as z:
                try:
                    excel_file=next(f for f in z.namelist() if f.endswith(".xlsx"))   
                    extracted_data=z.open(excel_file)
                    df=pd.read_excel(extracted_data)

                except Exception as e:
                    logging.exception("Corrupted File!!") 


                if not excel_file.endswith(".xlsx"):
                    logging.error("Unsupported File Format!")

            df.to_excel(file_path,index=False)
            if not file_path.endswith(".xlsx"):
                logging.error("Incorrect File Type!!")
            
            logging.info("\n%s",df)

            return df,"output.xlsx"
            break

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
