import unittest
from io import BytesIO
from user3 import main
import pandas as pd
from unittest.mock import Mock,patch
import zipfile
import os
from openpyxl import Workbook

class TestExcelFileExtraction(unittest.TestCase):
    #Verify Excel file Download
    @patch("user3.requests.get")
    def test_case1(self,mock_get):
        
        wb=Workbook()
        ws=wb.active
        ws['a1']="hello"
        ws['a2']="hitesh"
        ws['b1']=100
        ws["b2"]=200

        excel_bytes=BytesIO()
        wb.save(excel_bytes)

        zip_bytes=BytesIO()
        with zipfile.ZipFile(zip_bytes,"w") as z_f:
            z_f.writestr("output_mock.xlsx",excel_bytes.getvalue())

        mock_response=Mock()
        mock_response.status_code=200
        mock_response.content=zip_bytes.getvalue()
        mock_get.return_value=mock_response

        df,file_path=main()

        self.assertIsNotNone(df)
        self.assertTrue(os.path.exists(file_path))
        
    @patch("user3.requests.get")
    def test_case2(self,mock_get):
        wb=Workbook()
        ws=wb.active
        ws['a1']="hello"
        ws['a2']="hitesh"
        ws['b1']=100
        ws['b2']=200

        excel_bytes=BytesIO()
        wb.save(excel_bytes)

        zip_bytes=BytesIO()
        with zipfile.ZipFile(zip_bytes,"w") as z_f:
            z_f.writestr("output_mock.xlsx",excel_bytes.getvalue())

        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.content = zip_bytes.getvalue()
        mock_get.return_value = mock_response

        df, file_path = main()

        self.assertTrue(file_path.endswith(".xlsx"))
        self.assertGreater(len(df.columns),0)
        self.assertGreater(len(df),0)

    @patch("user3.requests.get")
    def test_case3(self, mock_get):
        wb = Workbook()
        ws = wb.active
        ws['a1'] = "hello"
        ws['a2'] = "hitesh"
        ws['b1'] = 100
        ws["b2"] = 200

        excel_bytes = BytesIO()
        wb.save(excel_bytes)

        zip_bytes = BytesIO()
        with zipfile.ZipFile(zip_bytes, "w") as z_f:
            z_f.writestr("output_mock.xlsx", excel_bytes.getvalue())

        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.content = zip_bytes.getvalue()
        mock_get.return_value = mock_response

        df, file_path = main()

        self.assertTrue(file_path.endswith(".xlsx"))
        self.assertTrue(os.path.exists(file_path))
        self.assertIsInstance(df,pd.DataFrame)

    @patch("user3.requests.get")
    def test_case4(self, mock_get):
        df_fake=pd.DataFrame({"Employee ID":["E1011"],"Full Name":["Hitesh"],"Job Title":["Manager"],"Department":["IT"],
                              "Business Unit":["xYZ"],"Gender":["male"],"Ethnicity":["asian"],"Age":[19],"Hire Date":["19-12-2005"],
                              "Annual Salary":[20000],"Bonus":[5000],"Country":["india"],"City":["mangalore"],
                              "Exit Date":["19-12-2024"]})

        excel_bytes = BytesIO()
        df_fake.to_excel(excel_bytes)
        excel_bytes.seek(0)
    

        zip_bytes = BytesIO()
        with zipfile.ZipFile(zip_bytes, "w") as z_f:
            z_f.writestr("output_mock.xlsx", excel_bytes.getvalue())
        zip_bytes.seek(0)

        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.content = zip_bytes.getvalue()
        mock_get.return_value = mock_response

        df, file_path = main()

        expected_columns=["Employee ID","Full Name","Job Title","Hire Date"]

        for col in expected_columns:
            self.assertIn(col,df.columns)
    
    @patch("user3.requests.get")
    def test_case5(self, mock_get):
        df_fake = pd.DataFrame(
            {"Employee ID": ["E1011"], "Full Name": ["Hitesh"], "Job Title": ["Manager"], "Department": ["IT"],
             "Business Unit": ["xYZ"], "Gender": ["male"], "Ethnicity": ["asian"], "Age": [19],
             "Hire Date": ["19-12-2005"],
             "Annual Salary": [20000], "Bonus": [5000], "Country": ["india"], "City": ["mangalore"],
             "Exit Date": ["19-12-2024"]})

        excel_bytes = BytesIO()
        df_fake.to_excel(excel_bytes)
        excel_bytes.seek(0)

        zip_bytes = BytesIO()
        with zipfile.ZipFile(zip_bytes, "w") as z_f:
            z_f.writestr("output_mock.xlsx", excel_bytes.getvalue())
        zip_bytes.seek(0)

        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.content = zip_bytes.getvalue()
        mock_get.return_value = mock_response

        df, file_path = main()

        expected_columns = ["Employee ID", "Full Name", "Job Title", "Hire Date"]

        for col in expected_columns:
            self.assertIn(col, df.columns,"columns is missing")

        self.assertFalse(df.isnull().any().any() ,"no null values Exists")

if __name__=="__main__":
    unittest.main()