import google_sheet as GoogleSheet
import argparse

class Sample:
    def __init__(self,args):
        self.args = args
    
    def main(self):
        google_sheet = GoogleSheet.GoogleSheet(args=self.args)
        result = google_sheet.insert_column(worksheet=1,number=1,column=0,values=["SAMPLE NAME 1","SAMPLE NAME 2","SAMPLE NAME 3","SAMPLE NAME 4","SAMPLE NAME 5"],inherit=False)
        print(result)
        

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Google Sheet")
    parser.add_argument("--credentials", type=str, help="Path to the credentials file.",default='Service_credentials.json')
    parser.add_argument("--spreadsheet_id", type=str, help="Id of the spreadsheet.",default='1N022ciXOUj6Cc3gU0bo2aGbjw88Quob2yJnYnW-A4Es')
    parser.add_argument("--worksheet", type=str, help="Name of the worksheet.",default='Sheet1')
    
    
    arg = parser.parse_args()
    s = Sample(arg)
    s.main()
        
