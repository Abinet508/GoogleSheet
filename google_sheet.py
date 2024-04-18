import pygsheets
import pandas as pd
import os
import click
import argparse

class GoogleSheet:
    """Google Sheet class to read, write, update and delete data from Google Sheet.
    """
    def __init__(self, args):
        self.args = args
        for root, dirs, files in os.walk(os.path.dirname(os.path.abspath(__file__))):
            if self.args.credentials in files:
                self.args.credentials = os.path.join(root, self.args.credentials)
                break
            
        if not os.path.exists(self.args.credentials):
            if not os.path.exists(self.args.credentials):
                self.args.credentials = click.prompt("Enter path to the credentials file", type=str)
                self.args.credentials = os.path.join(os.path.dirname(os.path.abspath(__file__)), "credentials", self.args.credentials)
                if not os.path.exists(self.args.credentials):
                    self.args.credentials = os.path.join(os.path.dirname(os.path.abspath(__file__)), "credentials", "wise-logic-415907-d45cefa39f1d.json")

        self.gc = pygsheets.authorize(service_file=self.args.credentials)
        if self.args.spreadsheet_id:
            self.spreadsheet = self.gc.open_by_key(self.args.spreadsheet_id)
        else:
            raise ValueError("Please provide either spreadsheet id.")
        if self.args.worksheet:
            self.worksheet = self.spreadsheet.worksheet_by_title(self.args.worksheet)
            self.worksheet = self.worksheet if isinstance(self.worksheet, pygsheets.Worksheet) else self.spreadsheet.sheet1
        else:
            self.worksheet = self.spreadsheet.sheet1 if isinstance(self.spreadsheet.sheet1, pygsheets.Worksheet) else self.create_new_worksheet(self.args.worksheet)
        
    def read_data_from_google_sheet_as_df(self):
        """
        Read data from Google Sheet as DataFrame.

        Returns:
        pd.DataFrame: DataFrame containing data from Google Sheet.
        """
        df = self.worksheet.get_as_df()
        df = df.dropna()
        df = df.reset_index(drop=True)
        return df
    
    def read_data_from_google_sheet_as_list(self):
        """
        Read data from Google Sheet as list.

        Returns:
        list: List containing data from Google Sheet.
        """
        df_remove_empty = self.worksheet.get_as_df()
        df_remove_empty = df_remove_empty.dropna()
        get_all_value = df_remove_empty.values.tolist()
        return get_all_value
    
    def read_data_from_google_sheet_as_dict(self):
        """
        Read data from Google Sheet as dictionary.

        Returns:
        dict: Dictionary containing data from Google Sheet.
        """
        df = self.worksheet.get_as_df()
        df = df.dropna()
        df = df.reset_index(drop=True)
        return df.to_dict()
    
    def create_new_worksheet(self, title, rows=100, cols=26, src_tuple=None,src_worksheet=None):
        """
        Create new worksheet in the spreadsheet.

        Args:
        title (str): Title of the new worksheet.
        """
        return self.spreadsheet.add_worksheet(title, rows, cols, src_tuple, src_worksheet)
    
    def delete_worksheet(self, worksheet):
        """
        Delete worksheet from the spreadsheet.

        Args:
        worksheet (pygsheets.Worksheet): Worksheet to be deleted.
        """
        self.spreadsheet.del_worksheet(worksheet)
        
    def create_new_spreadsheet(self, title, template=None, folder=None, folder_name=None):
        """
        Create new spreadsheet.

        Args:
        title (str): Title of the new spreadsheet.
        """
        return self.gc.create(title, template, folder, folder_name)
    
    def delete_spreadsheet(self, spreadsheet):
        """
        Delete spreadsheet.

        Args:
        spreadsheet (pygsheets.Spreadsheet): Spreadsheet to be deleted.
        """
        if isinstance(spreadsheet, pygsheets.Spreadsheet):
            spreadsheet.delete()
    
    def share_spreadsheet(self, email, role='reader',type ="anyone", spreadsheet=None, emailMessage="Sharing this spreadsheet with you."):
        """
        Share spreadsheet with user.

        Args:
        spreadsheet (pygsheets.Spreadsheet): Spreadsheet to be shared.
        email (str): Email of the user.
        role (str): Role of the user.
        """
        if spreadsheet:
            if isinstance(spreadsheet, pygsheets.Spreadsheet):
                if email:
                    spreadsheet.share(email, role, "user", emailMessage)
                if type == "anyone":
                    spreadsheet.share("", role, type)
        else:
            if email:
                self.spreadsheet.share(email, role, "user", emailMessage)
            if type == "anyone":
                self.spreadsheet.share("", role, type)
        
    def create_local_copy(self, path, type='excel'):
        """
        Create local copy of the spreadsheet.

        Args:
        path (str): Path of the local copy.
        """
        df = self.read_data_from_google_sheet_as_df()
        if isinstance(df, pd.DataFrame):
            if type == 'excel':
                df.to_excel(path, index=False)
            elif type == 'csv':
                df.to_csv(path, index=False)
            else:
                df.to_json(path, orient='records', indent=4)
        else:
            raise ValueError("Data is not in DataFrame format.")
    
    def read_local_copy(self, path, type='excel'):
        """
        Read local copy of the spreadsheet.

        Args:
        path (str): Path of the local copy.
        """
        if type == 'excel':
            return pd.read_excel(path)
        elif type == 'csv':
            return pd.read_csv(path)
        else:
            return pd.read_json(path)
        
    def update_title(self, spreadsheet, title):
        """
        Update title of the spreadsheet.

        Args:
        title (str): New title of the spreadsheet.
        """
        if isinstance(spreadsheet, pygsheets.Spreadsheet):
            spreadsheet.title = title
            return spreadsheet
        else:
            raise ValueError("Please provide a valid spreadsheet.")
        
    def update_worksheet_title(self, worksheet, title):
        """
        Update title of the worksheet.

        Args:
        worksheet (pygsheets.Worksheet): Worksheet to be updated.
        title (str): New title of the worksheet.
        """
        if isinstance(worksheet, pygsheets.Worksheet):
            worksheet.title = title
            return worksheet
        else:
            raise ValueError("Please provide a valid worksheet.")
        
    def update_worksheet_dimensions(self, worksheet, rows, cols):
        """
        Update dimensions of the worksheet.

        Args:
        worksheet (pygsheets.Worksheet): Worksheet to be updated.
        rows (int): Number of rows in the worksheet.
        cols (int): Number of columns in the worksheet.
        """
        if isinstance(worksheet, pygsheets.Worksheet):
            worksheet.rows = rows
            worksheet.cols = cols
            return worksheet
        else:
            raise ValueError("Please provide a valid worksheet.")
        
    def update_worksheet_frozen_rows(self, worksheet, rows):
        """
        Update frozen rows of the worksheet.

        Args:
        worksheet (pygsheets.Worksheet): Worksheet to be updated.
        rows (int): Number of frozen rows in the worksheet.
        """
        if isinstance(worksheet, pygsheets.Worksheet):
            worksheet.frozen_rows = rows
            return worksheet
        else:
            raise ValueError("Please provide a valid worksheet.")

    def update_worksheet_frozen_cols(self, worksheet, cols):
        """
        Update frozen columns of the worksheet.

        Args:
        worksheet (pygsheets.Worksheet): Worksheet to be updated.
        cols (int): Number of frozen columns in the worksheet.
        """
        if isinstance(worksheet, pygsheets.Worksheet):
            worksheet.frozen_cols = cols
            return worksheet
        else:
            raise ValueError("Please provide a valid worksheet.")
    
    def write_data_to_google_sheet(self, df, start,worksheet = None, copy_index=False, copy_head=True, extend=False, fit=False, escape_formulae=False, **kwargs):
        """
        Write data to Google Sheet.

        Args:
        df (pd.DataFrame): DataFrame to be written.
        start (str): Start cell of the data.
        copy_index (bool): Copy index of the DataFrame.
        copy_head (bool): Copy header of the DataFrame.
        extend (bool): Extend the worksheet if required.
        fit (bool): Fit the worksheet if required.
        escape_formulae (bool): Escape formulae in the data.
        """
        if worksheet:
            if isinstance(worksheet, pygsheets.Worksheet):
                worksheet.set_dataframe(df, start, copy_index, copy_head, extend, fit, escape_formulae, **kwargs)
            else:
                self.worksheet.set_dataframe(df, start, copy_index, copy_head, extend, fit, escape_formulae, **kwargs)
    
    def write_data_to_google_sheet_as_list(self, data, start,worksheet = None, extend=False, fit=False, escape_formulae=False, **kwargs):
        """
        Write data to Google Sheet as list.

        Args:
        data (list): List to be written.
        start (str): Start cell of the data.
        extend (bool): Extend the worksheet if required.
        fit (bool): Fit the worksheet if required.
        escape_formulae (bool): Escape formulae in the data.
        """
        if worksheet:
            if isinstance(worksheet, pygsheets.Worksheet):
                worksheet.update_values(start, data, extend, fit, escape_formulae, **kwargs)
            else:
                self.worksheet.update_values(start, data, extend, fit, escape_formulae, **kwargs)
    
    def write_data_to_google_sheet_from_dict(self, data, start,worksheet = None, extend=False, fit=False, escape_formulae=False, **kwargs):
        """
        Write data to Google Sheet from dictionary.

        Args:
        data (dict): Dictionary to be written.
        start (str): Start cell of the data.
        extend (bool): Extend the worksheet if required.
        fit (bool): Fit the worksheet if required.
        escape_formulae (bool): Escape formulae in the data.
        """
        if worksheet:
            if isinstance(worksheet, pygsheets.Worksheet):
                df = pd.DataFrame(data)
                worksheet.set_dataframe(df, start, copy_index=False, copy_head=True, extend=extend, fit=fit, escape_formulae=escape_formulae, **kwargs)
            else:
                df = pd.DataFrame(data)
                self.worksheet.set_dataframe(df, start, copy_index=False, copy_head=True, extend=extend, fit=fit, escape_formulae=escape_formulae, **kwargs)
    
    def update_cell(self, worksheet=None, cell="A1", value=""):
        """
        Update cell in the worksheet.

        Args:
        worksheet (pygsheets.Worksheet): Worksheet to be updated.
        cell (str): Cell to be updated.
        value (str): Value to be updated.
        """
        if isinstance(worksheet, pygsheets.Worksheet):
            worksheet.update_value(cell, value)
        else:
            self.worksheet.update_value(cell, value)
    
    def insert_row(self, worksheet=None, row=1, number=1, inherit=True, values=[]):
        """
        Insert row in the worksheet.

        Args:
        worksheet (pygsheets.Worksheet): Worksheet to be updated.
        row (int): Row to be inserted.
        number (int): Number of rows to be inserted.
        inherit (bool): Inherit values from the previous row. when true, variable row value shouldn't be 0.
        values (list): Values to be inserted.
        """
        if isinstance(worksheet, pygsheets.Worksheet):
            worksheet.insert_rows(row=row,number=number,values=values, inherit=inherit)
        else:
            self.worksheet.insert_rows(row=row,number=number,values=values, inherit=inherit)
            
    def insert_column(self, worksheet =None, column = 1, number = 1, inherit = True, values = []):
        """
        Insert column in the worksheet.

        Args:
        worksheet (pygsheets.Worksheet): Worksheet to be updated.
        column (int): Column to be inserted.
        number (int): Number of columns to be inserted.
        inherit (bool): Inherit values from the previous column. when true, variable column value shouldn't be 0.
        values (list): Values to be inserted.
        """
        if isinstance(worksheet, pygsheets.Worksheet):
            worksheet.insert_cols(col=column, number =number, values=values, inherit=inherit)
        else:
            self.worksheet.insert_cols(col=column, number =number, values=values, inherit=inherit)
            
    def update_range(self, worksheet=None, range=None, values=[], extend=False, parse=True):
        """
        Update range in the worksheet.

        Args:
        worksheet (pygsheets.Worksheet): Worksheet to be updated.
        range (str): Range to be updated.
        values (list): Values to be updated.
        extend (bool): Extend the worksheet if required.
        parse (bool): Parse the values to google sheet format.
        """
        if isinstance(worksheet, pygsheets.Worksheet):
            worksheet.update_values(crange =range, values=values, extend=extend, parse=parse,majordim='ROWS')
        else:
            self.worksheet.update_values(crange =range, values=values, extend=extend, parse=parse,majordim='ROWS')
    
    def update_row(self, worksheet, index, col_offset = 0, values = []):
        """
        Update row in the worksheet.

        Args:
        worksheet (pygsheets.Worksheet): Worksheet to be updated.
        index (int): Index of the starting row to be updated.
        col_offset (int): Columns to skip before updating values.
        values (list):  Values to be updated as matrix.
        """
        if isinstance(worksheet, pygsheets.Worksheet):
            worksheet.update_row(index, values, col_offset)
        else:
            self.worksheet.update_row(index, values, col_offset)
    
    def update_column(self, worksheet, index, values, row_offset = 0):
        """
        Update column in the worksheet.

        Args:
        worksheet (pygsheets.Worksheet): Worksheet to be updated.
        index (int): Index of the starting column to be updated.
        values (list): Values to be updated.
        row_offset (int): rows to skip before inserting values.
        """
        if isinstance(worksheet, pygsheets.Worksheet):
            worksheet.update_col(index, values, row_offset)
        else:
            self.worksheet.update_col(index, values, row_offset)
    
    def update_cells(self, worksheet, cells, values):
        """
        Update cells in the worksheet.

        Args:
        worksheet (pygsheets.Worksheet): Worksheet to be updated.
        cells (list): Cells to be updated.
        values (list): Values to be updated.
        """
        if isinstance(worksheet, pygsheets.Worksheet):
            worksheet.update_cells(cells, values)
        else:
            self.worksheet.update_cells(cells, values)
    
    def remove_rows(self, worksheet, start, end):
        """
        Remove rows from the worksheet.

        Args:
        worksheet (pygsheets.Worksheet): Worksheet to be updated.
        start (int): Start row to be removed.
        end (int): End row to be removed.
        """
        if isinstance(worksheet, pygsheets.Worksheet):
            worksheet.delete_rows(start, end)
        else:
            self.worksheet.delete_rows(start, end)
    
    def remove_columns(self, worksheet, start, end):
        """
        Remove columns from the worksheet.

        Args:
        worksheet (pygsheets.Worksheet): Worksheet to be updated.
        start (int): Start column to be removed.
        end (int): End column to be removed.
        """
        if isinstance(worksheet, pygsheets.Worksheet):
            worksheet.delete_cols(start, end)
        else:
            self.worksheet.delete_cols(start, end)
    
    def clear_worksheet(self, worksheet):
        """
        Clear worksheet.

        Args:
        worksheet (pygsheets.Worksheet): Worksheet to be cleared.
        """
        if isinstance(worksheet, pygsheets.Worksheet):
            worksheet.clear()
        else:
            self.worksheet.clear()
    
    def clear_by_field(self, worksheet, field):
        """
        Clear range in the worksheet.

        Args:
        worksheet (pygsheets.Worksheet): Worksheet to be cleared.
        range (str): Range to be cleared.
        """
        if isinstance(worksheet, pygsheets.Worksheet):
            worksheet.clear(fields=field)
        else:
            self.worksheet.clear(fields=field)           
                
    def clear_by_range(self, worksheet, start, end):
        """
        Clear range in the worksheet.

        Args:
        worksheet (pygsheets.Worksheet): Worksheet to be cleared.
        range (str): Range to be cleared.
        """
        if isinstance(worksheet, pygsheets.Worksheet):
            worksheet.clear(start, end)
        else:
            self.worksheet.clear(start, end)
    
    
    """
    Usage of GoogleSheet class.
    
    if __name__ == "__main__":
    
        parser = argparse.ArgumentParser(description="Google Sheet")
        parser.add_argument("--credentials", type=str, help="Path to the credentials file.")
        parser.add_argument("--spreadsheet", type=str, help="Name of the spreadsheet.")
        parser.add_argument("--spreadsheet_id", type=str, help="Id of the spreadsheet.")
        parser.add_argument("--worksheet", type=str, help="Name of the worksheet.")
        args = parser.parse_args()
        google_sheet = GoogleSheet(args)
    
        # Create GoogleSheet object.
        google_sheet = GoogleSheet(args)
        
        # Read data from Google Sheet as DataFrame.
        df = google_sheet.read_data_from_google_sheet_as_df()
        
        # Read data from Google Sheet as list.
        data = google_sheet.read_data_from_google_sheet_as_list()
        
        # Read data from Google Sheet as dictionary.
        data = google_sheet.read_data_from_google_sheet_as_dict()
        
        # Create new worksheet in the spreadsheet.
        worksheet = google_sheet.create_new_worksheet(title)

    """        
            
        
    
        
    
        
        
        