import os
import openpyxl
import xlsxwriter
import win32com.client


class Excel:

    def excel_automation(self, excel_file_path):


        # Open the Excel file
        wb = openpyxl.load_workbook(excel_file_path)
        sheet = wb.active

        # Check if there are rows to copy
        if sheet.max_row > 0:
            # Copy the first row (header)
            first_row = []
            for cell in sheet[1]:
                first_row.append(cell.value)

            # Remove the first row from the sheet
            sheet.delete_rows(1)

            # Save the modified Excel file
            wb.save(excel_file_path)

            # Convert the copied row to a string
            copied_user_name = ' '.join(map(str, first_row))
            print(copied_user_name)
            return copied_user_name
        else:
            print("The Excel file is empty.")
            return None


class Photoshop(Excel):

    def change_and_save_photoshop(self, psd_file_path, text_to_insert):
        if text_to_insert is None:
            return

        try:
            psApp = win32com.client.Dispatch("Photoshop.Application")
            psApp.Open(psd_file_path)

            doc = psApp.Application.ActiveDocument
            layerText = doc.ArtLayers["Bla"]
            text_of_layer = layerText.TextItem
            text_of_layer.contents = text_to_insert

            if text_to_insert:
                options = win32com.client.Dispatch("Photoshop.ExportOptionsSaveForWeb")
                options.Format = 6
                options.Quality = 100
                jpgfile = f"D:/Text_Files/output_projects/{text_to_insert}.jpg"
                doc.Export(ExportIn=jpgfile, ExportAs=2, Options=options)
                print(f"File saved as {text_to_insert}")

        except Exception as e:
            print("Error occurred", str(e))


if __name__ == "__main__":
    excel = Excel()
    photoshop = Photoshop()
    psd_file_path = "D:/Text_Files/ronaldo.psd"
    excel_file_path = 'D:/Text_Files/emo.xlsx'

    while True:
        copied_user_name = excel.excel_automation(excel_file_path)
        if copied_user_name == "None":
            break

        photoshop.change_and_save_photoshop(psd_file_path, copied_user_name)
