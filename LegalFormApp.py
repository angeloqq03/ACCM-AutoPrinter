from tkinter import *
from tkinter import ttk, messagebox
from PIL import Image, ImageTk, ImageWin
import cv2
import numpy as np
import io
import os
import win32print
import win32ui
import win32con
from google.oauth2 import service_account
from googleapiclient.discovery import build
from Mappings import (
    SourceCode, FirstName, MiddleName, LastName, MmDdYyyy,
    PlaceOfBirth, MotherFirstName, MotherMiddleName, MotherLastName, GenderMale, GenderFemale,
    CivilStatusSingle, CivilStatusWidowed, CivilStatusMarried, CivilStatusSeperated, DependentsNo,
    EducationHS,EducationCollege,EducationPostGrad, CitizenshipFilipino, CitizenshipOthers, Tin,
    Sss_Gsis, PresentHomeAddLine1,PresentHomeAddLine2, PermanentHomeAddLine1, PermanentHomeAddLine2,
    HomeOwnershipOwnedNotMortgaged, HomeOwnershipCompany, HomeOwnershipWithParents, YearsOfStay,
    CarsOwned, PersonalEmail, HomePhone1, HomePhone2, MobilePhone1, MobilePhone2, BestTimeToCall,
    SpouseFirstName, SpouseMiddlename, SpouseLastName, SpouseCompanyBusinessName, SpousePosition,
    BusinessAddSpouseLine1, BusinessAddSpouseLine2, BusinessPhoneSpouse, MobilePhoneSpouse, HomePhoneSpouse, AnnualIncome, EmploymentSEMP, EmploymentGOVT, EmploymentPrivate, YearsOfEmploymentOrBusinessPresent, TotalYearsOfWorking,
    PositionSeniorOff, PositionOwner, PositionDirector, PositionPresident, PositionTitle, OccupationAdmin, OccupationProfessional, OccupationSEMP, CompanyOrBusinessName, BusinessAddressLine1, BusinessAddressLine2, HrEmail, GrossAnnualIncome, SourceOfFundSalary, SourceOfFundBusiness, IsExistingClientOfEastWestYes, IsExistingClientOfEastWestNo, IfYesDeposit, IfYesCreditCard, IsExistingEWCreditCardHolderYes, IsExistingEWCreditCardHolderNo, CardIssuer, BusinessPhone, CardNumber, MemberSince, CreditLimit, Ref1FullName, Ref1CompanyOrBusinessName, Ref1BusinessPhone, Ref1HomePhone, Ref1MobilePhone, Ref2FullName, Ref2CompanyOrBusinessName, Ref2BusinessPhone, Ref2HomePhone, Ref2MobilePhone, PreferedPaymentTenor36, PreferedPaymentTenor48, PreferedPaymentTenor60, Ref2BusinessAddressLine1, Ref2BusinessAddressLine2,RefBusinessAddressLine1, RefBusinessAddressLine2
)
from Plottings import create_print_figure
from PrintControlWindow import PrintControlWindow

class LegalFormApp:
    def load_sheet_id(self):
        try:
            with open('sheet_id.txt', 'r') as f:
                sheet_id = f.read().strip()
                if not sheet_id:
                    raise ValueError("Sheet ID not found in sheet_id.txt")
                return sheet_id
        except Exception as e:
            print(f"Error loading sheet id: {str(e)}")
            raise ValueError("Sheet ID not found in sheet_id.txt")

    def update_sheet_id(self, new_sheet_id):
        """Callback to update sheet ID"""
        try:
            self.GoogleSheetID = new_sheet_id
            with open('sheet_id.txt', 'w') as f:
                f.write(new_sheet_id)
        except Exception as e:
            print(f"Error updating sheet id: {str(e)}")
            raise
        

    def __init__(self, root, front_page_path):
        self.root = root
        self.root.title("Legal Form Manager")
        self.paper_width = 8.5
        self.paper_height = 13.0
        self.printer_dpi = 344

        self.mappings = {
            'SourceCode': SourceCode,
            'FirstName': FirstName,
            'MiddleName': MiddleName,
            'LastName': LastName,
            'MmDdYyyy': MmDdYyyy,
            'PlaceOfBirth': PlaceOfBirth,
            'MotherFirstName': MotherFirstName,
            'MotherMiddleName': MotherMiddleName,
            'MotherLastName': MotherLastName,
            'GenderMale': GenderMale,
            'GenderFemale': GenderFemale,
            'CivilStatusSingle': CivilStatusSingle,
            'CivilStatusWidowed': CivilStatusWidowed,
            'CivilStatusMarried': CivilStatusMarried,
            'CivilStatusSeperated': CivilStatusSeperated,
            'DependentsNo': DependentsNo,
            'EducationHS': EducationHS,
            'EducationCollege': EducationCollege,
            'EducationPostGrad': EducationPostGrad,
            "CitizenshipFilipino" : CitizenshipFilipino,
            "CitizenshipOthers" : CitizenshipOthers,
            "Tin" : Tin,
            "Sss_Gsis" : Sss_Gsis,
            "PresentHomeAddLine1" : PresentHomeAddLine1,
            "PresentHomeAddLine2" : PresentHomeAddLine2,
            "PermanentHomeAddLine1" : PermanentHomeAddLine1,
            "PermanentHomeAddLine2" : PermanentHomeAddLine2,
            "HomeOwnershipOwnedNotMortgaged" : HomeOwnershipOwnedNotMortgaged,
            "HomeOwnershipCompany" : HomeOwnershipCompany,
            "HomeOwnershipWithParents" : HomeOwnershipWithParents,
            "YearsOfStay" : YearsOfStay,
            "CarsOwned" : CarsOwned,
            "PersonalEmail" : PersonalEmail,
            "HomePhone1" : HomePhone1,
            "HomePhone2" : HomePhone2,
            "MobilePhone1" : MobilePhone1,
            "MobilePhone2" : MobilePhone2,
            "BestTimeToCall" : BestTimeToCall,
            "SpouseFirstName" : SpouseFirstName,
            "SpouseMiddlename" : SpouseMiddlename,
            "SpouseLastName" : SpouseLastName,
            "SpouseCompanyBusinessName" : SpouseCompanyBusinessName,
            "SpousePosition" : SpousePosition,
            "BusinessAddSpouseLine1" : BusinessAddSpouseLine1,
            "BusinessAddSpouseLine2" : BusinessAddSpouseLine2,
            "BusinessPhoneSpouse" : BusinessPhoneSpouse,
            "MobilePhoneSpouse" : MobilePhoneSpouse,
            "HomePhoneSpouse" : HomePhoneSpouse,
            "AnnualIncome" : AnnualIncome,
            "EmploymentSEMP" : EmploymentSEMP,
            "EmploymentGOVT" : EmploymentGOVT,
            "EmploymentPrivate" : EmploymentPrivate,
            "YearsOfEmploymentOrBusinessPresent" : YearsOfEmploymentOrBusinessPresent,
            "TotalYearsOfWorking" : TotalYearsOfWorking,
            "PositionSeniorOff" : PositionSeniorOff,
            "PositionOwner" : PositionOwner,
            "PositionDirector" : PositionDirector,
            "PositionPresident" : PositionPresident,
            "PositionTitle" : PositionTitle,
            "OccupationAdmin" : OccupationAdmin,
            "OccupationProfessional" : OccupationProfessional,
            "OccupationSEMP" : OccupationSEMP,
            "CompanyOrBusinessName" : CompanyOrBusinessName,
            "BusinessAddressLine1" : BusinessAddressLine1,
            "BusinessAddressLine2" : BusinessAddressLine2,
            "BusinessPhone" : BusinessPhone,
            "HrEmail" : HrEmail,
            "GrossAnnualIncome" : GrossAnnualIncome,
            "SourceOfFundSalary" : SourceOfFundSalary,
            "SourceOfFundBusiness" : SourceOfFundBusiness,
            "IsExistingClientOfEastWestYes" : IsExistingClientOfEastWestYes,
            "IsExistingClientOfEastWestNo" : IsExistingClientOfEastWestNo,
            "IfYesDeposit" : IfYesDeposit,
            "IfYesCreditCard" : IfYesCreditCard,
            "IsExistingEWCreditCardHolderYes" : IsExistingEWCreditCardHolderYes,
            "IsExistingEWCreditCardHolderNo" : IsExistingEWCreditCardHolderNo,
            "CardIssuer" : CardIssuer,
            "CardNumber" : CardNumber,
            "MemberSince" : MemberSince,
            "CreditLimit" : CreditLimit,

            "Ref1FullName" : Ref1FullName,
            "Ref1CompanyOrBusinessName" : Ref1CompanyOrBusinessName,
            "RefBusinessAddressLine1" : RefBusinessAddressLine1,
            "RefBusinessAddressLine2" : RefBusinessAddressLine2,
            "Ref1BusinessPhone" : Ref1BusinessPhone,
            "Ref1HomePhone" : Ref1HomePhone,
            "Ref1MobilePhone" : Ref1MobilePhone,

            "Ref2FullName" :Ref2FullName ,
            "Ref2CompanyOrBusinessName" : Ref2CompanyOrBusinessName,
            "Ref2BusinessAddressLine1" : Ref2BusinessAddressLine1,
            "Ref2BusinessAddressLine2" : Ref2BusinessAddressLine2,
            "Ref2BusinessPhone" : Ref2BusinessPhone,
            "Ref2HomePhone" : Ref2HomePhone,
            "Ref2MobilePhone" : Ref2MobilePhone,

            "PreferedPaymentTenor36" : PreferedPaymentTenor36,
            "PreferedPaymentTenor48" : PreferedPaymentTenor48,
            "PreferedPaymentTenor60" : PreferedPaymentTenor60,
        }

        self.paper_dims = {
            'width': self.paper_width,
            'height': self.paper_height
        }
        self.GoogleSheetID = self.load_sheet_id()
        self.SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
        self.SERVICE_ACCOUNT_FILE = 'organic.json'
        self.client_data = {}


        try:
            self.load_image(front_page_path)
            self.print_window = PrintControlWindow(
                root,
                self.image,
                self.mappings,
                self.paper_dims,
                fetch_callback=self.fetch_single_client,
                print_callback=self.print_selected_clients,
                sheet_id_callback=self.update_sheet_id
            )
        except Exception as e:
            print(f"Initialization error: {str(e)}")
            raise

    def load_image(self, front_page_path):
        try:
            self.image_path = front_page_path
            self.image = cv2.imread(self.image_path)
            if self.image is None:
                raise ValueError(f"Failed to load image: {self.image_path}")
            self.image = cv2.cvtColor(self.image, cv2.COLOR_BGR2RGB)
        except Exception as e:
            print(f"Error loading image: {str(e)}")
            self.print_window.display_message(f"Error loading image: {str(e)}", "error")
            raise

    def fetch_single_client(self, client_no):
        """Fetch data for a single client"""
        try:
            if not os.path.exists(self.SERVICE_ACCOUNT_FILE):
                raise FileNotFoundError(f"Credentials file not found: {self.SERVICE_ACCOUNT_FILE}")
                
            creds = service_account.Credentials.from_service_account_file(
                self.SERVICE_ACCOUNT_FILE, scopes=self.SCOPES)
            
            service = build('sheets', 'v4', credentials=creds)
            sheet = service.spreadsheets()
            
            result = sheet.values().get(
                spreadsheetId=self.GoogleSheetID,
                range="'DETAILS FOR APP FORM'!A:ZZ"
            ).execute()
            
            values = result.get('values', [])
            if not values:
                return None
            
            headers = values[0]
            
            for row in values[1:]:
                if row and row[0].strip().upper() == client_no.strip().upper():
                    date_str = row[6] if len(row) > 6 else ''
                    date_parts = date_str.split('/') if date_str else []
                    formatted_date = ''.join(date_parts) if date_parts else ''
                    
                    client_data = {
                        'source_code': list(row[2].strip()) if len(row) > 2 else [],
                        'first_name': row[3] if len(row) > 3 else '',
                        'middle_name': row[4] if len(row) > 4 else '',
                        'last_name': row[5] if len(row) > 5 else '',
                        'date_of_birth': formatted_date,
                        'place_of_birth': row[7] if len(row) > 7 else '',
                        'mother_first_name': row[8] if len(row) > 8 else '',
                        'mother_middle_name': row[9] if len(row) > 9 else '',
                        'mother_last_name': row[10] if len(row) > 10 else '',
                        'gender': row[11] if len(row) > 11 else '',
                        'civil_status': row[12] if len(row) > 12 else '',
                        'dependents_no': row[13] if len(row) > 13 else '',
                        'education': row[14] if len(row) > 14 else '',
                        'citizenship': row[15] if len(row) > 15 else '',
                        'tin' : row[16] if len(row) > 16 else '',
                        'sss_gsis' : row[17] if len(row) > 17 else '',
                        'present_home_address' : row[18] if len(row) > 18 else '',
                        'permanent_home_address' : row[19] if len(row) > 19 else '',
                        'home_ownership' : row[20] if len(row) > 20 else '',
                        'years_of_stay' : row[21] if len(row) > 21 else '',
                        'cars_owned' : row[22] if len(row) > 22 else '',
                        'personal_email' : row[23] if len(row) > 23 else '',
                        'home_phone_1' : row[24] if len(row) > 24 else '',
                        'home_phone_2' : row[25] if len(row) > 25 else '',
                        'mobile_phone_1' : row[26] if len(row) > 26 else '',
                        'mobile_phone_2' : row[27] if len(row) > 27 else '',
                        'best_time_to_call' : row[28] if len(row) > 28 else '',
                        'name_of_spouse_first' : row[29] if len(row) > 29 else '',
                        'name_of_spouse_middle' : row[30] if len(row) > 30 else '',
                        'name_of_spouse_last' : row[31] if len(row) > 31 else '',
                        'company_business_name_spouse' : row[32] if len(row) > 32 else '',
                        'position_spouse' : row[33] if len(row) > 33 else '',
                        'business_address_spouse' : row[34] if len(row) > 34 else '',
                        'business_phone_spouse' : row[35] if len(row) > 35 else '',
                        'mobile_phone_spouse' : row[36] if len(row) > 36 else '',
                        'home_phone_spouse' : row[37] if len(row) > 37 else '',
                        'employment' : row[41] if len(row) > 41 else '',
                        'years_w_present' : row[42] if len(row) > 42 else '',
                        'total_years_working' : row[43] if len(row) > 43 else '',
                        'position' : row[44] if len(row) > 44 else '',
                        'position_title' : row[45] if len(row) > 45 else '',
                        'occupation' : row[47] if len(row) > 47 else '',
                        'company_business_name' : row[48] if len(row) > 48 else '',
                        'business_address' : row[49] if len(row) > 49 else '',
                        'business_phone' : row[50] if len(row) > 50 else '',
                        'business_hr_email' : row[51] if len(row) > 51 else '',
                        'gross_annual_income' : row[52] if len(row) > 52 else '',
                        'source_of_funds' : row[53] if len(row) > 53 else '',
                        'is_existing_ew_client' : row[54] if len(row) > 54 else '',
                        'if_is_existing_ew_client_yes' : row[55] if len(row) > 55 else '',
                        'is_existing_ew_cc_holder' : row[56] if len(row) > 56 else '',
                        'card_issuer' : row[60] if len(row) > 60 else '',
                        'card_number' : row[61] if len(row) > 61 else '',
                        'member_since'  : row[62] if len(row) > 62 else '',
                        'credit_limit' : row[63] if len(row) > 63 else '',
                        'ref1_full_name' : row[64] if len(row) > 64 else '',
                        'ref1_company_business_name' : row[65] if len(row) > 65 else '',
                        'ref1_comp_business_add' : row[66] if len(row) > 66 else '',
                        'ref1_business_phone' : row[67] if len(row) > 67 else '',
                        'ref1_home_phone' : row[68] if len(row) > 68 else '',
                        'ref1_mobile_phone' : row[69] if len(row) > 69 else '',
                        'ref2_full_name' : row[70] if len(row) > 70 else '',
                        'ref2_company_business_name' : row[71] if len(row) > 71 else '',
                        'ref2_comp_business_add' : row[72] if len(row) > 72 else '',
                        'ref2_business_phone' : row[73] if len(row) > 73 else '',
                        'ref2_home_phone' : row[74] if len(row) > 74 else '',
                        'ref2_mobile_phone' : row[75] if len(row) > 75 else '',
                        'prefered_payment_tenor' : row[78] if len(row) > 78 else '',
                    }
                    
                    self.client_data[client_no] = client_data
                    return client_data
            
            return None
            
        except Exception as e:
            print(f"Fetch error: {str(e)}")
            self.print_window.display_message(f"Fetch error: {str(e)}", "error")
            raise

    def print_selected_clients(self, client_numbers):
        """Print forms for selected clients"""
        if not client_numbers:
            self.print_window.display_message("No clients selected", "warning")
            return
            
        try:
            printer_name = win32print.GetDefaultPrinter()
            
            for client_no in client_numbers:
                if client_no not in self.client_data:
                    if not self.fetch_single_client(client_no):
                        continue
                
                client = self.client_data[client_no]
                dc = win32ui.CreateDC()
                dc.CreatePrinterDC(printer_name)
                
                dc.StartDoc(f'Legal Form - Client {client_no}')
                dc.StartPage()
                
                fig = create_print_figure(
                    self.image,
                    client,
                    self.mappings,
                    self.paper_dims
                )
                
                buf = io.BytesIO()
                fig.savefig(buf, format='png', dpi=344, bbox_inches='tight', pad_inches=0)
                buf.seek(0)

                img = Image.open(buf)
                dib = ImageWin.Dib(img)

                width = int(self.paper_width * self.printer_dpi)
                height = int(self.paper_height * self.printer_dpi)
                dib.draw(dc.GetHandleOutput(), (0, 0, width, height))
                
                dc.EndPage()
                dc.EndDoc()
                dc.DeleteDC()
            
            self.print_window.display_message(f"Printed {len(client_numbers)} forms at 344 DPI", "success")

            
        except Exception as e:
            print(f"Print error: {str(e)}")
            self.print_window.display_message(f"Print error: {str(e)}", "error")