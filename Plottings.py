from matplotlib.figure import Figure
from matplotlib.patches import Rectangle

def plot_text(ax, text, coordinates, font_size):
    """Plot text character by character at given coordinates"""
    for i, char in enumerate(text):
        if i < len(coordinates):
            x, y = list(coordinates)[i]
            ax.text(x, y, char,
                   fontsize=font_size, color="black",
                   ha='center', va='center',
                   fontweight='bold', fontname='Arial')

def plot_long_text(ax, text, coordinate, font_size):
    """Plot entire text at a single coordinate point"""
    if coordinate:
        x, y = list(coordinate.values())[0]
        ax.text(x, y, text,
                fontsize=font_size, color="black",
                ha='left', va='center',
                fontweight='bold', fontname='Arial')

def plot_exclusive_checkboxes(ax, value, options_map, mappings):
    """Plot mutually exclusive checkboxes where only one can be checked"""
    if not value:  
        for option, mapping_key in options_map.items():
            if mapping_key in mappings:
                plot_checkbox(ax, False, mappings.get(mapping_key, {}))
        return
    value = str(value).strip().upper()
    print(f"Processing value: {value}")  
    for option, mapping_key in options_map.items():
        option = str(option).strip().upper()
        is_checked = value == option
        print(f"Checking {option}: {is_checked}")  
        
        if mapping_key in mappings:
            plot_checkbox(ax, is_checked, mappings.get(mapping_key, {}))

def plot_checkbox(ax, is_checked, coordinate):
    """Plot a checkbox with or without check mark"""
    if coordinate:
        x, y = list(coordinate.values())[0]
        box_size = 20
        rect = Rectangle((x-box_size/2, y-box_size/2), box_size, box_size,
                        facecolor='white', edgecolor='white')
        ax.add_patch(rect)
        if is_checked:
            ax.text(x, y-2, "/", 
                   fontsize=9,
                   color="black",
                   ha='center',
                   va='center',
                   rotation=0)

def create_print_figure(image, client_data, mappings, paper_dims):
    """Create figure for printing"""
    fig = Figure(figsize=(paper_dims['width'], paper_dims['height']), dpi=344)
    ax = fig.add_subplot(111)
    ax.imshow(image)
    
    font_size = 6 * (paper_dims['height'] / 14.0)
    small_font = 4 * (paper_dims['height'] / 14.0)
    large_font = 6 * (paper_dims['height'] / 14.0)
    
    try:
        coords_source = list(mappings['SourceCode'].values())
        ax.scatter(*zip(*coords_source), color="red", s=30, marker="o")
        plot_text(ax, client_data['source_code'], coords_source, font_size)
        
        plot_text(ax, client_data['first_name'], mappings['FirstName'].values(), font_size)
        plot_text(ax, client_data['middle_name'], mappings['MiddleName'].values(), font_size)
        plot_text(ax, client_data['last_name'], mappings['LastName'].values(), font_size)
        plot_text(ax, client_data['date_of_birth'], mappings['MmDdYyyy'].values(), font_size)
        
        plot_long_text(ax, client_data.get('place_of_birth', ''), mappings.get('PlaceOfBirth', {}), small_font)
        plot_long_text(ax, client_data.get('mother_first_name', ''), mappings.get('MotherFirstName', {}), small_font)
        plot_long_text(ax, client_data.get('mother_middle_name', ''), mappings.get('MotherMiddleName', {}), small_font)
        plot_long_text(ax, client_data.get('mother_last_name', ''), mappings.get('MotherLastName', {}), small_font)
        
        gender_options = {
            'MALE': 'GenderMale',
            'FEMALE': 'GenderFemale'
        }
        plot_exclusive_checkboxes(ax, client_data.get('gender', ''), gender_options, mappings)
        
        civil_status_options = {
            'SINGLE': 'CivilStatusSingle',
            'MARRIED': 'CivilStatusMarried',
            'WIDOWED': 'CivilStatusWidowed',
            'SEPARATED': 'CivilStatusSeparated'
        }

        
        plot_exclusive_checkboxes(ax, client_data.get('civil_status', ''), civil_status_options, mappings)
        plot_long_text(ax, client_data.get('dependents_no', ''), mappings.get('DependentsNo', {}), large_font)

        education_options = {
            'HIGH SCHOOL': 'EducationHS',
            'COLLEGE': 'EducationCollege',
            'POST-GRADUATE': 'EducationPostGrad'
        }
        plot_exclusive_checkboxes(ax, client_data.get('education', ''), education_options, mappings)

        citizenship_options = {
            'FILIPINO': 'CitizenshipFilipino',
            'OTHERS': 'CitizenshipOthers'
        }
        plot_exclusive_checkboxes(ax, client_data.get('citizenship', ''), citizenship_options, mappings)
        plot_long_text(ax, client_data.get('tin', ''), mappings.get('Tin', {}), large_font)
        plot_long_text(ax, client_data.get('sss_gsis', ''), mappings.get('Sss_Gsis', {}), large_font)
        
        present_address = client_data.get('present_home_address', '')
        if present_address:
            words = present_address.split()
            if len(words) >= 6:
                mid = len(words) // 2
                line1 = ' '.join(words[:mid])
                line2 = ' '.join(words[mid:])
            else:
                line1 = present_address
                line2 = ""
            plot_long_text(ax, line1, mappings.get('PresentHomeAddLine1', {}), small_font)
            if line2:
                plot_long_text(ax, line2, mappings.get('PresentHomeAddLine2', {}), small_font)
        
        permanent_address = client_data.get('permanent_home_address', '')
        if permanent_address:
            words = permanent_address.split()
            if len(words) >= 6:
                mid = len(words) // 2
                line1 = ' '.join(words[:mid])
                line2 = ' '.join(words[mid:])
            else:
                line1 = permanent_address
                line2 = ""
            plot_long_text(ax, line1, mappings.get('PermanentHomeAddLine1', {}), small_font)
            if line2:
                plot_long_text(ax, line2, mappings.get('PermanentHomeAddLine2', {}), small_font)
        

        homeownerrship_options = {
            'OWNED, NOT MORTGAGED': 'HomeOwnershipOwnedNotMortgaged',
            'COMPANY PROVIDED': 'HomeOwnershipCompany',
            'LIVING W/ PARENTS / RELATIVES': 'HomeOwnershipWithParents'
        }
        plot_exclusive_checkboxes(ax, client_data.get('home_ownership', ''), homeownerrship_options, mappings)
        plot_long_text(ax, client_data.get('years_of_stay', ''), mappings.get('YearsOfStay', {}), small_font)

        # cars owned
        plot_long_text(ax, client_data.get('cars_owned', ''), mappings.get('CarsOwned', {}), small_font)

        plot_long_text(ax, client_data.get('personal_email', ''), mappings.get('PersonalEmail', {}), small_font)

        plot_long_text(ax, client_data.get('home_phone_1', ''), mappings.get('HomePhone1', {}), small_font)
        plot_long_text(ax, client_data.get('home_phone_2', ''), mappings.get('HomePhone2', {}), small_font)

        plot_long_text(ax, client_data.get('mobile_phone_1', ''), mappings.get('MobilePhone1', {}), small_font)
        plot_long_text(ax, client_data.get('mobile_phone_2', ''), mappings.get('MobilePhone2', {}), small_font)


        plot_long_text(ax, client_data.get('best_time_to_call', ''), mappings.get('BestTimeToCall', {}), small_font)


        plot_long_text(ax, client_data.get('name_of_spouse_first', ''), mappings.get('SpouseFirstName', {}), small_font)
        plot_long_text(ax, client_data.get('name_of_spouse_middle', ''), mappings.get('SpouseMiddlename', {}), small_font)
        plot_long_text(ax, client_data.get('name_of_spouse_last', ''), mappings.get('SpouseLastName', {}), small_font)

        plot_long_text(ax, client_data.get('company_business_name_spouse', ''), mappings.get('SpouseCompanyBusinessName', {}), small_font)
        plot_long_text(ax, client_data.get('position_spouse', ''), mappings.get('SpousePosition', {}), small_font)
        plot_long_text(ax, client_data.get('position_spouse', ''), mappings.get('SpousePosition', {}), small_font)
        
        spouse_business_add = client_data.get('business_address_spouse', '')
        if spouse_business_add:
            words = spouse_business_add.split()
            if len(words) >= 6:
                mid = len(words) // 2
                line1 = ' '.join(words[:mid])
                line2 = ' '.join(words[mid:])
            else:
                line1 = spouse_business_add
                line2 = ""
            plot_long_text(ax, line1, mappings.get('BusinessAddSpouseLine1', {}), small_font)
            if line2:
                plot_long_text(ax, line2, mappings.get('BusinessAddSpouseLine2', {}), small_font)

        plot_long_text(ax, client_data.get('business_phone_spouse', ''), mappings.get('BusinessPhoneSpouse', {}), small_font)
        plot_long_text(ax, client_data.get('mobile_phone_spouse', ''), mappings.get('MobilePhoneSpouse', {}), small_font)
        plot_long_text(ax, client_data.get('home_phone_spouse', ''), mappings.get('HomePhoneSpouse', {}), small_font)

        employment_options = {
            'SELF-EMPLOYED': 'EmploymentSEMP',
            'GOVERNMENT': 'EmploymentGOVT',
            'PRIVATE': 'EmploymentPrivate',
        }
        plot_exclusive_checkboxes(ax, client_data.get('employment', ''), employment_options, mappings)
        plot_long_text(ax, client_data.get('years_w_present', ''), mappings.get('YearsOfEmploymentOrBusinessPresent', {}), small_font)
        plot_long_text(ax, client_data.get('total_years_working', ''), mappings.get('TotalYearsOfWorking', {}), small_font)

        positions_options = {
            'SENIOR OFFICER': 'PositionSeniorOff',
            'OWNER / PART-OWNER' : 'PositionOwner',
            'DIRECTOR' : 'PositionDirector',
            'PRESIDENT' : 'PositionPresident',
        }
        plot_exclusive_checkboxes(ax, client_data.get('position', ''), positions_options, mappings)
        
        plot_long_text(ax, client_data.get('position_title', ''), mappings.get('PositionTitle', {}), small_font)
        
        occupation_options = {
            'ADMINISTRATOR / EXECUTIVE': 'OccupationAdmin',
            'PROFESSIONAL': 'OccupationProfessional',
            'SELF-EMPLOYED': 'OccupationSEMP',
        }
        plot_exclusive_checkboxes(ax, client_data.get('occupation', ''), occupation_options, mappings)
        

        plot_long_text(ax, client_data.get('company_business_name', ''), mappings.get('CompanyOrBusinessName', {}), small_font)


        business_address = client_data.get('business_address', '')
        if business_address:
            words = business_address.split()
            if len(words) >= 6:
                mid = len(words) // 2
                line1 = ' '.join(words[:mid])
                line2 = ' '.join(words[mid:])
            else:
                line1 = business_address
                line2 = ""
            plot_long_text(ax, line1, mappings.get('BusinessAddressLine1', {}), small_font)
            if line2:
                plot_long_text(ax, line2, mappings.get('BusinessAddressLine2', {}), small_font)


        plot_long_text(ax, client_data.get('business_phone', ''), mappings.get('BusinessPhone', {}), small_font)
        plot_long_text(ax, client_data.get('business_hr_email', ''), mappings.get('HrEmail', {}), small_font)
        plot_long_text(ax, client_data.get('gross_annual_income', ''), mappings.get('GrossAnnualIncome', {}), small_font)


        source_of_fund_option = {
            'SALARY / BENEFITS': 'SourceOfFundSalary',
            'BUSINESS INCOME': 'SourceOfFundBusiness',
        }

        plot_exclusive_checkboxes(ax, client_data.get('source_of_funds', ''), source_of_fund_option, mappings)

        existing_client_options = {
            'YES': 'IsExistingClientOfEastWestYes',
            'NO': 'IsExistingClientOfEastWestYes',
        }
        plot_exclusive_checkboxes(ax, client_data.get('is_existing_ew_client', ''), existing_client_options, mappings)

        if client_data.get('is_existing_ew_client') == 'YES':
            deposit_or_credit_options = {
                'DEPOSIT': 'IfYesDeposit',
                'CREDIT CARD': 'IfYesCreditCard',
            }
            plot_exclusive_checkboxes(ax, client_data.get('if_is_existing_ew_client_yes', ''), deposit_or_credit_options, mappings)  

        existing_credit_card_options = {
            'YES': 'IsExistingEWCreditCardHolderYes',
            'NO': 'IsExistingEWCreditCardHolderNo',
            }
        
        plot_exclusive_checkboxes(ax, client_data.get('is_existing_ew_cc_holder', ''), existing_credit_card_options, mappings)

        plot_long_text(ax, client_data.get('card_issuer', ''), mappings.get('CardIssuer', {}), small_font)
        plot_long_text(ax, client_data.get('card_number', ''), mappings.get('CardNumber', {}), small_font)
        plot_long_text(ax, client_data.get('member_since', ''), mappings.get('MemberSince', {}), small_font)
        plot_long_text(ax, client_data.get('credit_limit', ''), mappings.get('CreditLimit', {}), small_font)

        plot_long_text(ax, client_data.get('ref1_full_name', ''), mappings.get('Ref1FullName', {}), large_font)
        plot_long_text(ax, client_data.get('ref1_company_business_name', ''), mappings.get('Ref1CompanyOrBusinessName', {}), small_font)


        refbusiness_address1 = client_data.get('ref1_comp_business_add', '')
        if refbusiness_address1:
            words = refbusiness_address1.split()
            if len(words) >= 6:
                mid = len(words) // 2
                line1 = ' '.join(words[:mid])
                line2 = ' '.join(words[mid:])
            else:
                line1 = refbusiness_address1
                line2 = ""
            plot_long_text(ax, line1, mappings.get('RefBusinessAddressLine1', {}), small_font)
            if line2:
                plot_long_text(ax, line2, mappings.get('BusinessAddressLine2', {}), small_font)

            plot_long_text(ax, client_data.get('ref1_business_phone', ''), mappings.get('Ref1BusinessPhone', {}), small_font)
            plot_long_text(ax, client_data.get('ref1_home_phone', ''), mappings.get('Ref1HomePhone', {}), small_font)
            plot_long_text(ax, client_data.get('ref1_mobile_phone', ''), mappings.get('Ref1MobilePhone', {}), small_font)

            plot_long_text(ax, client_data.get('ref2_full_name', ''), mappings.get('Ref2FullName', {}), small_font)
            plot_long_text(ax, client_data.get('ref2_company_business_name', ''), mappings.get('Ref2CompanyOrBusinessName', {}), small_font)

            refbusiness_address2 = client_data.get('ref2_comp_business_add', '')
        if refbusiness_address2:
            words = refbusiness_address2.split()
            if len(words) >= 6:
                mid = len(words) // 2
                line1 = ' '.join(words[:mid])
                line2 = ' '.join(words[mid:])
            else:
                line1 = refbusiness_address2
                line2 = ""
            plot_long_text(ax, line1, mappings.get('Ref2BusinessAddressLine1', {}), small_font)
            if line2:
                plot_long_text(ax, line2, mappings.get('Ref2BusinessAddressLine2', {}), small_font)

        plot_long_text(ax, client_data.get('ref2_business_phone', ''), mappings.get('Ref2BusinessPhone', {}), small_font)
        plot_long_text(ax, client_data.get('ref2_home_phone', ''), mappings.get('Ref2HomePhone', {}), small_font)
        plot_long_text(ax, client_data.get('ref2_mobile_phone', ''), mappings.get('Ref2MobilePhone', {}), small_font)

        payment_tenor_options = {
            '36 months (3 YEARS)': 'PreferedPaymentTenor36',
            '48 months (4 YEARS)': 'PreferedPaymentTenor48',
            '60 months (5 YEARS)': 'PreferedPaymentTenor60',
        }
        plot_exclusive_checkboxes(ax, client_data.get('prefered_payment_tenor', ''), payment_tenor_options, mappings)

        



            
    except Exception as e:
        print(f"Error plotting fields: {str(e)}")
        print(f"Available mappings: {list(mappings.keys())}")
        raise
    
    ax.axis('off')
    return fig