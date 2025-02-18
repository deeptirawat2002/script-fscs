# Load required libraries
import pandas as pd
import re
from pathlib import Path
import os
import pycountry

# Load main FSCS CSV file
file_path = "fscs_scv_tables.xlsx"
data_inputs_df = pd.read_excel(file_path, sheet_name="Data inputs")

# Define validation functions
def is_alphanumeric(value):
    # Modified to accept parentheses, periods, and additional characters
    return bool(re.fullmatch(r"[A-Za-z0-9 '\-\(\)\.,]+", str(value))) if pd.notna(value) else True

def is_alpha(value):
    return bool(re.fullmatch(r"[A-Za-z '-]+", str(value))) if pd.notna(value) else True

def is_numeric(value):
    try:
        # Convert scientific notation or any number format to string and remove any spaces
        if isinstance(value, (int, float)):
            cleaned_value = f"{value:.0f}".strip()
        else:
            cleaned_value = str(value).strip().replace(" ", "")
        return bool(re.match(r'^[0-9]+$', cleaned_value)) if pd.notna(value) else True
    except:
        return False

def is_decimal(value):
    return bool(re.fullmatch(r'\d+\.\d+', str(value))) if pd.notna(value) else True

def is_valid_date(value):
    if pd.isna(value):
        return True
    try:
        date_str = str(value).strip()
        if isinstance(value, (int, float)):
            date_str = f"{int(value)}"
        
        if len(date_str) == 7:
            date_str = "0" + date_str
        elif len(date_str) < 7 or len(date_str) > 8:
            return False
            
        if len(date_str) != 8:
            return False
            
        day = int(date_str[:2])
        month = int(date_str[2:4])
        year = int(date_str[4:])
        
        return (1 <= day <= 31) and (1 <= month <= 12) and (1900 <= year <= 2099)
    except:
        return False

def is_valid_email(value):
    email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(email_pattern, str(value))) if pd.notna(value) else True


def is_valid_phone_number(value):
    """Validate phone number format including international prefix and scientific notation"""
    if pd.isna(value):
        return True
    
    try:
        # Handle scientific notation by converting to integer first
        if isinstance(value, (int, float)):
            value = f"{int(value)}"
        
        # Clean the string value
        clean_number = str(value).strip().replace(' ', '')
        
        # Allow +, 00 prefix and only digits
        phone_pattern = r'^(\+|00)?[0-9]{1,15}$'
        return bool(re.match(phone_pattern, clean_number))
    except:
        return False
    
def is_valid_iban(value):
    """Validate IBAN format"""
    if pd.isna(value):
        return True
    iban_pattern = r'^[A-Z]{2}[0-9]{2}[A-Z0-9]{11,30}$'
    clean_iban = str(value).upper().replace(' ', '')
    return bool(re.match(iban_pattern, clean_iban))

def is_valid_bic(value):
    """Validate BIC/SWIFT code format"""
    if pd.isna(value):
        return True
    bic_pattern = r'^[A-Z]{6}[A-Z2-9][A-NP-Z0-9]([A-Z0-9]{3})?$'
    return bool(re.match(bic_pattern, str(value).upper()))

def check_stp_eligibility(value):
    """Check for non-STP eligibility indicators"""
    if pd.isna(value):
        return True
    non_stp_keywords = {'DECEASED', 'TRUST', 'FUND', "DEC'D", 'C/O', 'STOP'}
    return not any(keyword in str(value).upper() for keyword in non_stp_keywords)

def is_short_name(value):
    """Check if name is too short (less than 3 characters)"""
    if pd.isna(value):
        return False
    return len(str(value).strip()) < 3

def contains_only_initials(value):
    """Check if name contains only initials"""
    if pd.isna(value):
        return False
    name_parts = str(value).split()
    return all(len(part.strip('.')) == 1 for part in name_parts)

def validate_ascii_range(value):
    """Validate that all characters are within ASCII 32-127 range"""
    if pd.isna(value):
        return True
    return all(32 <= ord(c) <= 127 for c in str(value))

validation_functions = {
    'AlphaNumeric': is_alphanumeric,
    'Alpha': is_alpha,
    'Numeric': is_numeric,
    'Decimal': is_decimal,
    'Email': is_valid_email,
    'Phone': is_valid_phone_number,
    'IBAN': is_valid_iban,
    'BIC': is_valid_bic,
    'ASCII': validate_ascii_range
}

def validate_file(file_path, rules_df):
    filename = Path(file_path).name
    is_exclusion_file = 'EX' in filename
    
    new_xls = pd.ExcelFile(file_path)
    new_data_df = new_xls.parse(new_xls.sheet_names[0])
    
    new_data_df['Exclusion_File'] = 'Yes' if is_exclusion_file else ''

    
    formatted_output = []
    seen_values = set()
    seen_account_numbers = set()
    seen_scv_records = set()
    
    valid_product_types = {'IAA', 'ISA', 'NA', 'FD1', 'FD2', 'FD4','FP4P', 'Other'}
    valid_exclusion_types = {'HMTS', 'LEGDIS', 'LEGDOR', 'BEN'}

    for index, row in new_data_df.iterrows():
        data_row = row.copy()
        validation_row = []
        
        individual_status = "Individual" if pd.notna(row.get("title")) else ""
        
        for col_name in row.index:
            if col_name in rules_df["Name in File"].values and col_name in new_data_df.columns:
                rule = rules_df[rules_df["Name in File"] == col_name].iloc[0]
                max_length = int(rule["Max Number of Characters"]) if pd.notna(rule["Max Number of Characters"]) else None
                data_type = rule["Type of data"]
                mandatory = rule["Mandate or not"] == "Yes"

                value = row[col_name]
                str_value = str(value) if pd.notna(value) else ""
                errors = []
                
                if col_name == 'sort_code' and pd.notna(value):
                    try:
                        if isinstance(value,(int,float)):
                            cleaned_value = f'{value:.0f}'.strip()
                        else:
                            cleaned_value = str(value).strip().replace(" ", "")
                        

                    except:
                        errors.append("Invalid Sort Code Format")
                    if not re.match(r'^[0-9]+$', cleaned_value):
                        errors.append("Invalid Sort Code Format")
                

                # Conditional mandatory checks
                if col_name == 'customer_second_forename':
                    if pd.notna(row.get('customer_third_forename')) and pd.isna(value):
                        errors.append("Mandatory when third forename is present")

                if col_name.startswith('address_line_'):
                    line_num = int(col_name[-1])
                    next_lines = [f'address_line_{i}' for i in range(line_num + 1, 7)]
                    if any(pd.notna(row.get(next_line)) for next_line in next_lines) and pd.isna(value):
                        errors.append(f"Mandatory when address line {line_num + 1} or higher is populated")                

                    # Care of address check
                    if line_num == 1 and 'C/O' in str_value.upper():
                        errors.append("Care of Address - NFFSTP")

                # Product type validation
                if col_name == 'product_type' and pd.notna(value):
                    if str(value) not in valid_product_types:
                        errors.append("Invalid product type")

                # Account branch jurisdiction validation
                if col_name == "account_branch_jurisdiction" and pd.notna(value):
                    if str(value).upper() not in ["GBR", "GIB"]:
                        errors.append("Invalid branch jurisdiction - Must be GBR or GIB")

                # Currency validation
               


                # Special handling for compensatable_amount
                if col_name == 'compensatable_amount':
                    exclusion_type = row.get('exclusion_type')
                    if mandatory and (not pd.notna(exclusion_type) or str(exclusion_type).upper() != 'BEN'):
                        if pd.isna(value):
                            errors.append("Mandatory unless exclusion_type is BEN")

                # Special handling for bank_recovery_and_resolution_marking
                if col_name == 'bank_recovery_and_resolution_marking':
                    if not is_exclusion_file:
                        if pd.isna(value):
                            errors.apped("Mandatory for non-exclusive files")
                    if pd.notna(value) and str(value).upper() not in {'YES', 'NO'}:
                        errors.append("Invalid value. Must be YES or NO")

                # Special handling for exclusion_type field
                if col_name == 'exclusion_type':
                    if is_exclusion_file:
                        if pd.isna(value):
                            errors.append("Exclusion Type is mandatory for exclusion files")
                        elif str(value).upper() not in valid_exclusion_types:
                            errors.append("Invalid Exclusion Type")

                validation_result = "Fail - " + ", ".join(errors) if errors else "Pass"
            else:
                validation_result = ""
            
            validation_row.append(validation_result)
        
        data_row["Individual_Status"] = individual_status
        formatted_output.append(data_row)
        formatted_output.append(pd.Series(validation_row, index=row.index))

    formatted_df = pd.DataFrame(formatted_output).reset_index(drop=True)
    return formatted_df

if __name__ == "__main__":
    # Get the downloads folder path and specify the fscs subfolder 
    downloads_path = Path.home() / "Downloads" / "fscs-testing"
    res_path = Path.home() / "Downloads" / "fscs-testing" / "results" 
    
    # Get all Excel files in the fscs files directory
    excel_files = [f for f in os.listdir(downloads_path) 
                  if f.endswith(('.xlsx', '.xls')) 
                  and not f.endswith('-result.xlsx')]
    
    print(f"Found {len(excel_files)} files to process")
    
    # Process each file
    for file_name in excel_files:
        try:
            input_file_path = os.path.join(downloads_path, file_name)
            
            # Get output file name (add -result before extension)
            output_name = file_name.rsplit('.', 1)[0] + '-result.xlsx'
            output_path = os.path.join(res_path, output_name)
            
            # Run validation
            validated_results = validate_file(input_file_path, data_inputs_df)
            
            # Save results
            validated_results.to_excel(output_path, index=False)
            print(f"Successfully processed {file_name} -> {output_name}")
            
        except Exception as e:
            print(f"Error processing {file_name}: {str(e)}")
            continue
