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

def is_valid_bfpo(value):
    """Validate BFPO address format"""
    if pd.isna(value):
        return True
    bfpo_pattern = r'^BFPO\s+\d+$'
    return bool(re.match(bfpo_pattern, str(value).upper()))

def is_valid_phone_number(value):
    """Validate phone number format including international prefix"""
    if pd.isna(value):
        return True
    # Allow +, 00 prefix and only digits
    phone_pattern = r'^(\+|00)?[0-9]{1,15}$'
    return bool(re.match(phone_pattern, str(value).strip()))

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

def is_valid_country_code(value):
    """Validate country codes"""
    if pd.isna(value):
        return True
    return str(value).upper() in {'GBR', 'GIB'}

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
    new_xls = pd.ExcelFile(file_path)
    new_data_df = new_xls.parse(new_xls.sheet_names[0])
    
    formatted_output = []
    seen_values = set()
    seen_account_numbers = set()
    seen_scv_records = set()
    seen_addresses = set()
    
    valid_product_types = {'IAA', 'ISA', 'NA', 'FD1', 'FD2', 'FD4', 'FP4P', 'Other'}
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

                # Previous validations...
                # Account number uniqueness check
                if col_name == "account_number" and pd.notna(value):
                    if value in seen_account_numbers:
                        errors.append("Duplicate Account Number")
                    else:
                        seen_account_numbers.add(value)

                # Add these new validations after the existing validation blocks:

                # THB (Temporary High Balance) validation
                if col_name == "account_balance_in_sterling" and pd.notna(value):
                    try:
                        balance = float(value)
                        if balance > 85000:  # FSCS compensation limit
                            errors.append("Potential THB - Balance exceeds compensation limit")
                    except (ValueError, TypeError):
                        pass

                # Sub-fund election validation for trusts
                if col_name == "account_title" and pd.notna(value):
                    if "TRUST" in str(value).upper() and "SUB" in str(value).upper():
                        if not any(keyword in str(value).upper() for keyword in ["HMRC", "ELECTION"]):
                            errors.append("Trust Sub-fund without election reference")

                # Junior ISA and Child Trust Fund validation
                if col_name == "product_type" and pd.notna(value):
                    if value == "ISA":
                        account_title = str(row.get("account_title", "")).upper()
                        if any(keyword in account_title for keyword in ["JUNIOR", "JISA", "CHILD TRUST"]):
                            errors.append("Junior ISA/Child Trust Fund should be in Exclusions View")

                # PO Box validation
                if col_name.startswith("address_line") and pd.notna(value):
                    if "PO BOX" in str(value).upper():
                        errors.append("PO Box address found - Verify delivery capability")

                # Prison address validation
                if col_name == "address_line_1" and pd.notna(value):
                    prison_keywords = ["HMP", "PRISON", "CORRECTIONAL"]
                    if any(keyword in str(value).upper() for keyword in prison_keywords):
                        if not re.match(r'^[A-Z0-9]+\s', str(value)):
                            errors.append("Missing prisoner number in prison address")

                # Account branch jurisdiction validation
                if col_name == "account_branch_jurisdiction" and pd.notna(value):
                    if str(value).upper() not in ["GBR", "GIB"]:
                        errors.append("Invalid branch jurisdiction - Must be GBR or GIB")

                # Continuity of access validation
                if col_name == "product_type" and pd.notna(value):
                    product_hierarchy = {
                        "IAA": 1,  # Instant Access highest priority
                        "ISA": 2,
                        "NA": 3,
                        "FD1": 4,
                        "FD2": 5,
                        "FD4": 6,
                        "Other": 7
                    }
                    if str(value) in product_hierarchy:
                        product_priority = product_hierarchy[str(value)]
                        transferable_eligible = row.get("transferable_eligible_deposit")
                        if pd.notna(transferable_eligible) and float(transferable_eligible) > 0:
                            for other_product in seen_values:
                                if product_hierarchy.get(other_product, 999) < product_priority:
                                    errors.append("Product hierarchy violation for continuity of access")
                        seen_values.add(str(value))

                # Currency conversion validation
                if col_name == "account_balance_in_sterling" and pd.notna(value):
                    orig_currency = row.get("currency_of_account")
                    if pd.notna(orig_currency) and orig_currency != "GBP":
                        orig_balance = row.get("account_balance_in_original_currency")
                        exchange_rate = row.get("exchange_rate")
                        if pd.notna(orig_balance) and pd.notna(exchange_rate):
                            expected_sterling = float(orig_balance) * float(exchange_rate)
                            if abs(float(value) - expected_sterling) > 0.01:  # Allow for rounding differences
                                errors.append("Currency conversion mismatch")

                # New validations
                if col_name.startswith('address_line_'):
                    # Address continuity check
                    line_num = int(col_name[-1])
                    next_line = f'address_line_{line_num + 1}'
                    if next_line in row and pd.notna(row[next_line]) and pd.isna(value):
                        errors.append("Address Line Continuity Error")

                    # BFPO validation
                    if line_num == 1 and 'BFPO' in str_value.upper():
                        if not is_valid_bfpo(value):
                            errors.append("Invalid BFPO Format")

                    # Care of address check
                    if line_num == 1 and 'C/O' in str_value.upper():
                        errors.append("Care of Address - NFFSTP")

                    # Duplicate address check
                    if line_num == 1:
                        address_key = f"{value}_{row.get('postcode', '')}"
                        if address_key in seen_addresses:
                            errors.append("Duplicate Address")
                        seen_addresses.add(address_key)

                # Phone number validation
                if col_name in ['main_phone_number', 'evening_phone_number', 'mobile_phone_number']:
                    if pd.notna(value) and not is_valid_phone_number(value):
                        errors.append("Invalid Phone Number Format")

                # IBAN validation
                if col_name == 'iban' and pd.notna(value):
                    if not is_valid_iban(value):
                        errors.append("Invalid IBAN Format")

                # BIC validation
                if col_name == 'bic' and pd.notna(value):
                    if not is_valid_bic(value):
                        errors.append("Invalid BIC Format")

                # BRRD flag validation
                if col_name == 'brrd_flag':
                    if pd.notna(value) and str(value).upper() not in ['YES', 'NO']:
                        errors.append("Invalid BRRD Flag")

                # Structured deposit validation
                if col_name == 'structured_deposit_accounts':
                    if pd.notna(value) and str(value).upper() not in ['YES', 'NO']:
                        errors.append("Invalid Structured Deposit Flag")

                # Name validations
                if col_name == 'customer_first_forename':
                    if contains_only_initials(value):
                        errors.append("First Name Contains Only Initials")
                    if pd.notna(value) and value in [row.get('customer_second_forename'), 
                                                   row.get('customer_third_forename')]:
                        errors.append("Repeated Forename")

                if col_name == 'surname':
                    if is_short_name(value):
                        errors.append("Surname Too Short")

                # Country code validation
                if col_name == 'country':
                    if pd.notna(value) and not is_valid_country_code(value):
                        errors.append("Invalid Country Code")

                # ASCII range validation for all fields
                if not validate_ascii_range(value):
                    errors.append("Invalid Characters Outside ASCII Range")

                validation_result = "Fail - " + ", ".join(errors) if errors else "Pass"
            else:
                validation_result = ""
            
            validation_row.append(validation_result)
        
        data_row["Individual_Status"] = individual_status
        formatted_output.append(data_row)
        formatted_output.append(pd.Series(validation_row, index=row.index))

    formatted_df = pd.DataFrame(formatted_output).reset_index(drop=True)

    # Validate file footer
    footer = '9' * 20
    if not str(formatted_df.iloc[-1:].to_string()).endswith(footer):
        print("Warning: Missing or invalid file footer (20 repeated '9's)")

    return formatted_df

if __name__ == "__main__":
    # Get the downloads folder path and specify the fscs subfolder
    downloads_path = Path.home() / "Downloads" / "fscs files"
    
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
            output_path = os.path.join(downloads_path, output_name)
            
            # Run validation
            validated_results = validate_file(input_file_path, data_inputs_df)
            
            # Save results
            validated_results.to_excel(output_path, index=False)
            print(f"Successfully processed {file_name} -> {output_name}")
            
        except Exception as e:
            print(f"Error processing {file_name}: {str(e)}")
            continue
        
