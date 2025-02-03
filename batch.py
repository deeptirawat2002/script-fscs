# Load required libraries
import pandas as pd
import re
from pathlib import Path
import os

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
   return bool(re.fullmatch(r'\d{2}\d{2}\d{4}', str(value))) if pd.notna(value) else True

def is_valid_email(value):
   # Basic email validation pattern
   email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
   return bool(re.match(email_pattern, str(value))) if pd.notna(value) else True

validation_functions = {
   'AlphaNumeric': is_alphanumeric,
   'Alpha': is_alpha,
   'Numeric': is_numeric,
   'Decimal': is_decimal,
   'Email': is_valid_email
}

def validate_file(file_path, rules_df):
   new_xls = pd.ExcelFile(file_path)
   new_data_df = new_xls.parse(new_xls.sheet_names[0]) 
   
   formatted_output = []
   seen_values = set()
   seen_account_numbers = set()
   seen_scv_records = set()  # New set to track unique SCV records
   
   valid_product_types = {'IAA', 'ISA', 'NA', 'FD1', 'FD2', 'FD4', 'Other'}
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

               # Account number uniqueness check
               if col_name == "account_number" and pd.notna(value):
                   if value in seen_account_numbers:
                       errors.append("Duplicate Account Number")
                   else:
                       seen_account_numbers.add(value)

               # Product type validation
               if col_name == "product_type" and pd.notna(value):
                   if value not in valid_product_types:
                       errors.append("Invalid Product Type")

               # Exclusion type validation
               if col_name == "exclusion_type" and pd.notna(value):
                   if value not in valid_exclusion_types:
                       errors.append("Invalid Exclusion Type")

               # Special handling for single_customer_view_record field
               if col_name == "single_customer_view_record":
                   if pd.notna(value):
                       # Convert to string, handling both numeric and text formats
                       str_val = str(value).strip()
                       # Check uniqueness
                       if str_val in seen_scv_records:
                           errors.append("Duplicate SCV Record")
                       else:
                           seen_scv_records.add(str_val)
                       # Check if it's either numeric or alphanumeric
                       if not (is_numeric(str_val) or is_alphanumeric(str_val)):
                           errors.append("Invalid Format - Must be Numeric or Alphanumeric")
                       # Check length if specified in rules
                       if max_length and len(str_val) > max_length:
                           errors.append("Exceeds Max Length")
               
               # Special handling for other numeric fields
               elif col_name in ["sort_code", "account_holder_indicator",
                              "account_balance_in_sterling", "authorised_negative_balances", 
                              "account_balance_in_original_currency",
                              "exchange_rate", "original_account_balance_before_interest"]:
                   if pd.notna(value):
                       try:
                           if col_name == "sort_code":
                               # First convert to string and remove any spaces/formatting
                               str_val = str(value).strip()
                               # Now force numeric conversion to handle "text formatted numbers"
                               num_val = int(float(str_val))
                               # Convert back to string to check length
                               clean_val = str(num_val)
                               if len(clean_val) > 6:  # Changed from != to >
                                   errors.append("Exceeds Max Length")
                           else:
                               # For other numeric fields
                               num_value = pd.to_numeric(str(value).strip())
                               str_val = str(num_value)
                               if max_length and len(str_val.replace(" ", "").replace("-", "")) > max_length:
                                   errors.append("Exceeds Max Length")
                       except:
                           errors.append(f"Invalid Numeric Format")

               # Length validation - for other fields
               elif max_length and len(str_value) > max_length:
                   errors.append("Exceeds Max Length")

               # Data type validation
               if col_name == "email_address" and pd.notna(value):
                   if not is_valid_email(value):
                       errors.append("Invalid Email Format")
               elif col_name == "main_phone_number":
                   if pd.notna(value):
                       if isinstance(value, (int, float)):
                           str_val = f"{value:.0f}"
                           if not str_val.isdigit():
                               errors.append("Invalid Phone Number Format")
                       else:
                           if not is_numeric(value):
                               errors.append("Invalid Phone Number Format")
               # Special handling for account_number and other alphanumeric fields that can be numeric
               elif data_type == 'AlphaNumeric':
                   if pd.notna(value):
                       # Pass if it's either numeric or alphanumeric
                       if not (is_numeric(value) or is_alphanumeric(value)):
                           errors.append(f"Invalid {data_type} Format")
               elif data_type in validation_functions and not validation_functions[data_type](value):
                   errors.append(f"Invalid {data_type} Format")
               
               # Individual-specific validations
               if col_name == "customer_first_forename":
                   if individual_status == "Individual" and pd.isna(value):
                       errors.append("Mandatory for Individual")
                   else:
                       mandatory = False
               
               if col_name == "other_national_identity_number":
                   if individual_status == "Individual" and pd.isna(value):
                       errors.append("Mandatory for Individual")
                   else:
                       mandatory = False
               
               if col_name == "other_national_identifier":
                   if individual_status == "Individual" and "other_national_identity_number" in new_data_df.columns:
                       if pd.notna(new_data_df.loc[index, "other_national_identity_number"]):
                           if pd.isna(value):
                               errors.append("Mandatory if other_national_identity_number is provided")
                           elif value not in ["NID", "DL", "O"]:
                               errors.append("Invalid Identifier Type")
                   else:
                       mandatory = False
               
               if col_name == "date_of_birth" and individual_status == "Individual" and not is_valid_date(value):
                   errors.append("Invalid Date Format (Should be DDMMYYYY)")
               
               # Modified address line validations
               if col_name == "address_line_3":
                   higher_lines = ["address_line_4", "address_line_5", "address_line_6"]
                   if any(pd.notna(row.get(col)) for col in higher_lines if col in new_data_df.columns):
                       if pd.isna(value):
                           errors.append("Mandatory if address_line_4 or address_line_5 is populated")
                   else:
                       mandatory = False
               
               if col_name == "address_line_4":
                   higher_lines = ["address_line_5", "address_line_6"]
                   if any(pd.notna(row.get(col)) for col in higher_lines if col in new_data_df.columns):
                       if pd.isna(value):
                           errors.append("Mandatory if address_line_5 or address_line_6 is populated")
                   else:
                       mandatory = False
               
               if col_name == "address_line_5":
                   if "address_line_6" in new_data_df.columns and pd.notna(row.get("address_line_6")):
                       if pd.isna(value):
                           errors.append("Mandatory if address_line_6 is populated")
                   else:
                       mandatory = False
               
               if mandatory and pd.isna(value):
                   errors.append("Missing Mandatory Value")
               
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
