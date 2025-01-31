# Cell 1: Load main FSCS CSV file
import pandas as pd
import re

file_path = "fscs_scv_tables.xlsx"
data_inputs_df = pd.read_excel(file_path, sheet_name="Data inputs")

# Cell 2: Define validation functions and logic
def is_alphanumeric(value):
   return bool(re.fullmatch(r"[A-Za-z0-9 '-]+", str(value))) if pd.notna(value) else True

def is_alpha(value):
   return bool(re.fullmatch(r"[A-Za-z '-]+", str(value))) if pd.notna(value) else True

def is_numeric(value):
   try:
       # Convert scientific notation or any number format to string and remove any spaces
       if isinstance(value, (int, float)):
           # Convert number to string without scientific notation
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

               # Length validation
               if max_length and len(str_value) > max_length:
                   errors.append("Exceeds Max Length")

               # Data type validation
               if col_name == "email_address" and pd.notna(value):
                   if not is_valid_email(value):
                       errors.append("Invalid Email Format")
               elif col_name == "main_phone_number":
                   if pd.notna(value):
                       if isinstance(value, (int, float)):
                           # Handle numeric values (including scientific notation)
                           str_val = f"{value:.0f}"
                           if not str_val.isdigit():
                               errors.append("Invalid Phone Number Format")
                       else:
                           # Handle text values
                           if not is_numeric(value):
                               errors.append("Invalid Phone Number Format")
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

# For second file
file_to_validate = "addtophonenum.xlsx"
validated_results = validate_file(file_to_validate, data_inputs_df)
print(validated_results)

validated_results.to_excel('addtophonenum-res-2nd.xlsx', index=False)
