import fitz  # Library name is PyMuPDF
import sys
import re

pdf_path = sys.argv[1]

def extract_form_fields(pdf_path):
    doc = fitz.open(pdf_path)
    fields = {}
    for page in doc:
        widgets = page.widgets()
        for widget in widgets:
            if widget.field_value:
                fields[widget.field_name] = widget.field_value
    return fields

form_fields = extract_form_fields(pdf_path)

# Combine the date of approval fields
date_of_approval = f"{form_fields.get('Date of Approval')}-{form_fields.get('undefined')}-{form_fields.get('undefined_2')}"

# Remove undefined fields after combining the date
if 'undefined' in form_fields:
    del form_fields['undefined']
if 'undefined_2' in form_fields:
    del form_fields['undefined_2']

# Add the combined date back to the fields
form_fields['Date of Approval'] = date_of_approval

# Define the options for "Type of Request" and "Requested SAP Roles"
type_of_request_options = {
    'Modify Existing User': 'Modify Existing User',
    'New User': 'New User',
    'Deactivate User': 'Deactivate User'
}

requested_sap_roles_options = {
    'SAP Basis Administrator': 'SAP Basis Administrator',
    'SAP FICO Consultant': 'SAP FICO Consultant',
    'SAP SD Consultant': 'SAP SD Consultant',
    'SAP HR Consultant': 'SAP HR Consultant',
    'SAP ABAP Developer': 'SAP ABAP Developer',
    'SAP QM Consultant': 'SAP QM Consultant',
    'SAP SCM Consultant': 'SAP SCM Consultant',
    'SAP GRC Consultant': 'SAP GRC Consultant',
    'SAP Security Consultant': 'SAP Security Consultant',
    'SAP WM Consultant': 'SAP WM Consultant',
    'SAP HANA Consultant': 'SAP HANA Consultant',
    'SAP Solution Architect': 'SAP Solution Architect',
    'SAP S/4HANA Consultant': 'SAP S/4HANA Consultant',
    'SAP CRM Consultant': 'SAP CRM Consultant',
    'SAP Project Manager': 'SAP Project Manager'
}

# Initialize variables to track the selected fields
selected_type_of_request = None
selected_sap_roles = []

# Check for the selected "Type of Request"
for key in type_of_request_options:
    if key in form_fields:
        selected_type_of_request = type_of_request_options[key]
        break

# Check for the selected "Requested SAP Roles"
for key in requested_sap_roles_options:
    if key in form_fields:
        selected_sap_roles.append(requested_sap_roles_options[key])

# Replace spaces with underscores in keys
formatted_fields = {re.sub(r'\s+', '_', key): value for key, value in form_fields.items()}

# Add the selected Type of Request and Requested SAP Roles to the output
if selected_type_of_request:
    formatted_fields['Type_of_Request'] = selected_type_of_request

if selected_sap_roles:
    formatted_fields['Requested_SAP_Roles'] = ', '.join(selected_sap_roles)

# Remove raw "On" fields that were processed
for key in list(formatted_fields.keys()):
    if formatted_fields[key] == 'On':
        del formatted_fields[key]

# Print fields with formatted keys
for field_name, field_value in formatted_fields.items():
    print(f"##gbStart##{field_name}##splitKeyValue##{field_value}##gbEnd##")

# Print field for Appstore node Evaluation
print("User Details Extracted Successfully")
