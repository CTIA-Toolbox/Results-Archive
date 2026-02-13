import pandas as pd

# Load the file from workspace (fresher copy)
input_file = r'C:\Users\dwhar\AppData\Local\Programs\Python\Python313\Scripts\Results Archive\Correlation All.xlsx'
output_file = r'C:\Users\dwhar\OneDrive\Documents\CTIA\Correlation All.xlsx'

print("Loading file...")
corr = pd.read_excel(input_file)

print("Setting Verizon and Google Chosen Location to NULL...")
# Set to None (which becomes NULL in Excel)
mask_verizon = corr['Participant'] == 'Verizon'
mask_google = corr['Participant'] == 'Google'

corr.loc[mask_verizon | mask_google, 'Chosen Location'] = None

print("Saving file...")
corr.to_excel(output_file, sheet_name='Sheet1', index=False, engine='openpyxl')

print("\nVerifying...")
corr_verify = pd.read_excel(output_file)
for p in ['AT&T', 'Google', 'T-Mobile', 'Verizon']:
    part_data = corr_verify[corr_verify['Participant'] == p]
    non_null = part_data['Chosen Location'].notna().sum()
    total = len(part_data)
    pct = (non_null / total * 100) if total > 0 else 0
    print(f'{p}: {non_null}/{total} filled ({pct:.1f}%)')

print("\nDone!")
