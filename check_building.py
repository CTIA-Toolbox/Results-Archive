import pandas as pd

# Check archive data
arch = pd.read_excel(r'C:\Users\dwhar\AppData\Local\Programs\Python\Python313\Scripts\Results Archive\Building Results.xlsx')

# Check if ATLUBC107 exists
building = 'ATLUBC107'
arch_building = arch[arch['Building ID'] == building]

if len(arch_building) > 0:
    print(f'Building {building} found in archive: {len(arch_building)} rows')
    print(f'Stages available: {list(arch_building["Stage"].unique())}')
    print(f'OEM rows: {len(arch_building[arch_building["Stage"] == "OEM"])}')
else:
    print(f'Building {building} NOT found in archive')
    print(f'Available buildings (first 20):')
    buildings = arch['Building ID'].unique()[:20]
    for b in buildings:
        print(f'  {b}')

# Check correlation data
print()
corr = pd.read_excel(r'C:\Users\dwhar\OneDrive\Documents\CTIA\Correlation All.xlsx')
corr_building = corr[corr['Building ID'] == building]
if len(corr_building) > 0:
    print(f'Building {building} found in correlation: {len(corr_building)} rows')
    print(f'Stages: {list(corr_building["Stage"].unique())}')
    print(f'Participants: {list(corr_building["Participant"].unique())}')
else:
    print(f'Building {building} NOT found in correlation')
