import pandas as pd

xl = pd.ExcelFile("TeleClinic_Candidate_Dataset_Month3.xlsx")

consults = pd.read_excel(xl, sheet_name=xl.sheet_names[2])

# health check on every table at once
for i, name in enumerate(xl.sheet_names):
    df = pd.read_excel(xl, sheet_name=xl.sheet_names[i])
    print(f"\n{'='*50}")
    print(f"SHEET: {name}")
    print(f"Shape: {df.shape}")
    print(f"Null counts:")
    print(df.isnull().sum())

print("\n=== CONSULTATIONS: BASIC HEALTH CHECK ===")
print("Rows and columns:", consults.shape)

print("\nNull count per column:")
print(consults.isnull().sum())

print("\nStatus breakdown:")
print(consults['status'].value_counts())

# Confirm: are the 307 nulls ONLY on no-shows and cancellations?
null_rows = consults[consults['call_type'].isna()]
print("\nStatus of rows where call_type is null:")
print(null_rows['status'].value_counts())


labs = pd.read_excel(xl, sheet_name=xl.sheet_names[4])

print("\n=== LAB TESTS: INVESTIGATING NULLS ===")
print(f"\nTotal lab records: {len(labs)}")
print(f"\nNull counts summary:")
print(f"  upload_datetime : {labs['upload_datetime'].isna().sum()}")
print(f"  tat_hours       : {labs['tat_hours'].isna().sum()}")
print(f"  clinician_viewed: {labs['clinician_viewed'].isna().sum()}")
print(f"  hours_to_view   : {labs['hours_to_view'].isna().sum()}")

print("\n--- When upload_datetime is null, what is result_uploaded?")
null_upload = labs[labs['upload_datetime'].isna()]
print(f"Rows affected: {len(null_upload)}")
print(null_upload['result_uploaded'].value_counts())

print("\n--- When tat_hours is null, what is result_uploaded?")
null_tat = labs[labs['tat_hours'].isna()]
print(f"Rows affected: {len(null_tat)}")
print(null_tat['result_uploaded'].value_counts())

print("\n--- When clinician_viewed is null, what is result_uploaded?")
null_viewed = labs[labs['clinician_viewed'].isna()]
print(f"Rows affected: {len(null_viewed)}")
print(null_viewed['result_uploaded'].value_counts())

print("\n--- When hours_to_view is null, what is clinician_viewed?")
null_hours = labs[labs['hours_to_view'].isna()]
print(f"Rows affected: {len(null_hours)}")
print(null_hours['clinician_viewed'].value_counts(dropna=False))

print("\n--- Results uploaded but hours_to_view still null")
uploaded_but_no_view = labs[(labs['result_uploaded'] == 'Yes') & (labs['hours_to_view'].isna())]
print(f"Rows affected: {len(uploaded_but_no_view)}")
print(uploaded_but_no_view['clinician_viewed'].value_counts(dropna=False))

print("\n--- What test types are uploaded but never viewed by clinician?")
uploaded_not_viewed = labs[(labs['result_uploaded'] == 'Yes') & (labs['clinician_viewed'] == 'No')]
print(f"Total unreviewed results: {len(uploaded_not_viewed)}")
print("\nBy test type:")
print(uploaded_not_viewed['test_type'].value_counts())
print("\nBy district:")
print(uploaded_not_viewed['district'].value_counts())

prescriptions = pd.read_excel(xl, sheet_name=xl.sheet_names[5])

print("\n=== PRESCRIPTIONS: INVESTIGATING NULLS ===")
print(f"\nTotal prescriptions: {len(prescriptions)}")
print(f"\nNull counts:")
print(f"  dispensed_datetime: {prescriptions['dispensed_datetime'].isna().sum()}")
print(f"  lag_hours         : {prescriptions['lag_hours'].isna().sum()}")

print("\n--- When dispensed_datetime is null, what is dispensed?")
null_dispensed = prescriptions[prescriptions['dispensed_datetime'].isna()]
print(f"Rows affected: {len(null_dispensed)}")
print(null_dispensed['dispensed'].value_counts())

print("\n--- What drug categories are not being dispensed?")
not_dispensed = prescriptions[prescriptions['dispensed'] == 'No']
print(f"Total not dispensed: {len(not_dispensed)}")
print(f"% of all prescriptions: {len(not_dispensed)/len(prescriptions)*100:.1f}%")
print("\nBy drug category:")
print(not_dispensed['drug_category'].value_counts())
print("\nBy district:")
print(not_dispensed['district'].value_counts())


print("\n--- Not dispensed by district, chronic only:")
chronic_not_dispensed = not_dispensed[not_dispensed['drug_category'].isin(['Antihypertensives', 'Antidiabetics'])]
print(f"Total chronic prescriptions not dispensed: {len(chronic_not_dispensed)}")
print("\nBy district:")
print(chronic_not_dispensed['district'].value_counts())
print("\nBy drug category and district:")
print(chronic_not_dispensed.groupby(['district', 'drug_category']).size())


referrals = pd.read_excel(xl, sheet_name=xl.sheet_names[6])

print("\n=== REFERRALS: BASIC HEALTH CHECK ===")
print(f"Total referrals: {len(referrals)}")
print(f"\nNull counts:")
print(referrals.isnull().sum())
print(f"\nAuthorised breakdown:")
print(referrals['authorised'].value_counts(dropna=False))


print("\n--- Datetime and processing hours by authorisation status:")
print(referrals.groupby('authorised', dropna=False)[['authorisation_datetime', 'processing_hours']].apply(lambda x: x.isnull().sum()))

print("\n--- Referral authorisation rate (honest calculation):")
total = len(referrals)
decided = referrals['authorised'].notna().sum()
authorized = (referrals['authorised'] == 'Yes').sum()
no_decision = referrals['authorised'].isna().sum()

print(f"Total referrals: {total}")
print(f"With a recorded decision: {decided}")
print(f"Authorized: {authorized} ({authorized/decided*100:.1f}% of decided)")
print(f"No recorded outcome: {no_decision} ({no_decision/total*100:.1f}% of all referrals)")

insurance = pd.read_excel(xl, sheet_name=xl.sheet_names[7])

print("\n=== INSURANCE LOG: INVESTIGATING NULLS ===")
print(f"Total insurance log entries: {len(insurance)}")
print(f"\nSuccess breakdown:")
print(insurance['success'].value_counts())
print(f"\nWhen failure_reason is null, what is success?")
null_reason = insurance[insurance['failure_reason'].isna()]
print(f"Rows affected: {len(null_reason)}")
print(null_reason['success'].value_counts())

print("\n--- When validation fails, what are the reasons?")
failed = insurance[insurance['success'] == 'No']
print(f"Total failed validations: {len(failed)}")
print(f"\nFailure reasons:")
print(failed['failure_reason'].value_counts())
print(f"\nAttempt number distribution:")
print(insurance['attempt_number'].value_counts().sort_index())
print(f"\nPatients who never successfully validated:")
never_succeeded = insurance.groupby('patient_id')['success'].apply(lambda x: (x=='Yes').any())
failed_all = never_succeeded[never_succeeded==False]
print(f"Count: {len(failed_all)}")


patients = pd.read_excel(xl, sheet_name=xl.sheet_names[1])
consults = pd.read_excel(xl, sheet_name=xl.sheet_names[2])
followups = pd.read_excel(xl, sheet_name=xl.sheet_names[3])
referrals = pd.read_excel(xl, sheet_name=xl.sheet_names[6])

print("\n=== CROSS TABLE CHECKS ===")

print("\n--- Do all patients in Consultations exist in Patients table?")
cons_patients = set(consults['patient_id'])
reg_patients = set(patients['patient_id'])
in_consults_not_registered = cons_patients - reg_patients
print(f"Patient IDs in Consultations but not in Patients: {len(in_consults_not_registered)}")

print("\n--- Do all Follow-Up consult IDs exist in Consultations?")
followup_consults = set(followups['original_consult_id'])
all_consults = set(consults['consultation_id'])
followup_not_matched = followup_consults - all_consults
print(f"Follow-up IDs not matching any consultation: {len(followup_not_matched)}")

print("\n--- Do all Referral consult IDs exist in Consultations?")
referral_consults = set(referrals['consultation_id'])
referral_not_matched = referral_consults - all_consults
print(f"Referral consult IDs not matching any consultation: {len(referral_not_matched)}")

print("\n=== CONSULTATIONS BY DISTRICT ===")
print(f"Total consultations by district: {consults['district'].value_counts().sum()}")
print(consults['district'].value_counts())


print("\n=== PATIENTS REGISTERED BY DISTRICT ===")
print(f"Total patients by district: {patients['district'].value_counts().sum()}")
print(patients['district'].value_counts())

print("\n=== MUSANZE: WHICH WEEKS DID CONSULTATIONS HAPPEN? ===")
musanze = consults[consults['district'] == 'Musanze']
print(f"Total Musanze consultations: {len(musanze)}")
print(f"\nBy week number:")
print(musanze['week_number'].value_counts().sort_index())
print(f"\nPlatform runs weeks 1-13. Musanze appears in weeks:")
print(sorted(musanze['week_number'].unique()))

print("\n=== BOOKED DATETIME: WHAT HOURS ARE CONSULTATIONS BOOKED? ===")
consults['booked_dt'] = pd.to_datetime(consults['booked_datetime'])
consults['hour'] = consults['booked_dt'].dt.hour
print(consults['hour'].value_counts().sort_index())

print("\n=== Consultation Completion Rate ===")
print(f"Total consultations: {len(consults)}")
print(consults['status'].value_counts())
print(consults['status'].value_counts(normalize=True).mul(100).round(1))


print("\n=== Lab Result Review Rate ===")
uploaded = labs[labs['result_uploaded'] == 'Yes']
print(uploaded['clinician_viewed'].value_counts())
print(uploaded['clinician_viewed'].value_counts(normalize=True).mul(100).round(1))


print("\n=== Clinical Protocol Compliance ===")
completed_consults = consults[consults['status'] == 'Completed']
print(completed_consults['icd_code_entered'].value_counts())
print(completed_consults['icd_code_entered'].value_counts(normalize=True).mul(100).round(1))
print(completed_consults['notes_entered'].value_counts())
print(completed_consults['notes_entered'].value_counts(normalize=True).mul(100).round(1))


print("\n=== Prescription Dispensing Rate ===")
print(prescriptions['dispensed'].value_counts())
print(prescriptions['dispensed'].value_counts(normalize=True).mul(100).round(1))

weekly = consults.groupby('week_number')['consultation_id'].count()
print("\n=== Platform Adoption Over Time (Consultations per Week) ===")
print(weekly)


print("\n=== URBAN/RURAL SPLIT ===")
print("Registered patients:")
print(patients['urban_rural'].value_counts(normalize=True).mul(100).round(1))

print("\nConsultations:")
print(consults['urban_rural'].value_counts(normalize=True).mul(100).round(1))

print("\n=== CHANNEL BY URBAN/RURAL ===")
print(patients.groupby(['urban_rural','channel']).size())

print("\n=== GENDER SPLIT ===")
print("Registered patients:")
print(patients['gender'].value_counts(normalize=True).mul(100).round(1))
print("\nConsultations:")
print(consults['gender'].value_counts(normalize=True).mul(100).round(1))

print("\n=== INSURANCE BY URBAN/RURAL ===")
print(patients.groupby(['urban_rural','insurance_scheme']).size())

print("=== UNINSURED RATE BY URBAN/RURAL ===")
rural = patients[patients['urban_rural'] == 'Rural']
urban = patients[patients['urban_rural'] == 'Urban']
print(f"Rural uninsured: {(rural['insurance_scheme']=='Uninsured').sum()} ({(rural['insurance_scheme']=='Uninsured').mean()*100:.1f}%)")
print(f"Urban uninsured: {(urban['insurance_scheme']=='Uninsured').sum()} ({(urban['insurance_scheme']=='Uninsured').mean()*100:.1f}%)")