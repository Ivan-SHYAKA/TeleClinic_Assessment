# =============================================================
# TeleClinic Platform — Data Quality Audit
# Author: Ivan Shyaka
# Assessment: MEL Associate | Irembo Digital Health
# Dataset: TeleClinic Candidate Dataset Month 3
# =============================================================

import pandas as pd

# Load all tables
xl = pd.ExcelFile("TeleClinic_Candidate_Dataset_Month3.xlsx")

patients      = pd.read_excel(xl, sheet_name=xl.sheet_names[1])
consults      = pd.read_excel(xl, sheet_name=xl.sheet_names[2])
followups     = pd.read_excel(xl, sheet_name=xl.sheet_names[3])
labs          = pd.read_excel(xl, sheet_name=xl.sheet_names[4])
prescriptions = pd.read_excel(xl, sheet_name=xl.sheet_names[5])
referrals     = pd.read_excel(xl, sheet_name=xl.sheet_names[6])
insurance     = pd.read_excel(xl, sheet_name=xl.sheet_names[7])

print("\nAll tables loaded successfully!")
print(f"\nPatients: {len(patients)} rows")
print(f"Consultations: {len(consults)} rows")
print(f"Follow-Ups: {len(followups)} rows")
print(f"Lab Tests: {len(labs)} rows")
print(f"Prescriptions: {len(prescriptions)} rows")
print(f"Referrals: {len(referrals)} rows")
print(f"Insurance Log: {len(insurance)} rows")

# =============================================================
# FINDING 1: Structural nulls in Consultations (307 rows)
# =============================================================
completed = consults[consults['status'] == 'Completed']
non_completed = consults[consults['status'] != 'Completed']

print("\n=== FINDING 1: Structural Nulls in Consultations ===")
print(f"\nCompleted consultations: {len(completed)}")
print(f"Non-completed (No-Show + Cancelled): {len(non_completed)}")
print(f"Null call_type on non-completed: {non_completed['call_type'].isna().sum()}")
print(f"\nConclusion: null values are only associated with non-completed consultations (No-Shows and Cancellations), not missing data.")

# =============================================================
# FINDING 2: Lab results uploaded but never viewed by clinician
# =============================================================
uploaded = labs[labs['result_uploaded'] == 'Yes']
not_viewed = uploaded[uploaded['clinician_viewed'] == 'No']

print("\n=== FINDING 2: Lab Results Uploaded but NOT Viewed by Clinician ===")
print(f"\nLab results uploaded: {len(uploaded)}")
print(f"Uploaded but NOT viewed by clinician: {len(not_viewed)}")
print(f"Percentage unreviewed: {len(not_viewed)/len(uploaded)*100:.1f}%")
print(f"\nLab results not viewed by test type:")
print(not_viewed['test_type'].value_counts())
print(f"\nLab results not viewed by district:")
print(not_viewed['district'].value_counts())
print(f"conclusion: 17.7% of uploaded lab results were never reviewed by the requesting clinician.")

# =============================================================
# FINDING 3: Lab results never uploaded
# =============================================================
not_uploaded = labs[labs['result_uploaded'] == 'No']

print("\n=== FINDING 3: Lab Results Never Uploaded ===")
print(f"\nTotal lab requests: {len(labs)}")
print(f"Results never uploaded: {len(not_uploaded)}")
print(f"Percentage: {len(not_uploaded)/len(labs)*100:.1f}%")
print(f"\nBy test type:")
print(not_uploaded['test_type'].value_counts())
print(f"\nBy district:")
print(not_uploaded['district'].value_counts())
print(f"\nConclusion: 6.2% of all lab requests have no result uploaded.")

# =============================================================
# FINDING 4: Prescriptions not dispensed
# =============================================================
not_dispensed = prescriptions[prescriptions['dispensed'] == 'No']
chronic_not_dispensed = not_dispensed[not_dispensed['drug_category'].isin(['Antihypertensives', 'Antidiabetics'])]

print("\n=== FINDING 4: Prescriptions Not Dispensed ===")
print(f"\nTotal prescriptions: {len(prescriptions)}")
print(f"Not dispensed: {len(not_dispensed)}")
print(f"Percentage: {len(not_dispensed)/len(prescriptions)*100:.1f}%")
print(f"\nBy drug category:")
print(not_dispensed['drug_category'].value_counts())
print(f"\nChronic medications not dispensed: {len(chronic_not_dispensed)}")
print(f"\nChronic medications not dispensed by district:")
print(chronic_not_dispensed.groupby(['district', 'drug_category']).size())
print(f"\nConclusion: 15.5% of prescriptions were never dispensed, including 27 chronic disease medications.")

# =============================================================
# FINDING 5: Referrals with no recorded outcome
# =============================================================
total_referrals = len(referrals)
decided = referrals['authorised'].notna().sum()
authorized = (referrals['authorised'] == 'Yes').sum()
rejected = (referrals['authorised'] == 'No').sum()
no_decision = referrals['authorised'].isna().sum()

print("\n=== FINDING 5: Referrals with No Recorded Outcome ===")
print(f"\nTotal referrals: {total_referrals}")
print(f"Authorized: {authorized}")
print(f"Rejected: {rejected}")
print(f"No recorded outcome: {no_decision}")
print(f"\nAuthorisation rate (of decided only): {authorized/decided*100:.1f}%")
print(f"Referrals with unknown outcome: {no_decision/total_referrals*100:.1f}%")

print(f"\nDatetime and processing hours by authorisation status:")
print(referrals.groupby('authorised', dropna=False)[['authorisation_datetime', 'processing_hours']].apply(lambda x: x.isnull().sum()))
print(f"\nConclusion: 17.6% of referrals have no recorded authorisation status, and these also have null values for authorisation datetime and processing hours.")

# =============================================================
# FINDING 6: Insurance API timeout failures
# =============================================================
total_entries = len(insurance)
failed = insurance[insurance['success'] == 'No']
api_timeout = failed[failed['failure_reason'] == 'API timeout']
never_succeeded = insurance.groupby('patient_id')['success'].apply(lambda x: (x=='Yes').any())
failed_all = never_succeeded[never_succeeded==False]

print("\n=== FINDING 6: Insurance API Timeout Failures ===")
print(f"\nTotal insurance validation attempts: {total_entries}")
print(f"Registered patients: {len(patients)}")
print(f"Extra attempts (retries): {total_entries - len(patients)}")
print(f"\nFailed validations: {len(failed)}")
print(f"\nFailure reasons:")
print(failed['failure_reason'].value_counts())
print(f"\nAPI timeout failures: {len(api_timeout)} ({len(api_timeout)/len(failed)*100:.1f}% of all failures)")
print(f"\nAttempt number distribution:")
print(insurance['attempt_number'].value_counts().sort_index())
print(f"\nPatients who never successfully validated: {len(failed_all)}")
print(f"\nConclusion: API timeout accounts for 22% of all failed validations, out of which 75% are first attempts.")

# =============================================================
# FINDING 7: Musanze implausibly low consultation volume
# =============================================================
musanze_patients = (patients['district'] == 'Musanze').sum()
musanze_consults = consults[consults['district'] == 'Musanze']
total_consults = len(consults)
total_patients = len(patients)

print("\n=== FINDING 7: Musanze Implausibly Low Consultation Volume ===")
print(f"\nConsultations by district:")
print(consults['district'].value_counts())
print(f"\nMusanze registered patients: {musanze_patients} ({musanze_patients/total_patients*100:.1f}% of all patients)")
print(f"Musanze consultations: {len(musanze_consults)} ({len(musanze_consults)/total_consults*100:.1f}% of all consultations)")
print(f"\nMusanze consultations by week:")
print(musanze_consults['week_number'].value_counts().sort_index())
print(f"\nConclusion: Musanze accounts for only 0.5% of consultations despite having 1.75% of registered patients, and consultations only appear from week 9 onwards.")

# =============================================================
# FINDING 8: Booking datetime plausibility issue
# =============================================================
consults['booked_dt'] = pd.to_datetime(consults['booked_datetime'])
consults['hour'] = consults['booked_dt'].dt.hour
night_bookings = consults[(consults['hour'] >= 0) & (consults['hour'] < 6)]

print("\n=== FINDING 8: Booking Datetime Plausibility Issue ===")
print(f"\nConsultations booked by hour:")
print(consults['hour'].value_counts().sort_index())
print(f"\nConsultations booked between midnight and 6am: {len(night_bookings)}")
print(f"Percentage of total: {len(night_bookings)/len(consults)*100:.1f}%")
print(f"\nConclusion: Bookings are distributed evenly across all 24 hours including midnight to 6am, which is implausible for a platform serving rural Rwanda.")

# =============================================================
# PART 2: METRICS & ANALYSIS
# =============================================================

# METRIC 1: Consultation Completion Rate
print("=== METRIC 1: CONSULTATION COMPLETION RATE ===")
print(consults['status'].value_counts())
print(consults['status'].value_counts(normalize=True).mul(100).round(1))

# METRIC 2: Lab Result Review Rate
print("\n=== METRIC 2: LAB RESULT REVIEW RATE ===")
uploaded = labs[labs['result_uploaded'] == 'Yes']
print(uploaded['clinician_viewed'].value_counts())
print(uploaded['clinician_viewed'].value_counts(normalize=True).mul(100).round(1))

# METRIC 3: Clinical Protocol Compliance
print("\n=== METRIC 3: CLINICAL PROTOCOL COMPLIANCE ===")
completed_consults = consults[consults['status'] == 'Completed']
print(completed_consults['icd_code_entered'].value_counts())
print(completed_consults['icd_code_entered'].value_counts(normalize=True).mul(100).round(1))

# METRIC 4: Prescription Dispensing Rate
print("\n=== METRIC 4: PRESCRIPTION DISPENSING RATE ===")
print(prescriptions['dispensed'].value_counts())
print(prescriptions['dispensed'].value_counts(normalize=True).mul(100).round(1))

# METRIC 5: Platform Adoption Over Time
print("\n=== METRIC 5: PLATFORM ADOPTION OVER TIME ===")
weekly = consults.groupby('week_number')['consultation_id'].count()
print(weekly)

# =============================================================
# PART 3: EQUITY ANALYSIS
# =============================================================

# Equity Measure 1: Urban/Rural representation gap
print("=== EQUITY MEASURE 1: URBAN/RURAL REPRESENTATION ===")
rural_patients = (patients['urban_rural'] == 'Rural').sum()
urban_patients = (patients['urban_rural'] == 'Urban').sum()
total_patients = len(patients)

print(f"Rural patients: {rural_patients} ({rural_patients/total_patients*100:.1f}%)")
print(f"Urban patients: {urban_patients} ({urban_patients/total_patients*100:.1f}%)")
print(f"Rwanda benchmark: 83% rural / 17% urban")
print(f"Gap: Rural underrepresented by {83 - rural_patients/total_patients*100:.1f} percentage points")

# Equity Measure 2: Uninsured rate by urban/rural
print("\n=== EQUITY MEASURE 2: UNINSURED RATE BY GEOGRAPHY ===")
rural = patients[patients['urban_rural'] == 'Rural']
urban = patients[patients['urban_rural'] == 'Urban']
print(f"Rural uninsured: {(rural['insurance_scheme']=='Uninsured').sum()} ({(rural['insurance_scheme']=='Uninsured').mean()*100:.1f}%)")
print(f"Urban uninsured: {(urban['insurance_scheme']=='Uninsured').sum()} ({(urban['insurance_scheme']=='Uninsured').mean()*100:.1f}%)")

# Equity Measure 3: Channel usage by urban/rural
print("\n=== EQUITY MEASURE 3: CHANNEL BY GEOGRAPHY ===")
print(patients.groupby(['urban_rural','channel']).size())
print("Note: Higher IremboApp usage in rural areas suggests most vulnerable patients may be excluded")

# Equity Measure 4: Gender equity
print("\n=== EQUITY MEASURE 4: GENDER SPLIT ===")
print(f"Registered patients:")
print(patients['gender'].value_counts(normalize=True).mul(100).round(1))
print(f"Consultations:")
print(consults['gender'].value_counts(normalize=True).mul(100).round(1))
print("Conclusion: Gender equity is maintained — 50/50 split across both registration and consultations")