import matplotlib.pyplot as plt
import pandas as pd

xl = pd.ExcelFile("TeleClinic_Candidate_Dataset_Month3.xlsx")
patients      = pd.read_excel(xl, sheet_name=xl.sheet_names[1])
consults      = pd.read_excel(xl, sheet_name=xl.sheet_names[2])
labs          = pd.read_excel(xl, sheet_name=xl.sheet_names[4])
prescriptions = pd.read_excel(xl, sheet_name=xl.sheet_names[5])

# ── Data prep ──────────────────────────────────────────────
completion_rate = (consults['status']=='Completed').mean()*100
uploaded = labs[labs['result_uploaded']=='Yes']
review_rate = (uploaded['clinician_viewed']=='Yes').mean()*100
completed = consults[consults['status']=='Completed']
icd_rate = (completed['icd_code_entered']=='Yes').mean()*100
dispense_rate = (prescriptions['dispensed']=='Yes').mean()*100
weekly = consults.groupby('week_number')['consultation_id'].count()
rural_pct = (patients['urban_rural']=='Rural').mean()*100
urban_pct = 100 - rural_pct
not_viewed = uploaded[uploaded['clinician_viewed']=='No']
district_counts = not_viewed['district'].value_counts()

# ── Figure ─────────────────────────────────────────────────
fig = plt.figure(figsize=(20, 11))
fig.patch.set_facecolor('#F0F2F5')

# Title
fig.text(0.5, 0.97,
         'TeleClinic Platform — Clinical Governance Dashboard | Month 3',
         ha='center', fontsize=17, fontweight='bold', color='#1A1A2E')
fig.text(0.5, 0.935,
         'Central Question: Is the platform delivering safe, quality care — and for whom?',
         ha='center', fontsize=11, style='italic', color='#555555')

# ── KPI Cards ──────────────────────────────────────────────
metrics = [
    ('Consultation\nCompletion Rate', completion_rate, 80),
    ('Lab Result\nReview Rate',       review_rate,    100),
    ('ICD Code\nCompliance',          icd_rate,       100),
    ('Prescription\nDispensing Rate', dispense_rate,  100),
]

card_colors = ['#E74C3C', '#E74C3C', '#E74C3C', '#E74C3C']
# Green if meets target
for i, (_, val, target) in enumerate(metrics):
    if val >= target:
        card_colors[i] = '#27AE60'

for i, ((label, value, target), color) in enumerate(zip(metrics, card_colors)):
    ax = fig.add_axes([0.02 + i*0.245, 0.77, 0.22, 0.13])
    ax.set_facecolor(color)
    ax.text(0.5, 0.62, f'{value:.1f}%',
            ha='center', va='center',
            fontsize=26, fontweight='bold', color='white',
            transform=ax.transAxes)
    ax.text(0.5, 0.22, label,
            ha='center', va='center',
            fontsize=9, color='white',
            transform=ax.transAxes,
            linespacing=1.4)
    ax.set_xticks([])
    ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_visible(False)

# ── Weekly Adoption ────────────────────────────────────────
ax2 = fig.add_axes([0.02, 0.08, 0.44, 0.60])
ax2.bar(weekly.index, weekly.values, color='#3498DB',
        edgecolor='white', linewidth=0.5)
ax2.set_title('Platform Adoption Over Time\n(Weekly Consultations)',
              fontweight='bold', fontsize=11, pad=10)
ax2.set_xlabel('Week Number', fontsize=9)
ax2.set_ylabel('Number of Consultations', fontsize=9)
ax2.axvline(x=12.5, color='#E67E22', linestyle='--',
            linewidth=1.5, label='Week 13: possible data truncation')
ax2.legend(fontsize=8, loc='upper left')
ax2.set_facecolor('#FFFFFF')
ax2.tick_params(labelsize=8)
ax2.spines['top'].set_visible(False)
ax2.spines['right'].set_visible(False)

# ── Rural vs Urban ─────────────────────────────────────────
ax3 = fig.add_axes([0.54, 0.44, 0.44, 0.26])
categories = ['Platform Reach', 'Rwanda Benchmark']
rural_vals = [rural_pct, 83]
urban_vals = [urban_pct, 17]
x = [0, 1]
width = 0.35
ax3.bar([i - width/2 for i in x], rural_vals,
        width=width, label='Rural', color='#E67E22')
ax3.bar([i + width/2 for i in x], urban_vals,
        width=width, label='Urban', color='#3498DB')
ax3.set_title('Rural vs Urban Reach vs Rwanda Benchmark',
              fontweight='bold', fontsize=10, pad=8)
ax3.set_xticks(x)
ax3.set_xticklabels(categories, fontsize=9)
ax3.set_ylabel('Percentage (%)', fontsize=9)
ax3.legend(fontsize=8)
ax3.set_facecolor('#FFFFFF')
ax3.tick_params(labelsize=8)
ax3.spines['top'].set_visible(False)
ax3.spines['right'].set_visible(False)

# ── Unreviewed Labs ────────────────────────────────────────
ax4 = fig.add_axes([0.54, 0.08, 0.44, 0.28])
colors = ['#E74C3C' if d == 'Gasabo' else '#E08080'
          for d in district_counts.index]
ax4.barh(district_counts.index, district_counts.values,
         color=colors, edgecolor='white')
ax4.set_title('Unreviewed Lab Results by District  ⚠ Patient Safety Risk',
              fontweight='bold', fontsize=10, pad=8)
ax4.set_xlabel('Number of Unreviewed Results', fontsize=9)
ax4.set_facecolor('#FFFFFF')
ax4.tick_params(labelsize=8)
ax4.spines['top'].set_visible(False)
ax4.spines['right'].set_visible(False)

plt.savefig('dashboard/teleclinic_dashboard.png',
            dpi=150, bbox_inches='tight',
            facecolor='#F0F2F5')
print("Dashboard saved.")
plt.show()