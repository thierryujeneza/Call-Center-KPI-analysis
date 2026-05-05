import os

os.makedirs('charts', exist_ok=True)
os.makedirs('call_center_analysis', exist_ok=True)


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# Load data
df = pd.read_excel('Data.xlsx')

# Clean column names
df.columns = df.columns.str.strip()

# Swap columns safely
if 'channel_type' in df.columns and 'Team' in df.columns:
    df[['channel_type', 'Team']] = df[['Team', 'channel_type']]

# Fix Names
df['Agent_Name_Clean'] = df['Agent_Name'].astype(str).str.title()

# Fill missing CSAT
df['CSAT_Score'] = df['CSAT_Score'].fillna(3)

# Fill missing Handle Time
mean_ht = df['Handle_Time_Min'].mean()
if pd.isna(mean_ht):
    mean_ht = 0
df['Handle_Time_Min'] = df['Handle_Time_Min'].fillna(mean_ht)

# Fix Date column
df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

# Missing date flag
df['Missing_Date_Flag'] = df['Date'].apply(lambda x: 'YES' if pd.isna(x) else 'NO')
print(df['Missing_Date_Flag'].value_counts())

# Handle time buckets
def ht_bucket(minutes):
    if minutes <= 5:
        return 'quick'
    elif minutes <= 10:
        return 'Standard'
    elif minutes <= 20:
        return 'Extended'
    else:
        return 'long'

df['HT_category'] = df['Handle_Time_Min'].apply(ht_bucket)
print(df['HT_category'].value_counts())

# ================== AGGREGATIONS ==================

fcr_by_team = df.groupby('Team').agg(
    Total_Cases=('Case_ID', 'count'),
    FCR_Yes=('FCR_Flag', lambda X: (X.astype(str).str.strip().str.lower() == 'yes').sum()),
    Avg_CSAT=('CSAT_Score', 'mean'),
    AVg_Handle_Time=('Handle_Time_Min', 'mean')
).reset_index()

fcr_by_team['FCR_Rate'] = fcr_by_team['FCR_Yes'] / fcr_by_team['Total_Cases']
print(fcr_by_team.round(2))

agent_scorecard = df.groupby(
    ['Agent_ID', 'Agent_Name', 'Team']
).agg(
    Total_Cases=('Case_ID', 'count'),
    Avg_CSAT=('CSAT_Score', 'mean'),
    Avg_Handle_Time=('Handle_Time_Min', 'mean'),
    FCR_Yes=('FCR_Flag', lambda x: (x.astype(str).str.strip().str.lower() == 'yes').sum()),
    Escalations=('Escalated', lambda X: (X.astype(str).str.strip().str.lower() == 'yes').sum())
).reset_index()

agent_scorecard['FCR_Rate'] = agent_scorecard['FCR_Yes'] / agent_scorecard['Total_Cases']
agent_scorecard['ESC_Rate'] = agent_scorecard['Escalations'] / agent_scorecard['Total_Cases']

print(agent_scorecard.round(2))

# ================== VISUAL 1 ==================

plt.style.use('dark_background')
sns.set_theme(style="dark", rc={"axes.facecolor": "#121212", "figure.facecolor": "#121212"})

fig, ax = plt.subplots(figsize=(10, 6))

colors = sns.color_palette("viridis", len(fcr_by_team))
sns.barplot(
    data=fcr_by_team,
    x='Team',
    y='FCR_Rate',
    palette=colors,
    edgecolor='white',
    linewidth=0.5,
    ax=ax
)

plt.title('First Call Resolution (FCR) Rate by Team', fontsize=18, fontweight='bold', color='white', pad=25)
plt.ylabel('FCR Rate (%)', fontsize=12, color='lightgray')
plt.xlabel('Support Channel', fontsize=12, color='lightgray')
plt.ylim(0, 1.0)

ax.tick_params(axis='x', colors='white')
ax.tick_params(axis='y', colors='white')

for i, v in enumerate(fcr_by_team['FCR_Rate']):
    ax.text(i, v + 0.02, f'{v:.1%}', color='white', ha='center')

plt.tight_layout()
plt.savefig('fcr_aesthetic_chart.png', dpi=300)
plt.show()
plt.close()

# ================== MONTHLY TREND ==================

monthly = df.dropna(subset=['Date']).groupby(
    df['Date'].dt.to_period('M')
).agg(
    Total_Cases=('Case_ID', 'count')
).reset_index()

monthly['Month'] = monthly['Date'].astype(str)

plt.figure(figsize=(12, 6))

plt.plot(monthly['Month'], monthly['Total_Cases'],
         marker='o', markersize=8,
         linewidth=3,
         markeredgecolor='white')

plt.fill_between(monthly['Month'], monthly['Total_Cases'], alpha=0.1)

plt.title('Monthly Case Volume Trend (2024)', fontsize=18, fontweight='bold', color='white')
plt.ylabel('Total Cases')
plt.xlabel('Month')

plt.xticks(rotation=45)
plt.ylim(0, monthly['Total_Cases'].max() * 1.2)

plt.tight_layout()
plt.savefig('monthly_trend_aesthetic.png', dpi=300)
plt.show()
plt.close()

# ================== SCATTER ==================

plt.figure(figsize=(10, 7))

sns.scatterplot(
    data=df,
    x='Handle_Time_Min',
    y='CSAT_Score',
    hue='Team',
    alpha=0.7,
    palette='flare',
    s=100,
    edgecolor='w'
)

plt.axhline(y=df['CSAT_Score'].mean(), linestyle=':', linewidth=2)
plt.axvline(x=df['Handle_Time_Min'].mean(), linestyle=':', linewidth=2)

plt.title('Correlation: Handle Time vs CSAT Score', fontsize=18, fontweight='bold')
plt.ylabel('CSAT Score')
plt.xlabel('Handle Time')

plt.legend()
plt.tight_layout()
plt.savefig('scatter_ht_csat_aesthetic.png', dpi=300)
plt.show()
plt.close()

# ================== EXTRA TABLE ==================

# Only create if column exists
if 'Issue_Type' in df.columns:
    csat_by_issue = df.groupby('Issue_Type').agg(
        Avg_CSAT=('CSAT_Score', 'mean'),
        Total_Cases=('Case_ID', 'count')
    ).reset_index()
else:
    csat_by_issue = pd.DataFrame()

# ================== SAVE TO EXCEL ==================

with pd.ExcelWriter('CCI_Python_Analysis.xlsx', engine='openpyxl') as writer:
    fcr_by_team.round(2).to_excel(writer, sheet_name='FCR_by_Team', index=False)

    if not csat_by_issue.empty:
        csat_by_issue.round(2).to_excel(writer, sheet_name='CSAT_by_Issue', index=False)

    agent_scorecard.round(2).to_excel(writer, sheet_name='Agent_Scorecard', index=False)

print('Excel file saved successfully')

# ================== CORRELATION ==================

numeric_cols = [col for col in [
    'Handle_Time_Min', 'Wait_Time_Min', 'CSAT_Score', 'Transfers'
] if col in df.columns]

if len(numeric_cols) > 1:
    corr_matrix = df[numeric_cols].corr()

    plt.figure(figsize=(7, 5))
    sns.heatmap(corr_matrix, annot=True, fmt='.2f',
                cmap='RdYlGn', center=0,
                square=True, linewidths=0.5)

    plt.title('Correlation Matrix — Call Center KPIs')
    plt.tight_layout()
    plt.savefig('correlation_heatmap.png', dpi=150)
    plt.show()
    plt.close()
    df.to_csv('call_center_cleaned_python.csv', index=False)
print(f'Saved {len(df)} rows to CSV')
print(df.info())