# ===============================
# NYC 311 Data Dashboard Project
# Role: PM/BA + Data Analyst
# Objective: Analyze NYC 311 complaints, visualize trends, and provide actionable insights
# Tools: Python (Pandas, Plotly)
# ===============================

import pandas as pd
import plotly.express as px

# -------------------------------
# Step 1: Load Dataset
# -------------------------------
# Download the CSV from: https://data.cityofnewyork.us/Social-Services/311-Service-Requests/erm2-nwe9
file_path = "311_Service_Requests.csv"  # Update with your path
df = pd.read_csv(file_path)

# -------------------------------
# Step 2: Keep Relevant Columns & Clean Data
# -------------------------------
df = df[['Created Date', 'Complaint Type', 'Borough', 'Status']]
df.dropna(inplace=True)  # Drop rows with missing values
df['Created Date'] = pd.to_datetime(df['Created Date'])
df['Month'] = df['Created Date'].dt.to_period('M')

# -------------------------------
# Step 3: Top Complaint Categories
# -------------------------------
top_categories = df['Complaint Type'].value_counts().head(5)
print("Top 5 Complaint Categories:\n", top_categories)

# Bar chart for top 5 complaint categories
fig1 = px.bar(top_categories, 
              x=top_categories.index, 
              y=top_categories.values,
              title='Top 5 NYC 311 Complaint Categories',
              labels={'x':'Complaint Type','y':'Number of Complaints'})
fig1.show()

# -------------------------------
# Step 4: Monthly Complaint Trends
# -------------------------------
monthly_trends = df.groupby('Month').size()
print("\nMonthly Complaint Trends:\n", monthly_trends)

# Line chart for monthly trend
fig2 = px.line(monthly_trends, 
               x=monthly_trends.index.astype(str), 
               y=monthly_trends.values,
               title='Monthly NYC 311 Complaint Trends',
               labels={'x':'Month','y':'Number of Complaints'})
fig2.show()

# -------------------------------
# Step 5: Complaint Distribution by Borough
# -------------------------------
borough_counts = df['Borough'].value_counts()
print("\nComplaint Counts by Borough:\n", borough_counts)

# Pie chart for borough distribution
fig3 = px.pie(borough_counts, 
              names=borough_counts.index, 
              values=borough_counts.values,
              title='Complaint Distribution by Borough')
fig3.show()

# -------------------------------
# Step 6: Optional - Top Complaint Types per Borough
# -------------------------------
top_per_borough = df.groupby(['Borough','Complaint Type']).size().reset_index(name='Count')
top_per_borough_sorted = top_per_borough.sort_values(['Borough','Count'], ascending=[True, False])
print("\nTop Complaint Types per Borough:\n", top_per_borough_sorted.groupby('Borough').head(3))

# -------------------------------
# Step 7: Insights & Recommendations (to include in portfolio)
# -------------------------------
# Example insights (replace with your observations from the charts):
insights = """
Insights:
- Noise complaints peak during summer months (June-August).
- Sanitation complaints are highest in Brooklyn and Queens.
- Targeted resource allocation can reduce response times for top complaint types.

Recommendations:
- Allocate additional patrol/resources during summer months for noise complaints.
- Prioritize sanitation resources in Brooklyn and Queens.
- Monitor trends monthly to adjust staffing and response strategies.
"""
print(insights)
