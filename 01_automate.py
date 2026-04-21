import pandas as pd
import numpy as np

file_path = 'input_data/Source File 4-13.xlsx'

sheet_names = [
    "EDFT",
    "TH Summary",
    "TH Hourly",
    "ML Store",
    "Emp List",
    "Apr Unapr OT",
    "Check Pivot",
    "Current Week OT",
    "Sheet10",
    "SCH",
    "SCH for MO"
]

dfs = {}

for sheet in sheet_names:
    dfs[sheet] = pd.read_excel(file_path, sheet_name=sheet)

dfs.keys()

# Access example:
df_edft = dfs["EDFT"]
df_th_summary = dfs["TH Summary"]
df_th_hourly = dfs["TH Hourly"]

# df_edft.head(10)

# ----------------------------
# STANDARDIZE COLUMNS
# ----------------------------
df_th_hourly["Date"] = pd.to_datetime(df_th_hourly["Date"], errors="coerce")
df_th_hourly["Hours"] = pd.to_numeric(df_th_hourly["Hours"], errors="coerce")

# OT IDENTIFICATION
ot_mask = df_th_hourly["Earning Desc"].str.contains("OT", case=False, na=False)

# ----------------------------
# 1. EMPLOYEE METRICS
# ----------------------------
emp_total_hours = df_th_hourly.groupby("EE Code")["Hours"].sum().rename("total_hours")

emp_ot_hours = df_th_hourly[ot_mask].groupby("EE Code")["Hours"].sum().rename("ot_hours")

emp_days_worked = df_th_hourly.groupby("EE Code")["Date"].nunique().rename("days_worked")

emp_daily_hours = df_th_hourly.groupby(["EE Code", "Date"])["Hours"].sum().reset_index()
emp_avg_daily_hours = emp_daily_hours.groupby("EE Code")["Hours"].mean().rename("avg_daily_hours")

emp_metrics = pd.concat([
    emp_total_hours,
    emp_ot_hours,
    emp_days_worked,
    emp_avg_daily_hours
], axis=1).fillna(0)

emp_metrics["regular_hours"] = emp_metrics["total_hours"] - emp_metrics["ot_hours"]
emp_metrics["ot_ratio"] = emp_metrics["ot_hours"] / emp_metrics["total_hours"].replace(0, np.nan)

# ----------------------------
# 2. POSITION METRICS
# ----------------------------
position_total_hours = df_th_hourly.groupby("Position_Title")["Hours"].sum().rename("total_hours")

position_ot_hours = df_th_hourly[ot_mask].groupby("Position_Title")["Hours"].sum().rename("ot_hours")

position_headcount = df_th_hourly.groupby("Position_Title")["EE Code"].nunique().rename("headcount")

position_metrics = pd.concat([
    position_total_hours,
    position_ot_hours,
    position_headcount
], axis=1).fillna(0)

position_metrics["avg_hours_per_employee"] = (
    position_metrics["total_hours"] / position_metrics["headcount"].replace(0, np.nan)
)

position_metrics["ot_ratio"] = (
    position_metrics["ot_hours"] / position_metrics["total_hours"].replace(0, np.nan)
)

# ----------------------------
# 3. DEPARTMENT METRICS
# ----------------------------
dept_total_hours = df_th_hourly.groupby("Dist Department Code")["Hours"].sum().rename("total_hours")

dept_ot_hours = df_th_hourly[ot_mask].groupby("Dist Department Code")["Hours"].sum().rename("ot_hours")

dept_headcount = df_th_hourly.groupby("Dist Department Code")["EE Code"].nunique().rename("headcount")

dept_metrics = pd.concat([
    dept_total_hours,
    dept_ot_hours,
    dept_headcount
], axis=1).fillna(0)

dept_metrics["avg_hours_per_employee"] = (
    dept_metrics["total_hours"] / dept_metrics["headcount"].replace(0, np.nan)
)

dept_metrics["ot_ratio"] = (
    dept_metrics["ot_hours"] / dept_metrics["total_hours"].replace(0, np.nan)
)

dept_metrics["ot_dependency"] = (
    dept_metrics["ot_hours"] /
    (dept_metrics["total_hours"] - dept_metrics["ot_hours"]).replace(0, np.nan)
)

# ----------------------------
# 4. PAY GROUP METRICS (MO vs ML)
# ----------------------------
pg_total_hours = df_th_hourly.groupby("Pay Group")["Hours"].sum().rename("total_hours")

pg_ot_hours = df_th_hourly[ot_mask].groupby("Pay Group")["Hours"].sum().rename("ot_hours")

pg_headcount = df_th_hourly.groupby("Pay Group")["EE Code"].nunique().rename("headcount")

paygroup_metrics = pd.concat([
    pg_total_hours,
    pg_ot_hours,
    pg_headcount
], axis=1).fillna(0)

paygroup_metrics["avg_hours_per_employee"] = (
    paygroup_metrics["total_hours"] / paygroup_metrics["headcount"].replace(0, np.nan)
)

paygroup_metrics["ot_ratio"] = (
    paygroup_metrics["ot_hours"] / paygroup_metrics["total_hours"].replace(0, np.nan)
)

# ----------------------------
# 5. TIME-BASED METRICS
# ----------------------------
daily_metrics = df_th_hourly.groupby("Date").agg(
    total_hours=("Hours", "sum"),
    active_employees=("EE Code", "nunique")
)

daily_metrics["avg_hours_per_employee"] = (
    daily_metrics["total_hours"] / daily_metrics["active_employees"].replace(0, np.nan)
)

daily_ot = df_th_hourly[ot_mask].groupby("Date")["Hours"].sum().rename("ot_hours")

daily_metrics = daily_metrics.join(daily_ot).fillna(0)

# WEEKLY
df_th_hourly["Week"] = df_th_hourly["Date"].dt.isocalendar().week

weekly_metrics = df_th_hourly.groupby("Week").agg(
    total_hours=("Hours", "sum"),
    active_employees=("EE Code", "nunique")
)

weekly_ot = df_th_hourly[ot_mask].groupby(df_th_hourly["Date"].dt.isocalendar().week)["Hours"].sum().rename("ot_hours")

weekly_metrics = weekly_metrics.join(weekly_ot).fillna(0)

weekly_metrics["wow_growth"] = weekly_metrics["total_hours"].pct_change()

# DAY OF WEEK
df_th_hourly["Day_Name"] = df_th_hourly["Date"].dt.day_name()

dow_metrics = df_th_hourly.groupby("Day_Name")["Hours"].mean().rename("avg_hours")

# ----------------------------
# 6. EARNING CODE METRICS
# ----------------------------
earning_metrics = df_th_hourly.groupby(["Earning Code", "Earning Desc"])["Hours"].sum().rename("total_hours").reset_index()

# ----------------------------
# 7. UTILIZATION METRICS
# ----------------------------
active_emp_daily = df_th_hourly.groupby("Date")["EE Code"].nunique()
total_hours_daily = df_th_hourly.groupby("Date")["Hours"].sum()

utilization_metrics = (total_hours_daily / active_emp_daily).rename("hours_per_employee")

# ----------------------------
# 8. CROSS DIMENSIONAL METRICS
# ----------------------------
# Employee × Department
emp_dept_matrix = df_th_hourly.groupby(["EE Code", "Dist Department Code"])["Hours"].sum().reset_index()

# Position × Department
pos_dept_matrix = df_th_hourly.groupby(["Position_Title", "Dist Department Code"])["Hours"].sum().reset_index()

# PayGroup × Department
pg_dept_matrix = df_th_hourly.groupby(["Pay Group", "Dist Department Code"])["Hours"].sum().reset_index()





import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Workforce Analytics Dashboard", layout="wide")

st.title("📊 Workforce Analytics Dashboard")

# ----------------------------
# LOAD DATA (ASSUMES df_th_hourly ALREADY PREPARED)
# ----------------------------
# You can replace this with file upload if needed
# df_th_hourly = pd.read_csv("your_file.csv")

# ----------------------------
# SIDEBAR FILTERS
# ----------------------------
st.sidebar.header("Filters")

if "Date" in df_th_hourly.columns:
    min_date = df_th_hourly["Date"].min()
    max_date = df_th_hourly["Date"].max()
    date_range = st.sidebar.date_input("Date Range", [min_date, max_date])
    df_filtered = df_th_hourly[(df_th_hourly["Date"] >= pd.to_datetime(date_range[0])) & 
                              (df_th_hourly["Date"] <= pd.to_datetime(date_range[1]))]
else:
    df_filtered = df_th_hourly.copy()

# ----------------------------
# KPI METRICS
# ----------------------------
total_hours = df_filtered["Hours"].sum()
ot_hours = df_filtered[df_filtered["Earning Desc"].str.contains("OT", case=False, na=False)]["Hours"].sum()
headcount = df_filtered["EE Code"].nunique()

col1, col2, col3 = st.columns(3)
col1.metric("Total Hours", f"{total_hours:,.0f}")
col2.metric("OT Hours", f"{ot_hours:,.0f}")
col3.metric("Headcount", f"{headcount}")

# ----------------------------
# DAILY TREND
# ----------------------------
st.subheader("📈 Daily Hours Trend")

daily = df_filtered.groupby("Date")["Hours"].sum().reset_index()
fig_daily = px.line(daily, x="Date", y="Hours", title="Daily Total Hours")
st.plotly_chart(fig_daily, use_container_width=True)

# ----------------------------
# WEEKLY TREND
# ----------------------------
st.subheader("📊 Weekly Trend")

df_filtered["Week"] = df_filtered["Date"].dt.isocalendar().week
weekly = df_filtered.groupby("Week")["Hours"].sum().reset_index()
fig_weekly = px.bar(weekly, x="Week", y="Hours", title="Weekly Hours")
st.plotly_chart(fig_weekly, use_container_width=True)

# ----------------------------
# POSITION ANALYSIS
# ----------------------------
st.subheader("👔 Position Analysis")

pos = df_filtered.groupby("Position_Title")["Hours"].sum().reset_index().sort_values(by="Hours", ascending=False)
fig_pos = px.bar(pos, x="Hours", y="Position_Title", orientation="h", title="Hours by Position")
st.plotly_chart(fig_pos, use_container_width=True)

# ----------------------------
# DEPARTMENT ANALYSIS
# ----------------------------
st.subheader("🏢 Department Analysis")

dept = df_filtered.groupby("Dist Department Code")["Hours"].sum().reset_index()
fig_dept = px.pie(dept, names="Dist Department Code", values="Hours", title="Department Share")
st.plotly_chart(fig_dept, use_container_width=True)

# ----------------------------
# OT ANALYSIS
# ----------------------------
st.subheader("⏱️ Overtime Analysis")

ot_df = df_filtered[df_filtered["Earning Desc"].str.contains("OT", case=False, na=False)]
ot_by_emp = ot_df.groupby("EE Code")["Hours"].sum().reset_index().sort_values(by="Hours", ascending=False).head(10)

fig_ot = px.bar(ot_by_emp, x="EE Code", y="Hours", title="Top 10 Employees by OT")
st.plotly_chart(fig_ot, use_container_width=True)

# ----------------------------
# DAY OF WEEK ANALYSIS
# ----------------------------
st.subheader("📅 Day of Week Analysis")

df_filtered["Day_Name"] = df_filtered["Date"].dt.day_name()
dow = df_filtered.groupby("Day_Name")["Hours"].mean().reset_index()

fig_dow = px.bar(dow, x="Day_Name", y="Hours", title="Average Hours by Day")
st.plotly_chart(fig_dow, use_container_width=True)

# ----------------------------
# UTILIZATION
# ----------------------------
st.subheader("⚙️ Utilization")

util = df_filtered.groupby("Date").agg(total_hours=("Hours", "sum"), employees=("EE Code", "nunique"))
util["hours_per_employee"] = util["total_hours"] / util["employees"]
util = util.reset_index()

fig_util = px.line(util, x="Date", y="hours_per_employee", title="Hours per Employee")
st.plotly_chart(fig_util, use_container_width=True)

# ----------------------------
# RAW DATA VIEW
# ----------------------------
st.subheader("📄 Raw Data")
st.dataframe(df_filtered.head(100))
