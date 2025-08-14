import pandas as pd
import os

# ====== CONFIGURATION ======
input_folder = r"C:\Users\shiva.darshini\OneDrive - OneWorkplace\Chanel\Documents\Data\Cleaned"
output_folder = os.path.join(input_folder, "Cleaned_Output")
os.makedirs(output_folder, exist_ok=True)


# ====== CLEANING FUNCTION ======
def clean_and_optimize(df, source_name):
    # Standardize column names
    df.columns = (
        df.columns.str.strip()
        .str.replace(r"[^0-9a-zA-Z]+", "_", regex=True)
        .str.lower()
    )

    # Add source column
    df["source"] = source_name

    # Standardize date format
    for col in df.columns:
        if "date" in col or "day" in col:
            try:
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d')
            except:
                pass

    # Remove empty columns
    df = df.dropna(axis=1, how='all')

    # Reduce float precision
    for col in df.select_dtypes(include=['float']).columns:
        df.loc[:, col] = df[col].round(4)

    return df


# ====== KPI CALCULATION FUNCTION ======
def calculate_metrics(df):
    # CTR
    if {"clicks", "impressions"}.issubset(df.columns):
        df["ctr_%"] = (df["clicks"] / df["impressions"] * 100).round(2)

    # CPM
    if {"media_cost", "impressions"}.issubset(df.columns):
        df["cpm"] = (df["media_cost"] / df["impressions"] * 1000).round(4)

    # CPC
    if {"media_cost", "clicks"}.issubset(df.columns):
        df["cpc"] = (df["media_cost"] / df["clicks"]).replace([float('inf'), -float('inf')], 0).round(4)

    # CPA
    for conv_col in ["conversions", "web_site_transactions"]:
        if {"media_cost", conv_col}.issubset(df.columns):
            df[f"cpa_{conv_col}"] = (df["media_cost"] / df[conv_col]).replace([float('inf'), -float('inf')], 0).round(4)

    # Bounce Rate (GA)
    if {"visitor_bounce", "visits"}.issubset(df.columns):
        df["bounce_rate_%"] = (df["visitor_bounce"] / df["visits"] * 100).round(2)

    # Pages per Visit (GA)
    if {"page_views", "visits"}.issubset(df.columns):
        df["pages_per_visit"] = (df["page_views"] / df["visits"]).round(2)

    # Avg Time on Page (GA)
    if {"time_on_site_total", "page_views"}.issubset(df.columns):
        df["avg_time_on_page_sec"] = (df["time_on_site_total"] / df["page_views"]).round(2)

    # Exit Rate (GA)
    if {"page_exits", "page_views"}.issubset(df.columns):
        df["exit_rate_%"] = (df["page_exits"] / df["page_views"] * 100).round(2)

    # Video Completion Rate (FB or Video Ads)
    if {"video_fully_played", "video_views"}.issubset(df.columns):
        df["video_completion_rate_%"] = (df["video_fully_played"] / df["video_views"] * 100).round(2)

    return df


# ====== PROCESS FILES ======
for file in os.listdir(input_folder):
    if file.lower().endswith(".xlsx") and "cleaned" in file.lower():
        file_path = os.path.join(input_folder, file)
        source_name = os.path.splitext(file)[0]

        # Load, clean, and calculate metrics
        df = pd.read_excel(file_path)
        df = clean_and_optimize(df, source_name)
        df = calculate_metrics(df)

        # Save cleaned + metrics file
        output_path = os.path.join(output_folder, f"{source_name}_Cleaned_With_KPIs.xlsx")
        df.to_excel(output_path, index=False)

print(f"✅ Cleaning + KPI calculation complete! Files saved in: {output_folder}")
