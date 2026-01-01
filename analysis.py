import pandas as pd
import re
# -------------------------------------------------
# Helpers
# -------------------------------------------------
def to_percent(x):
    """
    0.17  -> 17
    -0.92 -> -92
    '17%' -> 17
    """
    if pd.isna(x):
        return None
    if isinstance(x, str):
        s = x.strip().replace("%", "").replace(",", ".")
        try:
            return float(s)
        except:
            return None
    try:
        v = float(x)
    except:
        return None
    if -1 <= v <= 1:
        return v * 100
    return v
def detect_week_from_filename(filename: str):
    """
    Top_100_progress_OHL_w52.xlsx -> ("52", "51")
    Top_100_progress_OHL_w01.xlsx -> ("01", "52")
    """
    m = re.search(r"w(\d{1,2})", filename.lower())
    if not m:
        return None, None
    week = int(m.group(1))
    prev_week = 52 if week == 1 else week - 1
    return str(week), str(prev_week)
# -------------------------------------------------
# Main analysis
# -------------------------------------------------
def analyze(raw_path, progress_path, gms_path, progress_filename):
    # -----------------------------
    # LOAD FILES
    # -----------------------------
    raw_df = pd.read_excel(raw_path)
    progress_df = pd.read_excel(progress_path)
    raw_df.columns = raw_df.columns.str.lower()
    progress_df.columns = progress_df.columns.str.lower()
    # -----------------------------
    # WEEK (AUTOMATIC)
    # -----------------------------
    week, previous_week = detect_week_from_filename(progress_filename)
    if week is None:
        raise ValueError("Week could not be detected from filename")
    current_gms_col = f"gms_{week}"
    prev_gms_col = f"gms_{previous_week}"
    # güvenlik
    for col in [current_gms_col, prev_gms_col]:
        if col not in progress_df.columns:
            raise ValueError(f"Expected column '{col}' not found in progress file")
    # -----------------------------
    # GMS BASELINE (CONGRATS LOGIC)
    # -----------------------------
    progress_df[current_gms_col] = pd.to_numeric(
        progress_df[current_gms_col], errors="coerce"
    )
    progress_df[prev_gms_col] = pd.to_numeric(
        progress_df[prev_gms_col], errors="coerce"
    )
    progress_df["diff"] = (
        progress_df[current_gms_col] - progress_df[prev_gms_col]
    )
    progress_df["sas_flag"] = (
        progress_df["sas"].astype(str).str.lower() == "yes"
    )
    def top_n(df, n=3):
        return (
            df.head(n)[
                ["merchant_name", current_gms_col, prev_gms_col, "diff"]
            ]
            .rename(
                columns={
                    "merchant_name": "sp_name",
                    current_gms_col: f"gms_{week}",
                    prev_gms_col: f"gms_{previous_week}",
                }
            )
            .to_dict("records")
        )
    contributors_df = progress_df[progress_df["diff"] > 0].sort_values(
        "diff", ascending=False
    )
    detractors_df = progress_df[progress_df["diff"] < 0].sort_values("diff")
    contributors = {
        "sas": top_n(contributors_df[contributors_df["sas_flag"]]),
        "non_sas": top_n(contributors_df[~contributors_df["sas_flag"]]),
    }
    detractors = {
        "sas": top_n(detractors_df[detractors_df["sas_flag"]]),
        "non_sas": top_n(detractors_df[~detractors_df["sas_flag"]]),
    }
    # -----------------------------
    # PARITY (%30 / %50)
    # -----------------------------
    parity_df = progress_df[
        ["merchant_name", "selection_parity_comp"]
    ].dropna()
    parity_df["selection_parity_comp"] = parity_df[
        "selection_parity_comp"
    ].apply(to_percent)
    parity_df = parity_df.dropna(subset=["selection_parity_comp"])
    parity_increase = (
        parity_df[parity_df["selection_parity_comp"] >= 30]
        .sort_values("selection_parity_comp", ascending=False)
    )
    parity_decrease = (
        parity_df[parity_df["selection_parity_comp"] <= -50]
        .sort_values("selection_parity_comp")
    )
    # -----------------------------
    # FROM ZERO SELECTION
    # -----------------------------
    zero_selection = []
    if "amazon_ba" in raw_df.columns:
        zero_selection = (
            raw_df[raw_df["amazon_ba"] == 0]["merchant_name"]
            .dropna()
            .unique()
            .tolist()
        )
    # -----------------------------
    # PARITY DATA FOR NEW FORMAT
    # -----------------------------
    # From zero selection text
    if not zero_selection:
        from_zero_text = "N/A."
    else:
        from_zero_text = "\n".join(zero_selection)
    
    # WoW parity increase text
    if parity_increase.empty:
        parity_increase_text = "N/A."
    else:
        increase_lines = []
        for _, r in parity_increase.iterrows():
            increase_lines.append(f"{r['merchant_name']}\t{int(r['selection_parity_comp'])}%")
        parity_increase_text = "\n".join(increase_lines)
    
    # WoW parity decrease - structured data for table
    wow_parity_decrease = {}
    if not parity_decrease.empty:
        for _, r in parity_decrease.iterrows():
            wow_parity_decrease[r['merchant_name']] = f"{int(r['selection_parity_comp'])}%"
    
    # -----------------------------
    # FINAL JSON
    # -----------------------------
    return {
        "week": week,
        "previous_week": previous_week,
        "contributors": contributors,
        "detractors": detractors,
        "from_zero_selection_text": from_zero_text,
        "wow_parity_increase_text": parity_increase_text,
        "wow_parity_decrease": wow_parity_decrease,
    }