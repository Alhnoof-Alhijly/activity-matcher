import pandas as pd

def generate_statistics(descriptions_df):
    member_stats = descriptions_df.groupby("رقم العضوية").agg(
        total_activities=('matched_count', 'sum'),
        matched_activities=('Matched Codes', lambda x: sum([len(str(codes).split(",")) for codes in x if codes != ""])),
        unmatched_activities=('Matched?', lambda x: sum([1 for val in x if val == "❌"]))
    ).reset_index()
    member_stats["matching_percentage"] = (
        member_stats["matched_activities"] / 
        (member_stats["matched_activities"] + member_stats["unmatched_activities"])
    ) * 100
    member_stats["matching_percentage"] = member_stats["matching_percentage"].round(2)
    return member_stats