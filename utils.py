import pandas as pd

from matching import normalize

def prepare_activity_dict(activities_df, name_col, code_col):
    activity_dict = {}
    activity_set = set()
    for i, row in activities_df.iterrows():
        act_name = normalize(row[name_col])
        if act_name:
            # التأكد من أن الكود يُحفظ كنص
            activity_dict[act_name] = str(row[code_col]).zfill(6)  # إضافة أصفار بادئة إذا لزم الأمر
            activity_set.add(act_name)
    return activity_dict, activity_set