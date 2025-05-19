import pandas as pd
import re
from fuzzywuzzy import fuzz

def normalize(text):
    if pd.isnull(text) or str(text).strip().lower() in ['null', '']:
        return ""
    text = str(text)
    text = re.sub(r'[()\[\]{}ـ،؛:،\.\"\'“”]', '', text)
    text = text.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    text = text.replace("ة", "ه").replace("ى", "ي")
    return text.strip().lower()

def smart_match(description, activity_set, activity_dict):
    if pd.isnull(description) or not str(description).strip() or str(description).strip().lower() in ['null', '']:
        return "", "البيانات فارغة", ""
    norm_desc = normalize(description)
    if not norm_desc:
        return "", "البيانات فارغة", ""
    
    # المطابقة الكاملة
    matches = [str(activity_dict[activity]).zfill(6) for activity in activity_set if activity in norm_desc]
    if matches:
        return ", ".join(matches), "", ""
    
    # التطابق الجزئي (≥80%)
    partial_matches = []
    for activity in activity_set:
        score = fuzz.partial_ratio(norm_desc, activity)
        if score >= 80:
            partial_matches.append((activity, activity_dict[activity], score))
    
    if partial_matches:
        # ترتيب الأنشطة حسب درجة التشابه (تنازليًا)
        partial_matches.sort(key=lambda x: x[2], reverse=True)
        codes = [str(match[1]).zfill(6) for match in partial_matches]
        suggestions = [f"{match[0]} (تشابه: {match[2]}%)" for match in partial_matches]
        return ", ".join(codes), "تشابه جزئي", "; ".join(suggestions)
    
    return "", "لا توجد بيانات مطابقة في ملف النشاطات", ""