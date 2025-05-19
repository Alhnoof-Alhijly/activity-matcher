import pandas as pd
import re

# قراءة ملفات Excel مباشرة من المجلد
activities_df = pd.read_excel('activities.xlsx')
descriptions_df = pd.read_excel('descriptions.xlsx')

# عرض البيانات
print("✅ جدول النشاطات الأساسية:")
print(activities_df)
print("✅ جدول أوصاف النشاطات:")
print(descriptions_df)

# أسماء الأعمدة
activities_name_col = 'Name'         # اسم العمود في ملف النشاطات
activities_code_col = 'Code'        # الرمز في ملف النشاطات
descriptions_col = 'النشاط الرئيسي' # اسم العمود في ملف الأوصاف

# 1. تنظيف النص
def normalize(text):
    if pd.isnull(text):
        return ""
    text = str(text)
    text = re.sub(r'[()\[\]{}ـ،؛:،\.\"\'“”]', '', text)
    text = text.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    text = text.replace("ة", "ه")
    text = text.replace("ى", "ي")
    text = text.strip().lower()
    return text

# 2. تجهيز القاموس
activity_dict = {}
activity_set = set()
for i, row in activities_df.iterrows():
    act_name = normalize(row[activities_name_col])
    activity_dict[act_name] = row[activities_code_col]
    activity_set.add(act_name)

# 3. دالة المطابقة
def smart_match(description):
    if pd.isnull(description):
        return ""
    norm_desc = normalize(description)
    matches = []
    for activity in activity_set:
        if activity in norm_desc:
            matches.append(str(activity_dict[activity]))
    return ", ".join(matches)

# 4. تطبيق المطابقة على الوصف
descriptions_df["Matched Codes"] = descriptions_df[descriptions_col].apply(smart_match)

# 5. تحديد إذا تم المطابقة أم لا
descriptions_df["Matched?"] = descriptions_df["Matched Codes"].apply(lambda x: "✔️" if x else "❌")

# 6. إحصائيات المطابقة
total = len(descriptions_df)
matched = (descriptions_df["Matched?"] == "✔️").sum()
unmatched = (descriptions_df["Matched?"] == "❌").sum()

matched_pct = (matched / total) * 100
unmatched_pct = (unmatched / total) * 100

# 7. حساب عدد الأنشطة المطابقة داخل كل وصف
descriptions_df["matched_count"] = descriptions_df["Matched Codes"].apply(
    lambda x: len(str(x).split(",")) if x else 0
)

# 8. حساب إحصائيات لكل عضو (رقم العضوية)
member_stats = descriptions_df.groupby("رقم العضوية").agg(
    total_activities=('matched_count', 'sum'),
    matched_activities=('Matched Codes', lambda x: sum([len(str(codes).split(",")) for codes in x if codes != ""])),
    unmatched_activities=('Matched?', lambda x: sum([1 for val in x if val == "❌"])),
).reset_index()

# 9. إضافة عمود نسبة المطابقة
member_stats["matching_percentage"] = (member_stats["matched_activities"] /
                                      (member_stats["matched_activities"] + member_stats["unmatched_activities"])) * 100
member_stats["matching_percentage"] = member_stats["matching_percentage"].round(2)

# 10. إضافة عمود "مجموع الأنشطة الفعلية"
descriptions_df["مجموع الأنشطة الفعلية"] = descriptions_df["matched_count"]

# 11. عرض الإحصائيات
print("\n📊 إحصائيات الأعضاء:")
print(member_stats)
print("\n📊 نتائج المطابقة (أول 20 سجل):")
print(descriptions_df.head(20))
print(f"📊 إجمالي الوصف: {total}")
print(f"✅ تم مطابقة: {matched} ({matched_pct:.2f}%)")
print(f"❌ لم يتم مطابقة: {unmatched} ({unmatched_pct:.2f}%)")

# 12. تحويل رقم العضوية إلى نص
descriptions_df["رقم العضوية"] = descriptions_df["رقم العضوية"].astype(str)

# 13. حفظ النتائج في ملف Excel
descriptions_df.to_excel("matched_activitiestotel.xlsx", index=False)
print("✅ تم حفظ الملف: matched_activitiestotel.xlsx")