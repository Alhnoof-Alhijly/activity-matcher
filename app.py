# import streamlit as st
# import pandas as pd
# import re
# import os
# from fuzzywuzzy import fuzz
# # إعداد واجهة Streamlit
# st.title("مطابقة أنشطة الأعضاء")
# st.write("قم برفع ملف الانشطه ")

# # قراءة ملف النشاطات الثابت
# activities_file = "activities.xlsx"
# if not os.path.exists(activities_file):
#     st.error("⚠️ ملف activities.xlsx غير موجود في المجلد! تأكد من وضعه في نفس مجلد التطبيق.")
#     st.stop()

# activities_df = pd.read_excel(activities_file)

# # رفع ملف الأوصاف
# uploaded_file = st.file_uploader("اختر ملف الانشطه الغير مصنفه  (descriptions.xlsx)", type=["xlsx"])
# if uploaded_file is not None:
#     # قراءة ملف الأوصاف المرفوع
#     descriptions_df = pd.read_excel(uploaded_file)

#     # أسماء الأعمدة
#     activities_name_col = 'Name'         # اسم العمود في ملف النشاطات
#     activities_code_col = 'Code'        # الرمز في ملف النشاطات
#     descriptions_col = 'النشاط الرئيسي' # اسم العمود في ملف الأوصاف

#     # 1. تنظيف النص
#     def normalize(text):
#         if pd.isnull(text):
#             return ""
#         text = str(text)
#         text = re.sub(r'[()\[\]{}ـ،؛:،\.\"\'“”]', '', text)
#         text = text.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
#         text = text.replace("ة", "ه")
#         text = text.replace("ى", "ي")
#         text = text.strip().lower()
#         return text

#     # 2. تجهيز القاموس
#     activity_dict = {}
#     activity_set = set()
#     for i, row in activities_df.iterrows():
#         act_name = normalize(row[activities_name_col])
#         activity_dict[act_name] = row[activities_code_col]
#         activity_set.add(act_name)

#     # # 3. دالة المطابقة
#     # def smart_match(description):
#     #     if pd.isnull(description):
#     #         return "", "البيانات فارغة", ""
#     #     norm_desc = normalize(description)
#     #     matches = []
#     #     for activity in activity_set:
#     #         if activity in norm_desc:
#     #             matches.append(str(activity_dict[activity]))
#     #     return ", ".join(matches)
#     # 3. دالة المطابقة مع تحليل التشابه
#     def smart_match(description):
#         if pd.isnull(description):
#             return "", "البيانات فارغة", ""
#         norm_desc = normalize(description)
#         matches = []
#         for activity in activity_set:
#             if activity in norm_desc:
#                 matches.append(str(activity_dict[activity]))
#         if matches:
#             return ", ".join(matches), "", ""
        
#         # إذا لم يتم العثور على تطابق، تحقق من التشابه
#         best_match = ""
#         best_score = 0
#         for activity in activity_set:
#             score = fuzz.partial_ratio(norm_desc, activity)
#             if score > best_score and score >= 80:  # عتبة التشابه 80%
#                 best_score = score
#                 best_match = activity
#         if best_match:
#             return "", f"تشابه جزئي (هل تقصد: {best_match}؟)", best_match
#         return "", "لا توجد بيانات مطابقة في ملف النشاطات", ""
    
   
#     # 4. تطبيق المطابقة على الوصف
#     descriptions_df[["Matched Codes", "سبب عدم المطابقة", "اقتراح"]] = descriptions_df[descriptions_col].apply(
#         lambda x: pd.Series(smart_match(x))
#     )


#     # 5. تحديد إذا تم المطابقة أم لا
#     descriptions_df["Matched?"] = descriptions_df["Matched Codes"].apply(lambda x: "✔️" if x else "❌")

#     # 6. إحصائيات المطابقة
#     total = len(descriptions_df)
#     matched = (descriptions_df["Matched?"] == "✔️").sum()
#     unmatched = (descriptions_df["Matched?"] == "❌").sum()

#     matched_pct = (matched / total) * 100
#     unmatched_pct = (unmatched / total) * 100

#     # 7. حساب عدد الأنشطة المطابقة داخل كل وصف
#     descriptions_df["matched_count"] = descriptions_df["Matched Codes"].apply(
#         lambda x: len(str(x).split(",")) if x else 0
#     )

#     # 8. حساب إحصائيات لكل عضو (رقم العضوية)
#     member_stats = descriptions_df.groupby("رقم العضوية").agg(
#         total_activities=('matched_count', 'sum'),
#         matched_activities=('Matched Codes', lambda x: sum([len(str(codes).split(",")) for codes in x if codes != ""])),
#         unmatched_activities=('Matched?', lambda x: sum([1 for val in x if val == "❌"])),
#     ).reset_index()

#     # 9. إضافة عمود نسبة المطابقة
#     member_stats["matching_percentage"] = (member_stats["matched_activities"] /
#                                          (member_stats["matched_activities"] + member_stats["unmatched_activities"])) * 100
#     member_stats["matching_percentage"] = member_stats["matching_percentage"].round(2)

#     # 10. إضافة عمود "مجموع الأنشطة الفعلية"
#     descriptions_df["مجموع الأنشطة الفعلية"] = descriptions_df["matched_count"]

#     # 11. تحويل رقم العضوية إلى نص
#     descriptions_df["رقم العضوية"] = descriptions_df["رقم العضوية"].astype(str)

#     # 12. حفظ النتائج في ملف Excel
#     output_file = "matched_activitiestotel.xlsx"
#     descriptions_df.to_excel(output_file, index=False)

#    # عرض الإحصائيات
#     st.subheader("📊 إحصائيات الأعضاء")
#     st.dataframe(member_stats)

#     st.subheader("📊 نتائج المطابقة (أول 20 سجل)")
#     st.dataframe(descriptions_df.head(20))

#     # عرض الأوصاف غير المطابقة
#     unmatched_df = descriptions_df[descriptions_df["Matched?"] == "❌"][
#         ["رقم العضوية", descriptions_col, "سبب عدم المطابقة", "اقتراح"]
#     ]
#     st.subheader("📋 الأوصاف غير المطابقة")
#     if not unmatched_df.empty:
#         st.dataframe(unmatched_df)
#     else:
#         st.success("🎉 جميع الأوصاف تمت مطابقتها!")

#     st.write(f"📊 إجمالي الوصف: {total}")
#     st.write(f"✅ تم مطابقة: {matched} ({matched_pct:.2f}%)")
#     st.write(f"❌ لم يتم مطابقة: {unmatched} ({unmatched_pct:.2f}%)")
#     st.success(f"✅ تم حفظ الملف: {output_file}")

#     # زر لتحميل ملف الإخراج
#     with open(output_file, "rb") as file:
#         st.download_button(
#             label="📥 تحميل ملف النتائج (matched_activitiestotel.xlsx)",
#             data=file,
#             file_name=output_file,
#             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#         )
# else:
#     st.info("↥ الرجاء رفع ملف الأوصاف (descriptions.xlsx) للبدء.")

import streamlit as st
import pandas as pd
import os
from matching import normalize, smart_match
from stats import generate_statistics
from utils import prepare_activity_dict

# إعداد واجهة Streamlit
st.title("مطابقة الأنشطة مع دليل التصنيف")
st.write("قم برفع ملف الأنشطة لمعالجة البيانات وإنتاج ملفات النتائج.")

# قراءة ملف النشاطات الثابت
activities_file = "activities.xlsx"
if not os.path.exists(activities_file):
    st.error("⚠️ ملف activities.xlsx غير موجود في المجلد! تأكد من وضعه في نفس مجلد التطبيق.")
    st.stop()

activities_df = pd.read_excel(activities_file)

# أسماء الأعمدة
activities_name_col = 'Name'
activities_code_col = 'Code'
descriptions_col = 'النشاط الرئيسي'

# تجهيز القاموس
activity_dict, activity_set = prepare_activity_dict(activities_df, activities_name_col, activities_code_col)

# تخزين الحالة بين التفاعلات
if 'descriptions_df' not in st.session_state:
    st.session_state.descriptions_df = None
if 'unmatched_df' not in st.session_state:
    st.session_state.unmatched_df = None
if 'updated' not in st.session_state:
    st.session_state.updated = False

# رفع ملف الأوصاف
st.info("📄 يجب أن يحتوي الملف على الأعمدة التالية بالتحديد: 'رقم العضوية' و 'النشاط الرئيسي'.")
uploaded_file = st.file_uploader("اختر ملف الأوصاف", type=["xlsx"])
if uploaded_file is not None and not st.session_state.updated:
    # قراءة ملف الأوصاف
    descriptions_df = pd.read_excel(uploaded_file)
    st.session_state.descriptions_df = descriptions_df.copy()

    # التأكد من وجود الأعمدة المطلوبة
    required_columns = ["رقم العضوية", descriptions_col]
    missing_columns = [col for col in required_columns if col not in descriptions_df.columns]
    if missing_columns:
        st.error(f"⚠️ الملف لا يحتوي على الأعمدة التالية: {', '.join(missing_columns)}")
        st.stop()

    # حساب عدد الأنشطة لكل عضو
    member_activity_counts = descriptions_df.groupby("رقم العضوية").size().to_dict()

    # تطبيق المطابقة
    match_results = descriptions_df.apply(
        lambda row: smart_match(
            row[descriptions_col],
            activity_set,
            activity_dict,
            top_n=member_activity_counts.get(row["رقم العضوية"], 1)
        ), axis=1
    )
    descriptions_df["Matched Codes"] = match_results.apply(lambda x: x[0])
    descriptions_df["سبب عدم المطابقة"] = match_results.apply(lambda x: x[1])
    descriptions_df["اقتراح"] = match_results.apply(lambda x: x[2])
    descriptions_df["Matched?"] = descriptions_df["Matched Codes"].apply(lambda x: "✔️" if x else "❌")
    descriptions_df["matched_count"] = descriptions_df["Matched Codes"].apply(
        lambda x: len(str(x).split(",")) if x else 0
    )
    descriptions_df["مجموع الأنشطة الفعلية"] = descriptions_df["matched_count"]
    descriptions_df["رقم العضوية"] = descriptions_df["رقم العضوية"].astype(str)

    # إنشاء عمود "المطابقة بنسبة"
    descriptions_df["المطابقة بنسبة"] = descriptions_df.apply(
        lambda row: "100%" if row["Matched Codes"] and not row["سبب عدم المطابقة"].startswith("تشابه جزئي") else (
            f"≥80% (التطابق الجزئي: {row['اقتراح']})" if row["سبب عدم المطابقة"].startswith("تشابه جزئي") else "0%"
        ), axis=1
    )

    # دمج الاقتراحات تلقائيًا
    descriptions_df_updated = descriptions_df.copy()
    suggested_matches = []
    for idx, row in descriptions_df_updated.iterrows():
        if row["Matched?"] == "❌" and row.get("اقتراح", ""):
            suggestions = row.get("اقتراح", "").split("; ") if row.get("اقتراح", "") else []
            if suggestions:
                codes = []
                for suggestion in suggestions:
                    suggestion_text = suggestion.split(" (تشابه:")[0]
                    norm_suggestion = normalize(suggestion_text)
                    if norm_suggestion in activity_dict:
                        codes.append(str(activity_dict[norm_suggestion]))
                if codes:
                    descriptions_df_updated.loc[idx, "Matched Codes"] = ", ".join(codes)
                    descriptions_df_updated.loc[idx, "سبب عدم المطابقة"] = "تمت المطابقة بناءً على الاقتراحات (تشابه ≥80%)"
                    descriptions_df_updated.loc[idx, "Matched?"] = "✔️"
                    descriptions_df_updated.loc[idx, "matched_count"] = len(codes)
                    descriptions_df_updated.loc[idx, "مجموع الأنشطة الفعلية"] = len(codes)
                    descriptions_df_updated.loc[idx, "المطابقة بنسبة"] = f"≥80% (التطابق الجزئي: {row['اقتراح']})"
                    suggested_matches.append({
                        "رقم العضوية": row["رقم العضوية"],
                        descriptions_col: row[descriptions_col],
                        "Matched Codes": ", ".join(codes),
                        "تشابه": "≥80%"
                    })

    descriptions_df = descriptions_df_updated

    # إنتاج ملف المطابقة الكاملة (100%)
    exact_matches_df = descriptions_df[descriptions_df["المطابقة بنسبة"] == "100%"][
        ["رقم العضوية", descriptions_col, "Matched Codes", "matched_count"]
    ]
    exact_matches_df["Matched Codes"] = exact_matches_df["Matched Codes"].astype(str).str.split(",").apply(
        lambda x: ",".join([str(code).strip().zfill(6) for code in x])
    )
    exact_matches_file = "exact_matches.xlsx"
    exact_matches_df.to_excel(exact_matches_file, index=False)

    # إنتاج ملف المطابقة الجزئية (≥80%)
    suggested_matches_df = pd.DataFrame(suggested_matches)
    suggested_matches_file = "suggested_matches_80.xlsx"
    if not suggested_matches_df.empty:
        suggested_matches_df.to_excel(suggested_matches_file, index=False)

    # إنتاج ملف الأوصاف غير المطابقة
    unmatched_columns = ["رقم العضوية", descriptions_col, "سبب عدم المطابقة", "اقتراح"]
    unmatched_descriptions_df = descriptions_df[descriptions_df["Matched?"] == "❌"][unmatched_columns]
    unmatched_descriptions_file = "unmatched_descriptions.xlsx"
    if not unmatched_descriptions_df.empty:
        unmatched_descriptions_df.to_excel(unmatched_descriptions_file, index=False)

    # إنتاج ملف النتائج النهائية (100% + ≥80%)
    final_results_df = descriptions_df[descriptions_df["Matched?"] == "✔️"][
        ["رقم العضوية", descriptions_col, "Matched Codes", "المطابقة بنسبة", "matched_count"]
    ]
    final_results_df = final_results_df.rename(columns={"matched_count": "عدد الأنشطة المطابقة"})
    final_results_file = "final_results.xlsx"
    final_results_df.to_excel(final_results_file, index=False)

    # إنتاج ملف membership_matched_codes.xlsx (100% فقط)
    membership_matched_codes = []
    for _, row in exact_matches_df.iterrows():
        membership_number = row["رقم العضوية"]
        matched_codes = str(row["Matched Codes"]).split(",") if row["Matched Codes"] else []
        for code in matched_codes:
            code = str(code).strip().zfill(6)
            membership_matched_codes.append({
                "Membership Number": membership_number,
                "Matched Code": code
            })

    membership_matched_codes_df = pd.DataFrame(membership_matched_codes)
    membership_matched_codes_file = "membership_matched_codes.xlsx"
    if not membership_matched_codes_df.empty:
        membership_matched_codes_df.to_excel(membership_matched_codes_file, index=False)

    # إحصائيات الأعضاء
    member_stats = generate_statistics(descriptions_df)

    # إحصائيات المطابقة
    total = len(descriptions_df)
    matched_100 = len(descriptions_df[descriptions_df["المطابقة بنسبة"] == "100%"])
    matched_80 = len(descriptions_df[descriptions_df["المطابقة بنسبة"].str.startswith("≥80%")])
    unmatched = len(descriptions_df[descriptions_df["Matched?"] == "❌"])
    matched_100_pct = (matched_100 / total) * 100 if total > 0 else 0
    matched_80_pct = (matched_80 / total) * 100 if total > 0 else 0
    merged_pct = ((matched_100 + matched_80) / total) * 100 if total > 0 else 0

    # التحقق من صحة النسب
    if matched_100_pct + matched_80_pct + (unmatched / total * 100) > 100.01:  # السماح بتفاوت طفيف بسبب التقريب
        st.warning("⚠️ مجموع النسب يتجاوز 100%! يتم إعادة الحساب.")
        total_pct = matched_100_pct + matched_80_pct + (unmatched / total * 100)
        matched_100_pct = (matched_100_pct / total_pct) * 100
        matched_80_pct = (matched_80_pct / total_pct) * 100
        unmatched_pct = ((unmatched / total * 100) / total_pct) * 100
    else:
        unmatched_pct = (unmatched / total) * 100 if total > 0 else 0

    # إحصائيات مفصلة
    total_members = descriptions_df["رقم العضوية"].nunique()
    total_activities = descriptions_df["matched_count"].sum()
    avg_activities_per_member = total_activities / total_members if total_members > 0 else 0

    # عرض النتائج
    st.subheader("📊 إحصائيات المطابقة")
    st.write(f"✅ المطابقة 100%: {matched_100} ({matched_100_pct:.2f}%)")
    st.write(f"✅ المطابقة ≥80%: {matched_80} ({matched_80_pct:.2f}%)")
    st.write(f"✅ إجمالي الدمج (100% + ≥80%): {(matched_100 + matched_80)} ({merged_pct:.2f}%)")
    st.write(f"❌ غير مطابق: {unmatched} ({unmatched_pct:.2f}%)")

    st.subheader("📊 إحصائيات مفصلة")
    st.write(f"👥 عدد الأعضاء: {total_members}")
    st.write(f"🎯 إجمالي الأنشطة المطابقة: {total_activities}")
    st.write(f"📈 متوسط الأنشطة لكل عضو: {avg_activities_per_member:.2f}")

    st.subheader("📋 جدول جميع الأعضاء مع أنشطتهم")
    st.dataframe(member_stats)

    st.subheader("📋 جدول المطابقة 100%")
    if not exact_matches_df.empty:
        st.dataframe(exact_matches_df)
    else:
        st.write("⚠️ لا توجد مطابقات بنسبة 100%.")

    st.subheader("📋 جدول المطابقة ≥80%")
    if not suggested_matches_df.empty:
        st.dataframe(suggested_matches_df)
    else:
        # تصفية بديلة لعرض المطابقات ≥80% من descriptions_df
        temp_80_df = descriptions_df[descriptions_df["المطابقة بنسبة"].str.startswith("≥80%")][
            ["رقم العضوية", descriptions_col, "Matched Codes", "اقتراح"]
        ]
        if not temp_80_df.empty:
            st.write("ℹ️ عرض المطابقات ≥80% من البيانات الرئيسية بسبب مشكلة في جدول الاقتراحات:")
            st.dataframe(temp_80_df)
            # حفظ المطابقات ≥80% في ملف
            temp_80_df.to_excel(suggested_matches_file, index=False)
        else:
            st.warning(f"⚠️ لا توجد مطابقات بنسبة ≥80% في الجدول، رغم وجود {matched_80} مطابقة في الإحصائيات. تحقق من عمود 'اقتراح'.")

    st.subheader("📋 جدول الأوصاف غير المطابقة")
    if not unmatched_descriptions_df.empty:
        st.dataframe(unmatched_descriptions_df)
    else:
        st.success("🎉 جميع الأوصاف تمت مطابقتها بنسبة 100% أو ≥80%!")

    st.subheader("📋 جدول النتائج النهائية ")
    st.dataframe(final_results_df)

    # جدول المطابقة 100% (كل عضو مع نشاط واحد)
    st.subheader("📋 جدول المطابقة 100% (كل عضو مع نشاط واحد)")
    st.dataframe(membership_matched_codes_file)

    # تحميل الملفات
    st.subheader("📥 تحميل الملفات")
    for file_name, label in [
        (exact_matches_file, "تحميل ملف المطابقة الكاملة (exact_matches.xlsx)"),
        (suggested_matches_file, "تحميل ملف المطابقة الجزئية ≥80% (suggested_matches_80.xlsx)"),
        (unmatched_descriptions_file, "تحميل ملف الأوصاف غير المطابقة (unmatched_descriptions.xlsx)"),
        (final_results_file, "تحميل ملف النتائج النهائية (final_results.xlsx)"),
        (membership_matched_codes_file, "تحميل ملف أكواد الأعضاء المطابقة 100% (membership_matched_codes.xlsx)")
    ]:
        if os.path.exists(file_name):
            with open(file_name, "rb") as file:
                st.download_button(
                    label=label,
                    data=file,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning(f"⚠️ الملف {file_name} لم يتم إنشاؤه. تحقق من البيانات.")

    st.success("✅ تمت معالجة البيانات وحفظ جميع الملفات!")
    st.session_state.updated = True
    st.session_state.descriptions_df = descriptions_df

elif st.session_state.updated:
    st.info("✅ تم تحديث النتائج. يمكنك رفع ملف جديد لإعادة البدء.")
else:
    st.info("↥ الرجاء رفع ملف الأوصاف للبدء.")