import streamlit as st
import pandas as pd
import numpy as np
import os
from matching import normalize, smart_match
from stats import generate_statistics, generate_activity_recommendations
from utils import prepare_activity_dict

# إعداد واجهة Streamlit
st.title("مطابقة الأنشطة مع دليل التصنيف")
st.write("اختر بين رفع ملف الأوصاف أو إدخال نص نشاط واحد لتحليل التطابق والتوصيات.")

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
if 'single_activity_result' not in st.session_state:
    st.session_state.single_activity_result = None

# خيار التحليل
analysis_option = st.radio("اختر نوع التحليل:", ["رفع ملف الأوصاف", "إدخال نص نشاط واحد"])

# خيار 1: رفع ملف الأوصاف
if analysis_option == "رفع ملف الأوصاف":
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
        if matched_100_pct + matched_80_pct + (unmatched / total * 100) > 100.01:
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

        # توليد جدول التوصيات
        recommendations_df, summary = generate_activity_recommendations(descriptions_df, activity_dict, min_support=0.01, min_confidence=0.5)
        recommendations_file = "activity_recommendations.xlsx"
        if not recommendations_df.empty:
            recommendations_df.to_excel(recommendations_file, index=False)

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
            temp_80_df = descriptions_df[descriptions_df["المطابقة بنسبة"].str.startswith("≥80%")][
                ["رقم العضوية", descriptions_col, "Matched Codes", "اقتراح"]
            ]
            if not temp_80_df.empty:
                st.write("ℹ️ عرض المطابقات ≥80% من البيانات الرئيسية بسبب مشكلة في جدول الاقتراحات:")
                st.dataframe(temp_80_df)
                temp_80_df.to_excel(suggested_matches_file, index=False)
            else:
                st.warning(f"⚠️ لا توجد مطابقات بنسبة ≥80% في الجدول، رغم وجود {matched_80} مطابقة في الإحصائيات. تحقق من عمود 'اقتراح'.")

        st.subheader("📋 جدول الأوصاف غير المطابقة")
        if not unmatched_descriptions_df.empty:
            st.dataframe(unmatched_descriptions_df)
        else:
            st.success("🎉 جميع الأوصاف تمت مطابقتها بنسبة 100% أو ≥80%!")

        st.subheader("📋 جدول النتائج النهائية")
        st.dataframe(final_results_df)

        st.subheader("📋 جدول المطابقة 100% (كل عضو مع نشاط واحد)")
        st.dataframe(membership_matched_codes_df)

        st.subheader("📋 جدول توصيات الأنشطة")
        if not recommendations_df.empty:
            st.write("إذا كنت تمارس نشاطًا معينًا، هذه هي الأنشطة التي غالبًا تكون مرتبطة به:")
            st.markdown(summary)
            st.dataframe(recommendations_df)
        else:
            st.warning("⚠️ لا توجد توصيات متاحة بناءً على البيانات الحالية. حاول تعديل الحد الأدنى للدعم أو الثقة.")

        # تحميل الملفات
        st.subheader("📥 تحميل الملفات")
        for file_name, label in [
            (exact_matches_file, "تحميل ملف المطابقة الكاملة (exact_matches.xlsx)"),
            (suggested_matches_file, "تحميل ملف المطابقة الجزئية ≥80% (suggested_matches_80.xlsx)"),
            (unmatched_descriptions_file, "تحميل ملف الأوصاف غير المطابقة (unmatched_descriptions.xlsx)"),
            (final_results_file, "تحميل ملف النتائج النهائية (final_results.xlsx)"),
            (membership_matched_codes_file, "تحميل ملف أكواد الأعضاء المطابقة 100% (membership_matched_codes.xlsx)"),
            (recommendations_file, "تحميل ملف توصيات الأنشطة (activity_recommendations.xlsx)")
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

# خيار 2: إدخال نص نشاط واحد
elif analysis_option == "إدخال نص نشاط واحد":
    st.info("📝 أدخل نص النشاط لتحليل التطابق والتوصيات. مثال: 'الإنشاءات العامة للمباني غير السكنية، يشمل (المدارس، المستشفيات، الفنادق... إلخ)'")
    target_activity = st.text_input("أدخل نص النشاط:", "")

    if st.button("تحليل النشاط") and target_activity:
        # محاكاة بيانات الأوصاف
        simulated_data = [
            {"رقم العضوية": "203011201438", descriptions_col: target_activity, "Matched Codes": "410025,410026,551011"},
            {"رقم العضوية": "203011201439", descriptions_col: "إنشاء المدارس والمستشفيات", "Matched Codes": "410025,410026"},
            {"رقم العضوية": "203011201440", descriptions_col: "إنشاء الفنادق والمباني غير السكنية", "Matched Codes": "551011,410025"}
        ]
        descriptions_df = pd.DataFrame(simulated_data)

        # تطبيق المطابقة
        match_results = descriptions_df.apply(
            lambda row: smart_match(
                row[descriptions_col],
                activity_set,
                activity_dict,
                top_n=3  # السماح بمطابقة متعددة (مدارس، مستشفيات، فنادق)
            ), axis=1
        )
        descriptions_df["Matched Codes"] = match_results.apply(lambda x: x[0])
        descriptions_df["سبب عدم المطابقة"] = match_results.apply(lambda x: x[1])
        descriptions_df["اقتراح"] = match_results.apply(lambda x: x[2])
        descriptions_df["Matched?"] = descriptions_df["Matched Codes"].apply(lambda x: "✔️" if x else "❌")
        descriptions_df["matched_count"] = descriptions_df["Matched Codes"].apply(
            lambda x: len(str(x).split(",")) if x else 0
        )
        descriptions_df["المطابقة بنسبة"] = descriptions_df.apply(
            lambda row: "100%" if row["Matched Codes"] and not row["سبب عدم المطابقة"].startswith("تشابه جزئي") else (
                f"≥80% (التطابق الجزئي: {row['اقتراح']})" if row["سبب عدم المطابقة"].startswith("تشابه جزئي") else "0%"
            ), axis=1
        )

        # إحصائيات التطابق
        total = len(descriptions_df)
        matched_100 = len(descriptions_df[descriptions_df["المطابقة بنسبة"] == "100%"])
        matched_80 = len(descriptions_df[descriptions_df["المطابقة بنسبة"].str.startswith("≥80%")])
        unmatched = len(descriptions_df[descriptions_df["Matched?"] == "❌"])
        matched_100_pct = (matched_100 / total) * 100 if total > 0 else 0
        matched_80_pct = (matched_80 / total) * 100 if total > 0 else 0
        unmatched_pct = (unmatched / total) * 100 if total > 0 else 0

        # توليد التوصيات
        recommendations_df, summary = generate_activity_recommendations(
            descriptions_df, activity_dict, min_support=0.01, min_confidence=0.5
        )
        single_activity_file = "single_activity_analysis.xlsx"
        if not recommendations_df.empty:
            recommendations_df.to_excel(single_activity_file, index=False)

        # عرض النتائج
        st.subheader("📊 إحصائيات التطابق للنشاط المحدد")
        st.write(f"✅ المطابقة الكاملة (100%): {matched_100} ({matched_100_pct:.2f}%)")
        st.write(f"✅ المطابقة الجزئية (≥80%): {matched_80} ({matched_80_pct:.2f}%)")
        st.write(f"❌ غير مطابق: {unmatched} ({unmatched_pct:.2f}%)")

        st.subheader("📋 تفاصيل المطابقة")
        st.dataframe(descriptions_df[["رقم العضوية", descriptions_col, "Matched Codes", "المطابقة بنسبة", "اقتراح"]])

        st.subheader("📋 جدول توصيات الأنشطة")
        if not recommendations_df.empty:
            st.write("إذا كنت تمارس هذا النشاط، هذه هي الأنشطة التي قد تكون مرتبطة به:")
            # st.markdown(summary)
            # st.dataframe(recommendations_df)
            with open(single_activity_file, "rb") as file:
                st.download_button(
                    label="تحميل ملف تحليل النشاط (single_activity_analysis.xlsx)",
                    data=file,
                    file_name=single_activity_file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("⚠️ لا توجد توصيات متاحة بناءً على البيانات المحاكاة.")

        st.success("✅ تم تحليل النشاط المحدد!")
        st.session_state.single_activity_result = descriptions_df

    elif not target_activity:
        st.info("↥ الرجاء إدخال نص النشاط للبدء.")