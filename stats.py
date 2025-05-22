import pandas as pd
import numpy as np

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

try:
    from mlxtend.frequent_patterns import apriori, association_rules
    MLXTEND_AVAILABLE = True
except ImportError:
    MLXTEND_AVAILABLE = False

def generate_activity_recommendations(descriptions_df, activity_dict, min_support=0.01, min_confidence=0.5):
    """
    تحليل الأنشطة المرتبطة وإنشاء جدول توصيات محسن للقراء غير التقنيين.
    
    Parameters:
    - descriptions_df: DataFrame يحتوي على الأنشطة المطابقة (عمود 'Matched Codes').
    - activity_dict: قاموس يربط بين أسماء الأنشطة (مُعيارة) وأكوادها.
    - min_support: الحد الأدنى للدعم (نسبة التكرار).
    - min_confidence: الحد الأدنى للثقة في التوصيات.
    
    Returns:
    - DataFrame يحتوي على الأكواد، الأسماء، التوصيات، وأعمدة إضافية للتوضيح.
    """
    # تحويل الأكواد إلى قائمة لكل عضو
    member_activities = descriptions_df.groupby("رقم العضوية")["Matched Codes"].apply(
        lambda x: [code.strip() for codes in x if codes for code in str(codes).split(",")]
    ).reset_index()

    # إنشاء قائمة الأنشطة الفريدة
    all_codes = set()
    for codes in member_activities["Matched Codes"]:
        all_codes.update(codes)
    all_codes = sorted(all_codes)

    if not all_codes:
        return pd.DataFrame(columns=["كود النشاط", "اسم النشاط", "كود النشاط الموصى به", 
                                    "النشاط الموصى به", "نسبة التكرار (%)", "قوة الارتباط (%)", 
                                    "أهمية التوصية", "فرصة الأعمال", "النشاط الأكثر ارتباطًا"])

    # إنشاء قاموس عكسي لربط الأكواد بالأسماء
    code_to_name = {str(code).zfill(6): name for name, code in activity_dict.items()}

    if MLXTEND_AVAILABLE:
        # إنشاء مصفوفة ثنائية
        binary_matrix = pd.DataFrame(0, index=member_activities.index, columns=all_codes)
        for idx, codes in member_activities.iterrows():
            for code in codes["Matched Codes"]:
                binary_matrix.loc[idx, code] = 1

        # تطبيق خوارزمية Apriori
        frequent_itemsets = apriori(binary_matrix, min_support=min_support, use_colnames=True)
        if frequent_itemsets.empty:
            return pd.DataFrame(columns=["كود النشاط", "اسم النشاط", "كود النشاط الموصى به", 
                                        "النشاط الموصى به", "نسبة التكرار (%)", "قوة الارتباط (%)", 
                                        "أهمية التوصية", "فرصة الأعمال", "النشاط الأكثر ارتباطًا"])

        # استخراج قواعد الارتباط
        rules = association_rules(frequent_itemsets, metric="confidence", min_threshold=min_confidence)
        if rules.empty:
            return pd.DataFrame(columns=["كود النشاط", "اسم النشاط", "كود النشاط الموصى به", 
                                        "النشاط الموصى به", "نسبة التكرار (%)", "قوة الارتباط (%)", 
                                        "أهمية التوصية", "فرصة الأعمال", "النشاط الأكثر ارتباطًا"])

        # تنسيق جدول التوصيات
        recommendations = []
        for _, rule in rules.iterrows():
            antecedents = list(rule["antecedents"])
            consequents = list(rule["consequents"])
            for antecedent in antecedents:
                for consequent in consequents:
                    antecedent_name = code_to_name.get(antecedent, "غير معروف")
                    consequent_name = code_to_name.get(consequent, "غير معروف")
                    support_pct = round(rule["support"] * 100, 2)
                    confidence_pct = round(rule["confidence"] * 100, 2)
                    
                    # تحديد أهمية التوصية
                    if rule["confidence"] >= 0.85:
                        importance = "عالية"
                        opportunity = "فرصة قوية لتوسيع الأعمال أو تقديم خدمات إضافية"
                    elif rule["confidence"] >= 0.7:
                        importance = "متوسطة"
                        opportunity = "فرصة جيدة للنظر في إضافة هذا النشاط"
                    else:
                        importance = "منخفضة"
                        opportunity = "ارتباط ضعيف، قد يحتاج دراسة إضافية"
                    
                    recommendations.append({
                        "كود النشاط": antecedent,
                        "اسم النشاط": antecedent_name,
                        "كود النشاط الموصى به": consequent,
                        "النشاط الموصى به": consequent_name,
                        "نسبة التكرار (%)": support_pct,
                        "قوة الارتباط (%)": confidence_pct,
                        "أهمية التوصية": importance,
                        "فرصة الأعمال": opportunity
                    })

        recommendations_df = pd.DataFrame(recommendations)
        
        # إضافة عمود "النشاط الأكثر ارتباطًا"
        recommendations_df["النشاط الأكثر ارتباطًا"] = "لا"
        max_confidence_per_activity = recommendations_df.groupby("كود النشاط")["قوة الارتباط (%)"].max()
        for idx, row in recommendations_df.iterrows():
            if row["قوة الارتباط (%)"] == max_confidence_per_activity[row["كود النشاط"]]:
                recommendations_df.loc[idx, "النشاط الأكثر ارتباطًا"] = "نعم"
        
        # ترتيب الجدول حسب قوة الارتباط
        recommendations_df = recommendations_df.sort_values(by="قوة الارتباط (%)", ascending=False)
        
        # إنشاء ملخص نصي
        if not recommendations_df.empty:
            top_recommendation = recommendations_df.iloc[0]
            bottom_recommendation = recommendations_df.iloc[-1]
            summary = (
                f"**أعلى توصية**: إذا كنت تمارس '{top_recommendation['اسم النشاط']}'، يوصى بشدة بـ "
                f"'{top_recommendation['النشاط الموصى به']}' (قوة الارتباط: {top_recommendation['قوة الارتباط (%)']}%، "
                f"نسبة التكرار: {top_recommendation['نسبة التكرار (%)']}%). هذه فرصة قوية لتوسيع الأعمال.\n"
                f"**أقل توصية**: إذا كنت تمارس '{bottom_recommendation['اسم النشاط']}'، فإن التوصية بـ "
                f"'{bottom_recommendation['النشاط الموصى به']}' لها قوة ارتباط أقل "
                f"({bottom_recommendation['قوة الارتباط (%)']}%)، مما قد يتطلب دراسة إضافية."
            )
        else:
            summary = "لا توجد توصيات متاحة بناءً على البيانات الحالية."
        
        return recommendations_df, summary
    
    else:
        # نهج يدوي إذا لم تكن mlxtend متوفرة
        recommendations = []
        total_members = len(member_activities)
        
        # حساب التكرارات لكل زوج من الأكواد
        code_pairs = {}
        for _, row in member_activities.iterrows():
            codes = row["Matched Codes"]
            for i, code1 in enumerate(codes):
                for code2 in codes[i+1:]:
                    pair = tuple(sorted([code1, code2]))
                    code_pairs[pair] = code_pairs.get(pair, 0) + 1

        # تحويل التكرارات إلى توصيات
        for (code1, code2), count in code_pairs.items():
            support = count / total_members
            if support >= min_support:
                code1_count = sum(1 for codes in member_activities["Matched Codes"] if code1 in codes)
                confidence = count / code1_count if code1_count > 0 else 0
                if confidence >= min_confidence:
                    code1_name = code_to_name.get(code1, "غير معروف")
                    code2_name = code_to_name.get(code2, "غير معروف")
                    support_pct = round(support * 100, 2)
                    confidence_pct = round(confidence * 100, 2)
                    
                    # تحديد أهمية التوصية
                    if confidence >= 0.85:
                        importance = "عالية"
                        opportunity = "فرصة قوية لتوسيع الأعمال أو تقديم خدمات إضافية"
                    elif confidence >= 0.7:
                        importance = "متوسطة"
                        opportunity = "فرصة جيدة للنظر في إضافة هذا النشاط"
                    else:
                        importance = "منخفضة"
                        opportunity = "ارتباط ضعيف، قد يحتاج دراسة إضافية"
                    
                    recommendations.append({
                        "كود النشاط": code1,
                        "اسم النشاط": code1_name,
                        "كود النشاط الموصى به": code2,
                        "النشاط الموصى به": code2_name,
                        "نسبة التكرار (%)": support_pct,
                        "قوة الارتباط (%)": confidence_pct,
                        "أهمية التوصية": importance,
                        "فرصة الأعمال": opportunity
                    })

        recommendations_df = pd.DataFrame(recommendations)
        if recommendations_df.empty:
            return pd.DataFrame(columns=["كود النشاط", "اسم النشاط", "كود النشاط الموصى به", 
                                        "النشاط الموصى به", "نسبة التكرار (%)", "قوة الارتباط (%)", 
                                        "أهمية التوصية", "فرصة الأعمال", "النشاط الأكثر ارتباطًا"]), ""

        # إضافة عمود "النشاط الأكثر ارتباطًا"
        recommendations_df["النشاط الأكثر ارتباطًا"] = "لا"
        max_confidence_per_activity = recommendations_df.groupby("كود النشاط")["قوة الارتباط (%)"].max()
        for idx, row in recommendations_df.iterrows():
            if row["قوة الارتباط (%)"] == max_confidence_per_activity[row["كود النشاط"]]:
                recommendations_df["النشاط الأكثر ارتباطًا"] = "نعم"
        
        recommendations_df = recommendations_df.sort_values(by="قوة الارتباط (%)", ascending=False)
        
        # إنشاء ملخص نصي
        top_recommendation = recommendations_df.iloc[0]
        bottom_recommendation = recommendations_df.iloc[-1]
        summary = (
            f"**أعلى توصية**: إذا كنت تمارس '{top_recommendation['اسم النشاط']}'، يوصى بشدة بـ "
            f"'{top_recommendation['النشاط الموصى به']}' (قوة الارتباط: {top_recommendation['قوة الارتباط (%)']}%).\n"
            f"**أقل توصية**: إذا كنت تمارس '{bottom_recommendation['اسم النشاط']}'، فإن التوصية بـ "
            f"'{bottom_recommendation['النشاط الموصى به']}' لها قوة ارتباط أقل "
            f"({bottom_recommendation['قوة الارتباط (%)']}%)."
        )
        
        return recommendations_df, summary