
def analyze_financials(metrics_by_year):
    """
    Đưa ra nhận xét dựa trên các chỉ số tài chính theo năm.
    Input: metrics_by_year = {
        "2020": {"ROE": 12.5, "ROA": 6.2, ...},
        "2021": {"ROE": 13.8, ...}, ...
    }
    Output: dictionary chứa nhận xét phân tích
    """
    insights = {
        "profitability": "",
        "liquidity": "",
        "leverage": "",
        "efficiency": ""
    }

    roe_vals = [year_data.get("ROE", 0) for year_data in metrics_by_year.values()]
    roa_vals = [year_data.get("ROA", 0) for year_data in metrics_by_year.values()]

    if roe_vals[-1] > roe_vals[0]:
        insights["profitability"] = "Khả năng sinh lời có xu hướng cải thiện (ROE tăng)."
    else:
        insights["profitability"] = "Khả năng sinh lời giảm nhẹ theo thời gian."

    curr_ratios = [year_data.get("Current_Ratio", 0) for year_data in metrics_by_year.values()]
    if all(cr > 1.2 for cr in curr_ratios[-3:]):
        insights["liquidity"] = "Công ty duy trì thanh khoản tốt (Current Ratio > 1.2)."
    else:
        insights["liquidity"] = "Thanh khoản chưa ổn định, cần theo dõi thêm."

    debt_eq = [year_data.get("Debt_to_Equity", 0) for year_data in metrics_by_year.values()]
    if debt_eq[-1] > 100:
        insights["leverage"] = "Tỷ lệ nợ cao, có thể tiềm ẩn rủi ro tài chính."
    else:
        insights["leverage"] = "Cơ cấu vốn an toàn, đòn bẩy tài chính hợp lý."

    gross_margins = [year_data.get("Gross_Profit_Margin", 0) for year_data in metrics_by_year.values()]
    if gross_margins[-1] > gross_margins[0]:
        insights["efficiency"] = "Biên lợi nhuận gộp tăng, cho thấy hiệu quả cải thiện."
    else:
        insights["efficiency"] = "Hiệu quả hoạt động chưa cải thiện rõ rệt."

    return insights
