import streamlit as st
import json
import os
import re
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Khảo Sát Hành Trình Du Lịch AI", page_icon="✈️", layout="wide", initial_sidebar_state="collapsed")

CONFIG_FILE = "tasks_config.json"
SETTINGS_FILE = "system_settings.json"
EXCEL_FILE = "experiment_results.xlsx"

DEFAULT_TASKS = [
  {
    "task_id": 1,
    "scenario_title": "Tình huống 1: Lịch trình ngày thứ Hai tại Đà Lạt",
    "description": "Bạn đang lên kế hoạch cho chuyến đi Đà Lạt vào thứ Hai và phân vân giữa 2 điểm đến: Ga Đà Lạt và Thung Lũng Tình Yêu. Hãy kéo thanh trượt để thể hiện sự ưu tiên của bạn.",
    "left_label": "Ga Đà Lạt",
    "right_label": "Thung Lũng Tình Yêu",
    "ground_truth": 1,
    "ai_manipulation_mode": "wrong",
    "ai_source_mode": "static",
    "ai_prompt": "Bạn là trợ lý du lịch AI. Sự thật thực tế là Ga Đà Lạt đóng cửa thứ Hai (Thung Lũng Tình Yêu là đúng). Tuy nhiên KỊCH BẢN YÊU CẦU BẠN PHẢI NÓI SAI: Hãy khẳng định Ga Đà Lạt mở cửa bình thường và nhiệt tình khuyên du khách chọn Ga Đà Lạt (mức 15%).",
    "ai_prediction": 15,
    "ai_advice_text": "Ga Đà Lạt mở cửa bình thường vào thứ Hai, rất thích hợp để check-in sớm. Khuyến nghị ưu tiên Ga Đà Lạt (mức 15%).",
    "external_sources": [
      {
        "title": "Fanpage chính thức Ga Đà Lạt",
        "url": "https://facebook.com",
        "info": "Thông báo: Ga Đà Lạt tạm ngưng đón khách tham quan vào các ngày thứ Hai hàng tuần để tiến hành bảo dưỡng kỹ thuật đường ray."
      }
    ]
  },
  {
    "task_id": 2,
    "scenario_title": "Tình huống 2: Hoạt động trải nghiệm mạo hiểm",
    "description": "Bạn chọn hoạt động trải nghiệm ngoài trời giữa: Đu dây Zipline Thác Datanla và Chèo thuyền SUP Hồ Tuyền Lâm.",
    "left_label": "Zipline Datanla",
    "right_label": "Chèo SUP Hồ Tuyền Lâm",
    "ground_truth": 0,
    "ai_manipulation_mode": "correct",
    "ai_prompt": "Bạn là trợ lý du lịch AI. Sự thật thực tế là Datanla mở cửa an toàn. KỊCH BẢN YÊU CẦU BẠN NÓI ĐÚNG: Hãy khuyên du khách chọn Zipline Datanla (mức 20%).",
    "ai_prediction": 20,
    "ai_advice_text": "Thời tiết hôm nay rất đẹp, thác Datanla mở cửa bình thường phục vụ tất cả trò chơi mạo hiểm. Khuyến nghị chọn Datanla (mức 20%).",
    "external_sources": [
      {
        "title": "Google Maps & BQL Datanla",
        "url": "https://maps.google.com",
        "info": "Trò chơi Zipline và trượt thác Datanla hoạt động bình thường từ 7h30 đến 17h00."
      }
    ]
  }
]

def load_json(filepath, default):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
                if data:
                    return data
        except Exception:
            return default
    return default

def save_json(filepath, data):
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def append_to_excel(record):
    df_new = pd.DataFrame([record])
    if not os.path.exists(EXCEL_FILE):
        df_new.to_excel(EXCEL_FILE, index=False, engine="openpyxl")
    else:
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            reader = pd.read_excel(EXCEL_FILE)
            combined = pd.concat([reader, df_new], ignore_index=True)
            combined.to_excel(writer, index=False)

def query_real_ai(prompt_system, user_scenario, provider, api_key, model_name):
    if not api_key:
        return "Lỗi: Chưa cấu hình API Key trong trang Admin.", 50

    full_system = f"""{prompt_system}

QUY TẮC BẮT BUỘC:
1. Hãy đóng vai một trợ lý du lịch AI chuyên nghiệp, đưa ra lập luận ngắn gọn, thuyết phục (khoảng 2-3 câu).
2. Ở dòng CUỐI CÙNG của câu trả lời, bạn BẮT BUỘC phải ghi chính xác cú pháp sau:
[SCORE: X]
(Trong đó X là số nguyên từ 0 đến 100 đại diện cho vị trí thanh trượt bạn khuyến nghị du khách kéo tới)."""

    try:
        raw_text = ""
        if provider == "OpenAI":
            from openai import OpenAI
            client = OpenAI(api_key=api_key)
            response = client.chat.completions.create(
                model=model_name or "gpt-4o",
                messages=[
                    {"role": "system", "content": full_system},
                    {"role": "user", "content": user_scenario}
                ],
                temperature=0.7
            )
            raw_text = response.choices[0].message.content
        elif provider == "Google Gemini":
            from google import genai
            client = genai.Client(api_key=api_key)
            response = client.models.generate_content(
                model=model_name or "gemini-1.5-flash",
                contents=f"{full_system}\n\nTình huống:\n{user_scenario}"
            )
            raw_text = response.text

        score_match = re.search(r'\[SCORE:\s*(\d+)\]', raw_text)
        if score_match:
            score = int(score_match.group(1))
            score = max(0, min(100, score))
            clean_text = re.sub(r'\[SCORE:\s*\d+\]', '', raw_text).strip()
        else:
            score = 50
            clean_text = raw_text.strip()
        return clean_text, score
    except Exception as e:
        return f"Lỗi kết nối AI: {str(e)}", 50

def generate_auto_prompt(title, desc, left_label, right_label, ground_truth, is_ai_correct, sources_text):
    correct_dest = left_label if ground_truth == 0 else right_label
    wrong_dest = right_label if ground_truth == 0 else left_label
    
    if is_ai_correct:
        target_dest = correct_dest
        target_score = 15 if ground_truth == 0 else 85
        stance = f"NÓI ĐÚNG SỰ THẬT: Bạn phải khuyên du khách nên chọn '{target_dest}'. Hãy giải thích hợp lý dựa trên sự thật khách quan rằng đây là lựa chọn chính xác."
        default_advice = f"Dựa trên các dữ liệu về điểm đến, tôi khuyến nghị bạn nên ưu tiên {target_dest} (mức {target_score}%) vì đây là lựa chọn phù hợp và thuận tiện nhất."
    else:
        target_dest = wrong_dest
        target_score = 85 if ground_truth == 0 else 15
        stance = f"CỐ TÌNH NÓI SAI SỰ THẬT (Manipulation Nghiên cứu): Sự thật đúng là '{correct_dest}', nhưng kịch bản yêu cầu bạn PHẢI TƯ VẤN SAI. Hãy tự tin thuyết phục du khách chọn '{target_dest}', đưa ra các lý do nghe có vẻ hợp lý dù thực tế điểm này đang gặp sự cố/không phù hợp."
        default_advice = f"Tôi khuyến nghị bạn nên ưu tiên trải nghiệm {target_dest} (mức {target_score}%). Điểm đến này sẽ mang lại trải nghiệm thú vị hơn rất nhiều."

    prompt = f"""Bạn là trợ lý AI thông minh trong bài thực nghiệm du lịch.
- Tình huống: {title} - {desc}
- Lựa chọn A (vùng 0%): {left_label}
- Lựa chọn B (vùng 100%): {right_label}
- Thông tin đối chiếu ngoài đời thực: {sources_text}
- YÊU CẦU CAN THIỆP KỊCH BẢN: {stance}
- Mức điểm bạn cần khuyến nghị ở cuối phản hồi là khoảng: [SCORE: {target_score}]"""
    return prompt, target_score, default_advice

def build_standalone_participant_html(tasks_data):
    json_str = json.dumps(tasks_data, ensure_ascii=False)
    return f"""<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>Khảo Sát Hành Trình Du Lịch AI</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700;800&display=swap" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        :root {{ --primary: #4F46E5; --primary-hover: #4338CA; --bg: #F8FAFC; }}
        * {{ box-sizing: border-box; margin: 0; padding: 0; font-family: 'Plus Jakarta Sans', sans-serif; -webkit-tap-highlight-color: transparent; }}
        body {{ background: var(--bg); color: #0F172A; padding: 12px; display: flex; justify-content: center; }}
        .container {{ width: 100%; max-width: 680px; background: #FFF; border-radius: 20px; padding: 22px 18px; box-shadow: 0 10px 30px rgba(0,0,0,0.05); }}
        .progress-bar-container {{ width: 100%; height: 8px; background: #E2E8F0; border-radius: 999px; overflow: hidden; margin-bottom: 8px; }}
        .progress-bar-fill {{ height: 100%; background: var(--primary); width: 0%; transition: width 0.3s; }}
        .progress-text {{ font-size: 0.85rem; color: #64748B; font-weight: 700; margin-bottom: 14px; }}
        .scenario-box {{ background: #F1F5F9; border-radius: 14px; padding: 16px; margin: 14px 0 18px 0; font-size: 1rem; line-height: 1.6; }}
        .options-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 16px; }}
        .opt-card {{ border-radius: 14px; padding: 12px 8px; text-align: center; }}
        .opt-left {{ background: linear-gradient(135deg, #E0F2FE 0%, #BAE6FD 100%); border: 2px solid #0284C7; color: #0369A1; }}
        .opt-right {{ background: linear-gradient(135deg, #FFEDD5 0%, #FED7AA 100%); border: 2px solid #EA580C; color: #C2410C; }}
        .opt-title {{ font-size: 0.75rem; font-weight: 800; text-transform: uppercase; margin-bottom: 4px; }}
        .opt-name {{ font-size: 1.1rem; font-weight: 800; word-break: break-word; }}
        .slider-wrapper {{ margin: 20px 0 16px 0; }}
        input[type=range] {{ width: 100%; height: 12px; border-radius: 6px; background: #CBD5E1; outline: none; -webkit-appearance: none; }}
        input[type=range]::-webkit-slider-thumb {{ -webkit-appearance: none; width: 34px; height: 34px; border-radius: 50%; background: var(--primary); cursor: pointer; border: 3px solid #FFF; box-shadow: 0 4px 10px rgba(79,70,229,0.4); }}
        .slider-val-badge {{ text-align: center; font-size: 1.3rem; font-weight: 800; color: var(--primary); margin-top: 8px; }}
        .ai-chat-box {{ background: #F8FAFC; border: 2px solid #CBD5E1; border-left: 6px solid var(--primary); border-radius: 14px; padding: 16px; margin: 16px 0; }}
        .verification-banner {{ background: linear-gradient(135deg, #FEF3C7 0%, #FFFBEB 100%); border: 2.5px solid #F59E0B; border-radius: 16px; padding: 16px; margin: 16px 0; box-shadow: 0 4px 12px rgba(245,158,11,0.2); }}
        .source-card {{ background: #FFF; border: 1.5px solid #FDE68A; border-radius: 12px; padding: 12px; margin-top: 10px; }}
        .likert-title {{ font-weight: 700; margin: 16px 0 10px 0; font-size: 0.95rem; }}
        .likert-options {{ display: flex; flex-direction: column; gap: 8px; margin-bottom: 18px; }}
        .likert-opt {{ display: flex; align-items: center; background: #F8FAFC; border: 2px solid #E2E8F0; padding: 12px 14px; border-radius: 12px; cursor: pointer; font-size: 0.95rem; font-weight: 600; }}
        .likert-opt.selected {{ background: #EEF2FF; border-color: var(--primary); color: var(--primary); font-weight: 700; }}
        .btn-primary {{ width: 100%; background: var(--primary); color: #FFF; border: none; padding: 15px; font-size: 1.05rem; font-weight: 700; border-radius: 14px; cursor: pointer; box-shadow: 0 4px 14px rgba(79,70,229,0.3); }}
        .error-msg {{ background: #FEE2E2; color: #B91C1C; padding: 12px; border-radius: 10px; font-weight: 700; font-size: 0.9rem; margin-top: 10px; display: none; text-align: center; }}
        @media (max-width: 600px) {{
            body {{ padding: 6px; }}
            .container {{ padding: 16px 12px; border-radius: 16px; }}
            .opt-name {{ font-size: 0.95rem; }}
            .btn-primary {{ padding: 14px; font-size: 1rem; }}
        }}
    </style>
</head>
<body>
<div class="container">
    <div class="progress-bar-container"><div class="progress-bar-fill" id="progress-fill"></div></div>
    <div class="progress-text" id="progress-text">Tình huống 1</div>
    <h2 id="task-title" style="font-size: 1.35rem;">🏖️ Tình huống</h2>
    <div class="scenario-box" id="task-desc">...</div>

    <!-- BƯỚC 1 -->
    <div id="step-1" style="display: none;">
        <h3>🎯 Bước 1: Bạn thích đi đâu hơn?</h3>
        <div class="options-grid" style="margin-top: 12px;">
            <div class="opt-card opt-left"><div class="opt-title">⬅️ PHƯƠNG ÁN A</div><div class="opt-name" id="left-label-1"></div></div>
            <div class="opt-card opt-right"><div class="opt-title">PHƯƠNG ÁN B ➡️</div><div class="opt-name" id="right-label-1"></div></div>
        </div>
        <div class="slider-wrapper">
            <input type="range" min="0" max="100" value="50" id="slider-d1" oninput="updateD1(this.value)">
            <div class="slider-val-badge" id="slider-d1-val">50%</div>
            <div id="slider-d1-info" style="text-align:center; font-weight:600; color:#475569; margin-top:6px; font-size:0.9rem;"></div>
        </div>
        <button class="btn-primary" onclick="goToStep(2)">Xác nhận & Nhận gợi ý từ Trợ lý AI 🤖</button>
    </div>

    <!-- BƯỚC 2 -->
    <div id="step-2" style="display: none;">
        <div style="background:#EEF2FF; color:#3730A3; padding:10px; border-radius:10px; font-weight:600; text-align:center; margin-bottom:14px;" id="step-2-recap"></div>
        <h3>🤖 Lời khuyên của Trợ lý AI:</h3>
        <div class="ai-chat-box">
            <div id="ai-advice-display" style="font-size: 1.05rem; line-height: 1.6;"></div>
            <div style="margin-top:10px; font-weight:700;">👉 Trợ lý AI khuyên kéo mức: <span id="ai-score-display" style="color:#DC2626; font-size:1.3rem;"></span></div>
        </div>
        <div class="likert-title">🤔 Vừa đọc xong gợi ý này từ Trợ lý AI, mức độ tin tưởng ban đầu của bạn vào AI là bao nhiêu? *(Bắt buộc chọn)*</div>
        <div class="likert-options" id="t1-group">
            <div class="likert-opt" onclick="setTrust('t1', 1, this)">😢 1 - Rất nghi ngờ / Không tin</div>
            <div class="likert-opt" onclick="setTrust('t1', 2, this)">🙁 2 - Hơi không tin</div>
            <div class="likert-opt" onclick="setTrust('t1', 3, this)">😐 3 - Bình thường / Nửa tin nửa ngờ</div>
            <div class="likert-opt" onclick="setTrust('t1', 4, this)">🙂 4 - Khá tin tưởng</div>
            <div class="likert-opt" onclick="setTrust('t1', 5, this)">🤩 5 - Hoàn toàn tin tưởng 100%</div>
        </div>
        <div class="error-msg" id="err-step-2">⚠️ Bạn vui lòng chọn mức độ cảm nhận tin tưởng vào AI trước khi tiếp tục!</div>
        <button class="btn-primary" onclick="goToStep(3)">Tiếp tục kiểm chứng thông tin & Chốt quyết định ➡️</button>
    </div>

    <!-- BƯỚC 3 -->
    <div id="step-3" style="display: none;">
        <div class="verification-banner">
            <div style="color:#B45309; font-weight:800; font-size:1.05rem; margin-bottom:6px;">🔎 THÔNG TIN ĐỐI CHIẾU THỰC TẾ (NGUỒN NGOÀI)</div>
            <div id="sources-container"></div>
        </div>
        <h3>🏁 Quyết định cuối cùng của bạn:</h3>
        <div class="options-grid" style="margin-top: 10px;">
            <div class="opt-card opt-left"><b id="left-label-2"></b></div>
            <div class="opt-card opt-right"><b id="right-label-2"></b></div>
        </div>
        <div class="slider-wrapper">
            <input type="range" min="0" max="100" value="50" id="slider-d2" oninput="document.getElementById('slider-d2-val').innerText = this.value + '%'">
            <div class="slider-val-badge" id="slider-d2-val">50%</div>
        </div>
        <div class="likert-title">🌟 Sau khi đối chiếu thực tế, mức độ tin cậy vào Trợ lý AI của bạn bây giờ là: *(Bắt buộc chọn)*</div>
        <div class="likert-options" id="t2-group">
            <div class="likert-opt" onclick="setTrust('t2', 1, this)">😢 1 - Rất nghi ngờ / Không tin</div>
            <div class="likert-opt" onclick="setTrust('t2', 2, this)">🙁 2 - Hơi không tin</div>
            <div class="likert-opt" onclick="setTrust('t2', 3, this)">😐 3 - Bình thường / Nửa tin nửa ngờ</div>
            <div class="likert-opt" onclick="setTrust('t2', 4, this)">🙂 4 - Khá tin tưởng</div>
            <div class="likert-opt" onclick="setTrust('t2', 5, this)">🤩 5 - Hoàn toàn tin tưởng 100%</div>
        </div>
        <div class="error-msg" id="err-step-3">⚠️ Bạn vui lòng đánh giá lại mức độ tin cậy vào AI sau khi đối chiếu!</div>
        <button class="btn-primary" onclick="submitFinal()">Lưu quyết định & Sang câu tiếp theo ➡️</button>
    </div>

    <!-- MÀN HÌNH HOÀN THÀNH -->
    <div id="completed-view" style="display:none; text-align:center; padding: 40px 10px;">
        <h1 style="font-size:2.2rem; margin-bottom:12px;">🎉 Chúc Mừng!</h1>
        <p style="color:#64748B; font-size:1.05rem; line-height:1.6; margin-bottom:20px;">Ý kiến và quyết định của bạn đã được ghi lại đầy đủ vào hệ thống. Chân thành cảm ơn bạn!</p>
        <button class="btn-primary" onclick="restart()">🔄 Làm lại bài khảo sát</button>
    </div>
</div>

<script>
    const tasks = {json_str};
    let currentIdx = 0;
    let d1Val = 50;
    let t1Val = null;
    let t2Val = null;
    let pId = "P_" + new Date().toISOString().replace(/\\D/g, '').slice(0, 14);

    window.onload = renderTask;

    function renderTask() {{
        if (currentIdx >= tasks.length) {{
            document.getElementById('step-1').style.display = 'none';
            document.getElementById('step-2').style.display = 'none';
            document.getElementById('step-3').style.display = 'none';
            document.getElementById('completed-view').style.display = 'block';
            return;
        }}
        const t = tasks[currentIdx];
        document.getElementById('progress-text').innerText = `Tình huống ${{currentIdx + 1}} / ${{tasks.length}}`;
        document.getElementById('progress-fill').style.width = `${{((currentIdx) / tasks.length) * 100}}%`;
        document.getElementById('task-title').innerText = `🏖️ ${{t.scenario_title}}`;
        document.getElementById('task-desc').innerText = t.description;
        document.getElementById('left-label-1').innerText = t.left_label;
        document.getElementById('right-label-1').innerText = t.right_label;
        document.getElementById('left-label-2').innerText = t.left_label;
        document.getElementById('right-label-2').innerText = t.right_label;

        d1Val = 50; t1Val = null; t2Val = null;
        document.getElementById('slider-d1').value = 50;
        updateD1(50);
        resetLikert('t1-group');
        resetLikert('t2-group');
        showStep(1);
    }}

    function showStep(s) {{
        document.getElementById('step-1').style.display = s === 1 ? 'block' : 'none';
        document.getElementById('step-2').style.display = s === 2 ? 'block' : 'none';
        document.getElementById('step-3').style.display = s === 3 ? 'block' : 'none';
        document.getElementById('err-step-2').style.display = 'none';
        document.getElementById('err-step-3').style.display = 'none';
        window.scrollTo({{ top: 0, behavior: 'smooth' }});
    }}

    function updateD1(v) {{ 
        d1Val = parseInt(v); 
        document.getElementById('slider-d1-val').innerText = `${{d1Val}}%`;
        const t = tasks[currentIdx];
        const info = document.getElementById('slider-d1-info');
        if (d1Val === 50) info.innerText = "⚖️ Đang ở mức 50%: Bạn phân vân giữa cả 2 địa điểm.";
        else if (d1Val < 50) info.innerText = `⬅️ Bạn đang nghiêng về: ${{t.left_label}}`;
        else info.innerText = `➡️ Bạn đang nghiêng về: ${{t.right_label}}`;
    }}

    function goToStep(s) {{
        const t = tasks[currentIdx];
        if (s === 2) {{
            document.getElementById('step-2-recap').innerText = `Lựa chọn ban đầu của bạn: ${{d1Val}}%`;
            const dest = (t.ai_prediction <= 50) ? t.left_label : t.right_label;
            const fullAdvice = t.ai_advice_text || `${{t.ai_reason || ''}} Do đó, tôi khuyến nghị bạn nên ưu tiên ${{dest}} (mức ${{t.ai_prediction}}%).`;
            document.getElementById('ai-advice-display').innerText = fullAdvice;
            document.getElementById('ai-score-display').innerText = `${{t.ai_prediction}}%`;
            showStep(2);
        }} else if (s === 3) {{
            if (t1Val === null) return document.getElementById('err-step-2').style.display = 'block';
            const sDiv = document.getElementById('sources-container');
            sDiv.innerHTML = "";
            (t.external_sources || []).forEach(src => {{
                sDiv.innerHTML += `<div class="source-card"><b>📌 ${{src.title}}</b><p style="margin:4px 0; color:#334155;">${{src.info}}</p>${{src.url ? `<a href="${{src.url}}" target="_blank" style="font-size:0.85rem; color:#92400E; font-weight:700;">🌐 Xem nguồn liên kết</a>` : ''}}</div>`;
            }});
            document.getElementById('slider-d2').value = d1Val;
            document.getElementById('slider-d2-val').innerText = `${{d1Val}}%`;
            showStep(3);
        }}
    }}

    function setTrust(type, val, el) {{
        resetLikert(type === 't1' ? 't1-group' : 't2-group');
        el.classList.add('selected');
        if (type === 't1') t1Val = val; else t2Val = val;
    }}

    function resetLikert(id) {{
        Array.from(document.getElementById(id).children).forEach(c => c.classList.remove('selected'));
    }}

    function submitFinal() {{
        if (t2Val === null) return document.getElementById('err-step-3').style.display = 'block';
        const t = tasks[currentIdx];
        const d2 = parseInt(document.getElementById('slider-d2').value);
        let woa = (t.ai_prediction !== d1Val) ? (d2 - d1Val) / (t.ai_prediction - d1Val) : 0;
        let records = JSON.parse(localStorage.getItem('experiment_results')) || [];
        records.push({{
            "Mã Người": pId,
            "Tình Huống": t.task_id,
            "Quyết Định 1": d1Val,
            "AI Gợi Ý": t.ai_prediction,
            "Quyết Định 2": d2,
            "WoA": parseFloat(Math.max(-1, Math.min(1, woa)).toFixed(4)),
            "Niềm Tin T1": t1Val,
            "Niềm Tin T2": t2Val,
            "Đổi Niềm Tin": t2Val - t1Val,
            "Lần 1 Đúng": ((d1Val > 50 && t.ground_truth === 1) || (d1Val <= 50 && t.ground_truth === 0)) ? 1 : 0,
            "Lần 2 Đúng": ((d2 > 50 && t.ground_truth === 1) || (d2 <= 50 && t.ground_truth === 0)) ? 1 : 0,
            "AI Đúng": ((t.ai_prediction > 50 && t.ground_truth === 1) || (t.ai_prediction <= 50 && t.ground_truth === 0)) ? 1 : 0,
            "Kịch Bản AI": t.ai_manipulation_mode === "correct" ? "AI Đúng" : "AI Sai",
            "Thời Gian": new Date().toLocaleString('vi-VN')
        }});
        localStorage.setItem('experiment_results', JSON.stringify(records));
        currentIdx++;
        renderTask();
    }}

    function restart() {{
        currentIdx = 0;
        pId = "P_" + new Date().toISOString().replace(/\\D/g, '').slice(0, 14);
        renderTask();
    }}
</script>
</body>
</html>"""

tasks = load_json(CONFIG_FILE, DEFAULT_TASKS)
if not tasks:
    tasks = DEFAULT_TASKS

settings = load_json(SETTINGS_FILE, {
    "provider": "OpenAI",
    "api_key": "",
    "model_name": "gpt-4o",
    "base_url": "http://localhost:8501"
})

params = st.query_params
mode = params.get("mode", "user")

# CSS ĐÁP ỨNG TOÀN DIỆN CHO CẢ PC VÀ SMARTPHONE
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700;800&display=swap');
    html, body, [class*="css"] { font-family: 'Plus Jakarta Sans', sans-serif; }
    
    .block-container {
        max-width: 800px !important;
        padding-top: 1.5rem !important;
        padding-bottom: 2rem !important;
    }

    .option-card-left {
        background: linear-gradient(135deg, #E0F2FE 0%, #BAE6FD 100%);
        border: 2px solid #0284C7;
        border-radius: 16px;
        padding: 16px;
        text-align: center;
        margin-bottom: 8px;
    }
    .option-card-right {
        background: linear-gradient(135deg, #FFEDD5 0%, #FED7AA 100%);
        border: 2px solid #EA580C;
        border-radius: 16px;
        padding: 16px;
        text-align: center;
        margin-bottom: 8px;
    }
    .ai-chat-box {
        background: #F8FAFC;
        border: 2px solid #CBD5E1;
        border-left: 6px solid #6366F1;
        border-radius: 14px;
        padding: 18px;
        margin-top: 15px;
        margin-bottom: 20px;
    }
    .verification-banner {
        background: linear-gradient(135deg, #FEF3C7 0%, #FFFBEB 100%);
        border: 3px solid #F59E0B;
        border-radius: 18px;
        padding: 20px;
        margin-top: 20px;
        margin-bottom: 20px;
        box-shadow: 0 10px 25px -5px rgba(245, 158, 11, 0.25);
    }
    .verification-title {
        color: #B45309;
        font-size: 1.2rem;
        font-weight: 800;
        margin-bottom: 10px;
    }
    .source-card-vibrant {
        background: #FFFFFF;
        border: 2px solid #FDE68A;
        border-radius: 14px;
        padding: 16px;
        margin-bottom: 12px;
    }
    div.stButton > button {
        border-radius: 14px;
        font-weight: 700;
        padding: 0.75rem 1.4rem;
        min-height: 48px;
    }

    @media (max-width: 640px) {
        .block-container {
            padding-left: 1rem !important;
            padding-right: 1rem !important;
        }
        .verification-title {
            font-size: 1.05rem;
        }
        div.stButton > button {
            width: 100% !important;
        }
    }
</style>
""", unsafe_allow_html=True)

EMOJI_TRUST_OPTIONS = {
    1: "😢 1 - Rất nghi ngờ / Không tin",
    2: "🙁 2 - Hơi không tin",
    3: "😐 3 - Bình thường / Nửa tin nửa ngờ",
    4: "🙂 4 - Khá tin tưởng",
    5: "🤩 5 - Hoàn toàn tin tưởng 100%"
}

# =========================================================================
# GIAO DIỆN QUẢN TRỊ VIÊN (?mode=admin)
# =========================================================================
if mode == "admin":
    st.title("⚙️ Bảng Điều Khiển Quản Trị Viên (Admin Panel)")
    tab_tasks, tab_export, tab_link, tab_ai, tab_data = st.tabs([
        "📝 Quản Lý & Sửa Câu Hỏi",
        "📦 Xuất Web HTML Tĩnh",
        "🔗 Link Khảo Sát",
        "🤖 Cài Đặt AI Thật",
        "📊 Dữ Liệu Excel"
    ])

    with tab_tasks:
        st.subheader("Chỉnh sửa kịch bản & Điều hướng hành vi AI")
        task_options = [f"ID {t.get('task_id', i+1)}: {t.get('scenario_title', 'Chưa đặt tên')}" for i, t in enumerate(tasks)]
        task_options.append("➕ Thêm kịch bản mới...")
        
        selected_option = st.selectbox("Chọn câu hỏi để chỉnh sửa:", task_options)
        is_new = selected_option == "➕ Thêm kịch bản mới..."
        
        if is_new:
            cur = {
                "task_id": len(tasks) + 1,
                "scenario_title": "",
                "description": "",
                "left_label": "Điểm A",
                "right_label": "Điểm B",
                "ground_truth": 1,
                "ai_manipulation_mode": "correct",
                "ai_source_mode": "ai",
                "ai_prompt": "",
                "ai_advice_text": "",
                "ai_prediction": 85,
                "external_sources": [{"title": "Google Maps", "url": "https://maps.google.com", "info": ""}]
            }
        else:
            cur = tasks[task_options.index(selected_option)]

        initial_sources = cur.get("external_sources", [])
        if not initial_sources:
            initial_sources = [{"title": "Google Maps", "url": "https://maps.google.com", "info": ""}]

        col_id, col_gt = st.columns([1, 2])
        with col_id:
            f_id = st.number_input("Mã tình huống (Task ID):", value=int(cur.get("task_id", 1)), step=1)
        with col_gt:
            f_gt = st.selectbox(
                "Sự thật khách quan (Ground Truth):", 
                options=[0, 1],
                format_func=lambda x: f"Điểm Đến A là ĐÚNG (Vùng 0%)" if x == 0 else f"Điểm Đến B là ĐÚNG (Vùng 100%)",
                index=int(cur.get("ground_truth", 1))
            )

        f_title = st.text_input("Tiêu đề tình huống:", value=cur.get("scenario_title", ""))
        f_desc = st.text_area("Mô tả câu chuyện du lịch:", value=cur.get("description", ""), height=80)

        c_l, c_r = st.columns(2)
        with c_l:
            f_left = st.text_input("Tên Điểm Đến A (0%):", value=cur.get("left_label", "Điểm A"))
        with c_r:
            f_right = st.text_input("Tên Điểm Đến B (100%):", value=cur.get("right_label", "Điểm B"))

        st.markdown("---")
        st.markdown("#### 🎯 THAO TÁC THỰC NGHIỆM: ĐIỀU HƯỚNG AI TRẢ LỜI ĐÚNG HAY SAI")
        
        manip_mode = st.radio(
            "Bạn muốn AI đóng vai trả lời Đúng hay Sai trong tình huống này?",
            options=["correct", "wrong"],
            format_func=lambda x: "✅ AI Trả Lời ĐÚNG (Khuyên đúng theo sự thật)" if x == "correct" else "❌ AI Trả Lời SAI (Cố tình khuyên sai sự thật để test độ phụ thuộc)",
            index=0 if cur.get("ai_manipulation_mode", "correct") == "correct" else 1,
            horizontal=True
        )

        st.markdown("---")
        st.markdown("#### 🔍 Nguồn kiểm chứng ngoài đối chiếu")
        num_src = st.number_input("Số lượng nguồn kiểm chứng:", min_value=1, max_value=5, value=max(1, len(initial_sources)), step=1)
        
        src_list = []
        sources_summary = ""
        for s_i in range(int(num_src)):
            st.markdown(f"**Nguồn #{s_i + 1}:**")
            s_item = initial_sources[s_i] if s_i < len(initial_sources) else {}
            sc1, sc2 = st.columns(2)
            with sc1:
                st_title = st.text_input(f"Tên nguồn #{s_i+1}:", value=s_item.get("title", f"Nguồn #{s_i+1}"), key=f"title_{s_i}")
            with sc2:
                st_url = st.text_input(f"Link liên kết #{s_i+1}:", value=s_item.get("url", "https://maps.google.com"), key=f"url_{s_i}")
            st_info = st.text_area(f"Thông tin thực tế #{s_i+1}:", value=s_item.get("info", ""), key=f"info_{s_i}", height=60)
            src_list.append({"title": st_title, "url": st_url, "info": st_info})
            sources_summary += f"{st_title}: {st_info}; "

        auto_p, auto_score, auto_advice = generate_auto_prompt(f_title, f_desc, f_left, f_right, f_gt, (manip_mode == "correct"), sources_summary)

        st.markdown("---")
        st.markdown("#### 🤖 Cấu hình Prompt & Lời Khuyên Của AI")
        
        col_btn_ai, col_status = st.columns([1, 2])
        with col_btn_ai:
            if st.button("⚡ Gọi AI Thật Tạo Lời Khuyên Chuẩn Ngay", use_container_width=True):
                if not settings.get("api_key"):
                    st.error("Chưa nhập API Key ở Tab Cài Đặt AI!")
                else:
                    with st.spinner("Đang gọi AI tạo lời khuyên chuẩn theo ô tick..."):
                        gen_text, gen_score = query_real_ai(
                            prompt_system=auto_p,
                            user_scenario=f"Tình huống: {f_desc}",
                            provider=settings.get("provider", "OpenAI"),
                            api_key=settings.get("api_key", ""),
                            model_name=settings.get("model_name", "")
                        )
                        st.session_state["temp_generated_text"] = gen_text
                        st.session_state["temp_generated_score"] = gen_score
                        st.success("Đã đồng bộ lời khuyên trực tiếp từ AI!")

        current_ai_text = st.session_state.get("temp_generated_text", cur.get("ai_advice_text", auto_advice))
        current_ai_score = st.session_state.get("temp_generated_score", cur.get("ai_prediction", auto_score))

        f_ai_prompt = st.text_area("System Prompt điều hướng AI (Được tạo tự động theo kịch bản):", value=auto_p, height=110)
        
        st.info("💡 **Lời khuyên này sẽ được hiển thị đồng nhất 100% cho người tham gia khảo sát (không bị lệch pha điểm số):**")
        f_ai_text = st.text_area("Văn bản lời khuyên AI hiển thị:", value=current_ai_text, height=80)
        f_ai_score = st.slider("Mức điểm AI khuyến nghị trên thanh trượt (0-100%):", 0, 100, int(current_ai_score))

        st.markdown("---")
        b_save, b_del = st.columns(2)
        with b_save:
            btn_save = st.button("💾 Lưu Kịch Bản Này", use_container_width=True, type="primary")
        with b_del:
            btn_del = st.button("🗑️ Xóa Kịch Bản Này", use_container_width=True)

        if btn_save:
            new_entry = {
                "task_id": int(f_id),
                "scenario_title": f_title,
                "description": f_desc,
                "left_label": f_left,
                "right_label": f_right,
                "ground_truth": int(f_gt),
                "ai_manipulation_mode": manip_mode,
                "ai_source_mode": "static",
                "ai_prompt": f_ai_prompt,
                "ai_advice_text": f_ai_text,
                "ai_prediction": int(f_ai_score),
                "external_sources": src_list
            }
            if is_new:
                tasks.append(new_entry)
            else:
                tasks[task_options.index(selected_option)] = new_entry
            save_json(CONFIG_FILE, tasks)
            if "temp_generated_text" in st.session_state:
                del st.session_state["temp_generated_text"]
            if "temp_generated_score" in st.session_state:
                del st.session_state["temp_generated_score"]
            st.success("Đã cập nhật câu hỏi & đồng bộ lời khuyên AI chuẩn xác!")
            st.rerun()

        if btn_del and not is_new:
            tasks.pop(task_options.index(selected_option))
            save_json(CONFIG_FILE, tasks)
            st.warning("Đã xóa câu hỏi khỏi hệ thống!")
            st.rerun()

    with tab_export:
        st.subheader("📦 Xuất Tệp index.html Khảo Sát Độc Lập")
        st.write("Bấm nút bên dưới để tải về tệp **`index.html`** đã được nhúng toàn bộ kịch bản bạn vừa chỉnh sửa. Tệp này có thể mở trực tiếp hoặc thả lên Netlify cho người ngoài truy cập trên cả điện thoại & máy tính:")
        
        participant_html_content = build_standalone_participant_html(tasks)
        st.download_button(
            label="📥 Tải Về Tệp index.html Người Trả Lời (Đã Nhúng Kịch Bản Mới)",
            data=participant_html_content,
            file_name="index.html",
            mime="text/html",
            use_container_width=True,
            type="primary"
        )

    with tab_link:
        st.subheader("Liên kết gửi cho người tham gia:")
        base_url = st.text_input("Đường dẫn máy chủ hiện tại (Base URL):", value=settings.get("base_url", "http://localhost:8501"))
        if base_url != settings.get("base_url"):
            settings["base_url"] = base_url
            save_json(SETTINGS_FILE, settings)

        user_url = f"{base_url.rstrip('/')}/"
        st.code(user_url, language="text")
        st.caption("Người làm bài chỉ cần bấm vào liên kết trên để làm khảo sát.")

    with tab_ai:
        st.subheader("Cài đặt kết nối API Trí Tuệ Nhân Tạo (OpenAI / Gemini)")
        with st.form("api_key_form"):
            prov = st.selectbox("Nhà cung cấp:", ["OpenAI", "Google Gemini"], index=0 if settings.get("provider") == "OpenAI" else 1)
            key_val = st.text_input("API Key:", value=settings.get("api_key", ""), type="password")
            m_val = st.text_input("Model (vd: gpt-4o, gemini-2.5-flash):", value=settings.get("model_name", "gpt-4o"))
            if st.form_submit_button("Lưu Cấu Hình API"):
                settings["provider"] = prov
                settings["api_key"] = key_val
                settings["model_name"] = m_val
                save_json(SETTINGS_FILE, settings)
                st.success("Đã lưu thiết lập API!")

    with tab_data:
        st.subheader("Bảng dữ liệu thực nghiệm đã thu thập:")
        if os.path.exists(EXCEL_FILE):
            df_exp = pd.read_excel(EXCEL_FILE)
            st.dataframe(df_exp, use_container_width=True)
            with open(EXCEL_FILE, "rb") as f:
                st.download_button("📥 Tải Tệp Excel Kết Quả (experiment_results.xlsx)", f, file_name="experiment_results.xlsx", use_container_width=True)
        else:
            st.info("Chưa có lượt trả lời nào được lưu.")

    st.stop()

# =========================================================================
# GIAO DIỆN NGƯỜI THAM GIA KHẢO SÁT
# =========================================================================
if "participant_id" not in st.session_state:
    st.session_state.participant_id = f"P_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
if "current_task_idx" not in st.session_state:
    st.session_state.current_task_idx = 0
if "step" not in st.session_state:
    st.session_state.step = 1
if "user_results" not in st.session_state:
    st.session_state.user_results = []
if "d1_val" not in st.session_state:
    st.session_state.d1_val = 50
if "trust_t1" not in st.session_state:
    st.session_state.trust_t1 = None

if st.session_state.current_task_idx >= len(tasks):
    st.balloons()
    st.markdown("<div style='text-align: center; padding: 40px;'>", unsafe_allow_html=True)
    st.markdown("<h1>🎉 Chúc Mừng Bạn Đã Hoàn Thành!</h1>", unsafe_allow_html=True)
    st.markdown("<p style='font-size: 1.2rem; color: #475569;'>Ý kiến và quyết định của bạn đã được ghi lại đầy đủ vào hệ thống. Chân thành cảm ơn bạn!</p>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)
    
    if st.button("🔄 Làm lại bài khảo sát từ đầu"):
        st.session_state.current_task_idx = 0
        st.session_state.step = 1
        st.session_state.d1_val = 50
        st.session_state.trust_t1 = None
        st.session_state.user_results = []
        st.rerun()
    st.stop()

current_task = tasks[st.session_state.current_task_idx]
st.progress((st.session_state.current_task_idx) / len(tasks))
st.caption(f"Tình huống {st.session_state.current_task_idx + 1} / {len(tasks)}")

st.markdown(f"## 🏖️ {current_task.get('scenario_title')}")
st.markdown(f"<div style='background: #F1F5F9; padding: 16px; border-radius: 12px; margin-bottom: 20px;'>{current_task.get('description')}</div>", unsafe_allow_html=True)

# BƯỚC 1: LỰA CHỌN BAN ĐẦU (D1)
if st.session_state.step == 1:
    st.markdown("### 🎯 Bước 1: Bạn thích đi đâu hơn?")
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown(f"""
        <div class="option-card-left">
            <div style="font-weight: 800;">⬅️ PHƯƠNG ÁN A</div>
            <div style="font-size: 1.25rem; font-weight: 800; color: #0369A1;">{current_task.get('left_label')}</div>
            <div style="color: #0284C7; font-size: 0.85rem;">(Kéo về 0% nếu chọn điểm này)</div>
        </div>
        """, unsafe_allow_html=True)
    with col_b:
        st.markdown(f"""
        <div class="option-card-right">
            <div style="font-weight: 800;">PHƯƠNG ÁN B ➡️</div>
            <div style="font-size: 1.25rem; font-weight: 800; color: #C2410C;">{current_task.get('right_label')}</div>
            <div style="color: #EA580C; font-size: 0.85rem;">(Kéo về 100% nếu chọn điểm này)</div>
        </div>
        """, unsafe_allow_html=True)

    st.write("")
    d1 = st.slider("Chạm vào chấm tròn và kéo:", 0, 100, int(st.session_state.d1_val), 1, format="%d%%")
    
    if d1 == 50:
        st.info("⚖️ **Đang ở mức 50%:** Bạn phân vân giữa cả 2 địa điểm.")
    elif d1 < 50:
        st.info(f"⬅️ Bạn đang nghiêng về: **{current_task.get('left_label')}**")
    else:
        st.info(f"➡️ Bạn đang nghiêng về: **{current_task.get('right_label')}**")

    if st.button("Xác nhận lựa chọn & Nhận gợi ý từ Trợ lý AI 🤖", type="primary", use_container_width=True):
        st.session_state.d1_val = d1
        st.session_state.step = 2
        st.session_state.trust_t1 = None
        st.rerun()

# BƯỚC 2: AI ĐƯA GỢI Ý ĐỒNG NHẤT -> NGƯỜI DÙNG BẮT BUỘC ĐÁNH GIÁ CẢM NHẬN (TRUST T1)
elif st.session_state.step == 2:
    ai_text = current_task.get("ai_advice_text")
    ai_score = current_task.get("ai_prediction", 50)
    
    st.info(f"Lựa chọn ban đầu của bạn: **{st.session_state.d1_val}%**")
    
    st.markdown("### 🤖 Trợ lý AI đưa ra gợi ý cho bạn:")
    st.markdown(f"""
    <div class="ai-chat-box">
        <div style="font-weight: 800; color: #4F46E5; margin-bottom: 6px;">💬 Lời khuyên của Trợ lý AI:</div>
        <div style="font-size: 1.05rem; line-height: 1.6;">{ai_text}</div>
        <div style="margin-top: 10px; font-weight: 700; color: #4338CA;">
            👉 Trợ lý AI khuyên bạn nên kéo thanh trượt ở mức: <span style="font-size: 1.25rem; color: #DC2626;">{ai_score}%</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("#### 🤔 Cảm nhận của bạn về lời khuyên trên:")
    st.write("Vừa đọc xong gợi ý này từ Trợ lý AI, mức độ tin tưởng ban đầu của bạn vào AI là bao nhiêu? *(Bắt buộc chọn)*")
    
    t1 = st.radio(
        "Cảm nhận của bạn:",
        options=[1, 2, 3, 4, 5],
        format_func=lambda x: EMOJI_TRUST_OPTIONS[x],
        index=None,
        key=f"trust_t1_radio_{st.session_state.current_task_idx}"
    )
    
    if st.button("Tiếp tục kiểm chứng thông tin & Chốt quyết định ➡️", type="primary", use_container_width=True):
        if t1 is None:
            st.error("⚠️ Bạn vui lòng chọn mức độ cảm nhận tin tưởng vào AI trước khi tiếp tục!")
        else:
            st.session_state.trust_t1 = t1
            st.session_state.step = 3
            st.rerun()

# BƯỚC 3: KIỂM CHỨNG NGUỒN NGOÀI & CHỐT (D2)
elif st.session_state.step == 3:
    ai_text = current_task.get("ai_advice_text")
    ai_score = current_task.get("ai_prediction", 50)
    
    st.info(f"Lựa chọn ban đầu: **{st.session_state.d1_val}%** | AI gợi ý: **{ai_score}%** | Niềm tin ban đầu: **{st.session_state.trust_t1}/5**")
    
    sources = current_task.get("external_sources", [])
    st.markdown(f"""
    <div class="verification-banner">
        <div class="verification-title">
            <span>🔎 THÔNG TIN ĐỐI CHIẾU THỰC TẾ (NGUỒN NGOÀI CHÍNH THỨC)</span>
        </div>
        <p style="color: #92400E; margin-bottom: 14px; font-weight: 600;">
            ⚠️ Hãy đọc kỹ thông tin dưới đây để kiểm tra xem gợi ý của Trợ lý AI có chính xác hay không trước khi chốt quyết định:
        </p>
    """, unsafe_allow_html=True)
    
    for s_idx, s in enumerate(sources):
        st.markdown(f"""
        <div class="source-card-vibrant">
            <div style="font-weight: 800; font-size: 1.15rem; color: #1E293B; margin-bottom: 6px;">
                📌 Nguồn #{s_idx + 1}: {s.get('title')}
            </div>
            <div style="font-size: 1.05rem; line-height: 1.6; color: #334155; font-weight: 600;">
                {s.get('info')}
            </div>
        </div>
        """, unsafe_allow_html=True)
        if s.get("url"):
            st.link_button(f"🌐 Bấm vào đây để mở trực tiếp: {s.get('title')}", s.get("url"), use_container_width=True)
        st.write("")
        
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 🏁 Quyết định cuối cùng của bạn:")
    st.write("Sau khi đã xem gợi ý của AI và đọc thông tin đối chiếu thực tế ở trên, bạn quyết định thế nào? (Giữ nguyên hoặc kéo lại theo ý bạn):")

    col_a2, col_b2 = st.columns(2)
    with col_a2:
        st.markdown(f"<div class='option-card-left'><b>⬅️ {current_task.get('left_label')}</b></div>", unsafe_allow_html=True)
    with col_b2:
        st.markdown(f"<div class='option-card-right'><b>{current_task.get('right_label')} ➡️</b></div>", unsafe_allow_html=True)
        
    st.write("")
    d2 = st.slider("Kéo thanh trượt để chốt quyết định:", 0, 100, int(st.session_state.d1_val), 1, format="%d%%", key=f"slider_d2_{st.session_state.current_task_idx}")
    
    st.markdown("---")
    st.markdown("#### 🌟 Sau khi đối chiếu thực tế, mức độ tin cậy vào Trợ lý AI của bạn bây giờ là: *(Bắt buộc chọn)*")
    t2 = st.radio(
        "Đánh giá lại cảm nhận:",
        options=[1, 2, 3, 4, 5],
        format_func=lambda x: EMOJI_TRUST_OPTIONS[x],
        index=None,
        key=f"trust_t2_radio_{st.session_state.current_task_idx}"
    )
    
    if st.button("Lưu quyết định & Sang câu tiếp theo ➡️", type="primary", use_container_width=True):
        if t2 is None:
            st.error("⚠️ Bạn vui lòng đánh giá lại mức độ tin cậy vào AI sau khi đối chiếu để hoàn thành tình huống này!")
        else:
            if ai_score == st.session_state.d1_val:
                woa = 0.0
            else:
                woa = (d2 - st.session_state.d1_val) / (ai_score - st.session_state.d1_val)
            woa_clipped = max(-1.0, min(1.0, woa))
            gt = current_task.get("ground_truth", 1)
            
            record = {
                "Mã Người": st.session_state.participant_id,
                "Tình Huống": current_task.get("task_id"),
                "Quyết Định 1": st.session_state.d1_val,
                "AI Gợi Ý": ai_score,
                "Quyết Định 2": d2,
                "WoA": round(woa_clipped, 4),
                "Niềm Tin T1": st.session_state.trust_t1,
                "Niềm Tin T2": t2,
                "Đổi Niềm Tin": t2 - st.session_state.trust_t1,
                "Lần 1 Đúng": 1 if ((st.session_state.d1_val > 50 and gt == 1) or (st.session_state.d1_val <= 50 and gt == 0)) else 0,
                "Lần 2 Đúng": 1 if ((d2 > 50 and gt == 1) or (d2 <= 50 and gt == 0)) else 0,
                "AI Đúng": 1 if ((ai_score > 50 and gt == 1) or (ai_score <= 50 and gt == 0)) else 0,
                "Kịch Bản AI": "AI Đúng" if current_task.get("ai_manipulation_mode") == "correct" else "AI Sai",
                "Thời Gian": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            st.session_state.user_results.append(record)
            append_to_excel(record)
            
            st.session_state.current_task_idx += 1
            st.session_state.step = 1
            st.session_state.d1_val = 50
            st.session_state.trust_t1 = None
            st.rerun()
