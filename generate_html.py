import json
import os

def main():
    # === JSON読み込み ===
    with open("ai_mapped_result.json", "r", encoding="utf-8") as f:
        data = json.load(f)["columns"]

    # === HTMLテンプレート ===
    html = f"""
    <!DOCTYPE html>
    <html lang="ja">
    <head>
    <meta charset="utf-8">
    <title>{data['title']}</title>
    <style>
    body {{
      font-family: "Hiragino Sans", "Segoe UI", sans-serif;
      background: linear-gradient(180deg, #f0f4f8, #ffffff);
      color: #222;
      line-height: 1.7;
      margin: 40px auto;
      max-width: 900px;
      padding: 0 20px;
    }}
    .card {{
      background: #fff;
      border-radius: 16px;
      padding: 24px;
      margin-bottom: 32px;
      box-shadow: 0 4px 12px rgba(0,0,0,0.08);
      transition: transform 0.2s ease;
    }}
    .card:hover {{
      transform: translateY(-2px);
    }}
    h1 {{
      color: #0056b3;
      font-size: 1.8rem;
      margin-bottom: 0.5rem;
    }}
    h2 {{
      border-bottom: 3px solid #0078d7;
      padding-bottom: 6px;
      color: #005fa3;
      font-size: 1.3rem;
      margin-top: 1.5rem;
    }}
    .date {{
      font-weight: bold;
      color: #0078d7;
      margin-top: 16px;
      font-size: 1.1rem;
    }}
    .time {{
      font-weight: bold;
      color: #333;
      margin-top: 8px;
      font-size: 1rem;
    }}
    .event {{
      background: #f5f9ff;
      border-left: 5px solid #0078d7;
      border-radius: 8px;
      padding: 12px 16px;
      margin: 10px 0;
    }}
    .images {{
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      margin-top: 10px;
    }}
    .images img {{
      max-width: 200px;
      border-radius: 8px;
      box-shadow: 0 2px 6px rgba(0,0,0,0.15);
      transition: transform 0.2s ease;
    }}
    .images img:hover {{
      transform: scale(1.03);
    }}
    .summary-item {{
      margin: 6px 0;
    }}
    </style>
    </head>
    <body>

    <div class="card">
      <h1>📋 {data['title']}</h1>
      <p>📅 発生日: {data['event_date']}　📍 {data['location']}　👤 {data['reporter']}</p>
    </div>

    <div class="card">
      <h2>📆 Progress（進捗）</h2>
    """

    # === progress部分 ===
    for prog in data.get("progress", []):
        html += f"<div class='date'>🗓️ {prog.get('date', '（日付不明）')}</div>"

        for timeline in prog.get("timeline", []):
            time_val = timeline.get("time")
            description = timeline.get("description", "").replace("\n", "<br>")

            # ⏰ time があれば表示、なければスキップ
            if time_val and str(time_val).strip().lower() not in ["none", "null", ""]:
                html += f"<div class='time'>⏰ {time_val}</div>"

            # description が空ならスキップ（完全空のイベント除外）
            if not description.strip():
                continue

            html += f"<div class='event'>{description}"

            imgs = timeline.get("related_images", [])
            if imgs:
                html += "<div class='images'>"
                for img in imgs:
                    if isinstance(img, dict):
                        img_path = img.get("image_path") or img.get("image_name")
                    else:
                        img_path = str(img)
                    if not os.path.exists(img_path):
                        img_path = f"output_images/{os.path.basename(img_path)}"
                    html += f"<img src='{img_path}' alt='{os.path.basename(img_path)}'>"
                html += "</div>"

            html += "</div>"

    html += f"""
    </div>

    <div class="card">
      <h2>🧭 Summary（要約）</h2>
      <p class="summary-item"><b>結果:</b> {data['summary'].get('result', '')}</p>
      <p class="summary-item"><b>原因:</b> {data['summary'].get('cause', '')}</p>
      <p class="summary-item"><b>対策:</b> {data['summary'].get('countermeasure', '')}</p>
    </div>

    </body>
    </html>
    """

    # === 保存 ===
    with open("report.html", "w", encoding="utf-8") as f:
        f.write(html)

    print("✅ report.html を生成しました。")


if __name__ == "__main__":
    main()
