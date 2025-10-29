import os
import json
import re
from openai import OpenAI
from dotenv import load_dotenv


def main():
    # === 1️⃣ 環境設定 ===
    load_dotenv()
    client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

    # === 2️⃣ Excel解析結果を読み込み ===
    with open("excel_analysis.json", "r", encoding="utf-8") as f:
        excel_data = json.load(f)

    # === 3️⃣ スキーマ定義 ===
    schema = {
        "table": "incident_report",
        "columns": {
            "title": "string",
            "event_date": "date",
            "location": "string",
            "reporter": "string",
            "progress": [
                {
                    "date": "date",
                    "timeline": [
                        {
                            "time": "time",
                            "description": "string",
                            "related_images": "list"
                        }
                    ]
                }
            ],
            "summary": {
                "result": "string",
                "cause": "string",
                "countermeasure": "string"
            }
        }
    }

    # === 4️⃣ AIへの指示 ===
    prompt = f"""
    あなたはExcel報告書を構造化するAIアシスタントです。
    次のExcel抽出データ（セルごとの内容と画像情報）をもとに、
    指定スキーマに従って構造化されたJSONを生成してください。

    # 🧭 基本方針
    - Excelの構造は一律ではなく、列や行の配置は自由です。
    - あなたの判断で、セルの内容や位置関係（上下・左右）をもとに、
      日付・時間・イベント・画像の関係を自然に推定してください。
    - 項目名の右側または下側に値があるケースが多いですが、固定ではありません。
    - スキーマに厳密に従い、欠損や重複を避けて階層化します。

    # 🧭 ルール（重要）
    1. 「発生日」「発生日時」などのキーワードを含むセルがあれば、
       その値を "event_date" に設定（1件のみ）。
    2. 日付のように見えるセル（例: "2025.8.26", "8月26日", "45895" など）は progress[].date に設定。
       - Excelシリアル値（例: 45894）は1900年基準で日付に変換。
    3. 各 progress 内の "timeline" には、その日付に属する出来事を格納する。
       - 時間（例: "14:17", "9:30", "21時35分"など）を timeline[].time に設定。
       - その時間セルの近く（右側または下側）のテキストを timeline[].description に入れる。
       - **セルが異なる場合（行が違う、列が違う）は、たとえ時間や内容が似ていても別のイベントとして扱う。**
       - **縦方向に連続したセルであっても、セルが分かれている限り、それぞれ独立した timeline として登録する。**
       - **同じ時間の記述が複数行に分かれていても、それぞれ別 timeline として扱う。**
       - セル内に「/」「／」「・」「,」「、」「改行（\\n）」などで複数の出来事が区切られている場合は、
         それぞれを別々の timeline として分割する。
       - 絶対に、複数セルの内容を勝手に統合して1つのイベントにまとめてはいけない。
    4. 画像の関連付け：
       - すべての画像(images[].anchor_cell)を対象に処理する。
       - 各画像について、上方向にたどって最初に見つかる「時間またはイベントセル」に関連付ける。
       - 列が異なる場合でも、同じ行または近い行内で最も近いイベントに紐づける。
       - **同じ行や近接した行に複数の画像がある場合は、それらを同じイベントの related_images にまとめてもよい。**
       - 1つのイベントに複数画像を紐づけることを許可する。
       - すべての画像は必ずいずれかのイベントに属する。
    5. "結果"・"原因"・"対策"・"対応"などの語を含むセルは summary に分類し、
       result / cause / countermeasure に対応させる。
    6. 出力は指定スキーマに厳密に従うこと。
       出力はJSONのみで、説明文・コメントは禁止。

    # スキーマ
    {json.dumps(schema, ensure_ascii=False, indent=2)}

    # セルデータ
    {json.dumps(excel_data["text_cells"], ensure_ascii=False, indent=2)}

    # 画像データ
    {json.dumps(excel_data["images"], ensure_ascii=False, indent=2)}
    """


    # === 5️⃣ OpenAI呼び出し ===
    response = client.chat.completions.create(
        model="gpt-4.1",
        temperature=0.2,
        messages=[
            {"role": "system", "content": "あなたは階層的な時系列データ整理に優れたAIアナライザーです。"},
            {"role": "user", "content": prompt}
        ]
    )

    raw_output = response.choices[0].message.content.strip()


    # Markdownの ```json ``` を除去
    clean_output = re.sub(r"```json|```", "", raw_output).strip()

    # === 7️⃣ JSONとして保存 ===
    try:
        result = json.loads(clean_output)
        with open("ai_mapped_result.json", "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print("\n✅ 結果を ai_mapped_result.json に保存しました。")
    except json.JSONDecodeError as e:
        print("\n⚠️ JSONパースに失敗しました。エラー内容:", e)


# === 単体実行対応 ===
if __name__ == "__main__":
    main()
