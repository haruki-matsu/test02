import excel_to_json
import json_to_schema_by_llm
import generate_html

print("=== Step 1: Excel → JSON ===")
excel_to_json.main()

print("\n=== Step 2: JSON → スキーマ変換（LLM） ===")
json_to_schema_by_llm.main()

print("\n=== Step 3: HTMLレポート生成 ===")
generate_html.main()

print("\n🎉すべて完了しました！ report.html を確認してください。")
