import pandas as pd
import json

# อ่านไฟล์ CSV
csv_filename = 'json_to_excel_output.csv'  # เปลี่ยนเป็นชื่อไฟล์ CSV ของคุณ
df = pd.read_csv(csv_filename)

# แปลง DataFrame เป็น JSON
json_data = df.to_json(orient='records', date_format='iso', default_handler=str)

# แสดงผลลัพธ์ JSON
print(json.dumps(json.loads(json_data), indent=4, ensure_ascii=False))

# บันทึกไฟล์ JSON
with open('output_data.json', 'w', encoding='utf-8') as json_file:
    json.dump(json.loads(json_data), json_file, ensure_ascii=False, indent=4)

print(f"Data has been written to output_data.json")
