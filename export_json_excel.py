import json
import pandas as pd

# อ่านข้อมูล JSON ที่มีค่า null
json_data = """
[
    {
        "accNam": null,
        "accNo": null,
        "accTyp": null,
        "aprDat": null,
        "aprReason": null,
        "bAddNo": "8/63",
        "bAmpCod": "103600",
        "bAmpNam": "ดอนเมือง",
        "bBldgNam": "-",
        "bBraNam": "สำนักงาน  เหมือนฝัน น้อยไขขำ (ทดสอบ)",
        "bEmail": null,
        "bFloorNo": "-",
        "bMooNo": "-",
        "bPosCod": "10210",
        "bProvCod": "100000",
        "bProvNam": "กรุงเทพมหานคร",
        "bRoomNo": "-",
        "bSoiNam": "-",
        "bTamCod": "103605",
        "bTamNam": "สนามบิน",
        "bTelNo": "0835985588",
        "bThnNam": "วิภาวดีรังสิต",
        "bTitCod": "00002130",
        "bVillage": "ชวนชื่น โมดัส-วิภาวดี",
        "bWebsite": null,
        "bYaek": "",
        "bnkBraCod": null,
        "bnkBraNam": "",
        "bnkCod": null,
        "bnkNam": "",
        "braNo": 0,
        "chgTyp": "1",
        "dln": null,
        "docAprDat": null,
        "docReqDat": null,
        "docReqNo": null,
        "engAccNam": null,
        "firNam": "เหมือนฝัน น้อยไขขำ (ทดสอบ)",
        "fullName": "นางสาว  เหมือนฝัน น้อยไขขำ (ทดสอบ)",
        "hAddNo": "8/63",
        "hAmpCod": "103600",
        "hAmpNam": "ดอนเมือง",
        "hBldgNam": "-",
        "hBraNam": "สำนักงาน  เหมือนฝัน น้อยไขขำ (ทดสอบ)",
        "hEmail": null,
        "hFloorNo": "-",
        "hMooNo": "-",
        "hPosCod": "10210",
        "hProvCod": "100000",
        "hProvNam": "กรุงเทพมหานคร",
        "hRoomNo": "-",
        "hSoiNam": "-",
        "hTamCod": "103605",
        "hTamNam": "สนามบิน",
        "hTelNo": "0835985588",
        "hThnNam": "วิภาวดีรังสิต",
        "hTitCod": "00002130",
        "hVillage": "ชวนชื่น โมดัส-วิภาวดี",
        "hWebsite": null,
        "hYaek": "",
        "hasVal": null,
        "ipReg": "10.1.233.182",
        "lasNam": "",
        "ltoFlg": "0",
        "midNam": "",
        "nid": "0105565176530",
        "offCod": "01009410",
        "pin": "0105565176530",
        "pp13Id": "0",
        "pp13PdfUrl": null,
        "regTyp": "3",
        "titleNam": "นางสาว",
        "toAccNam": "นางสาวเหมือนฝัน น้อยไขขำ",
        "toAccNo": "0984245234",
        "toAccTyp": "ออมทรัพย์",
        "toBnkBraCod": "056 ",
        "toBnkBraNam": "สี่แยกบางนา",
        "toBnkCod": "004",
        "toBnkNam": " ธนาคารกสิกรไทย จำกัด (มหาชน)",
        "toDocAprDat": null,
        "toDocDat": null,
        "toDocNo": null,
        "toEngAccNam": "Mueanfan Noikaikam",
        "traDat": null,
        "uid": null,
        "uploads": [],
        "userId": "0105565176530"
    }
]
"""

# แปลง JSON string ให้เป็น Python dictionary
data = json.loads(json_data)

# แปลงข้อมูล JSON เป็น pandas DataFrame
df = pd.json_normalize(data)

# บันทึก DataFrame ไปยังไฟล์ Excel
excel_filename = "output_data.xlsx"
df.to_excel(excel_filename, index=False, engine='openpyxl')

print(f"Data has been written to {excel_filename}")
