import pandas as pd
import numpy as np
import os
import shutil

##########################################
# 1) 원본 파일 경로와 새 파일 경로 생성
##########################################
file_path = r"C:\Users\slee\OneDrive - SBP\Tax Returns\xSungkeun\Tax Return\2024\TB\Q4 Alex US Trial Balances.xlsx"
base, ext = os.path.splitext(file_path)
new_file_path = base + "_Final" + ext  # 예: Q4 Alex US Trial BalancesFinal.xlsx

# 2) 원본 파일을 새 파일로 복사 (기존 시트들을 그대로 유지)
shutil.copy2(file_path, new_file_path)

##########################################
# 3) 처리할 시트 및 결과 테이블 컬럼명 매핑
##########################################
sheet_map = {
    "West": "West",
    "NE": "NE",
    "MW": "MW",
    "NSS": "NSS",
    "Direct LLC": "Direct LLC"
}

##########################################
# 4) 각 시트를 읽어 "Main" 기준으로 병합(outer join)
##########################################
df_merged = None
for sheet_name, col_name in sheet_map.items():
    # 각 시트에서 A열(Main)과 B열(Closing balance)만 읽어옵니다.
    df_temp = pd.read_excel(
        new_file_path,
        sheet_name=sheet_name,
        usecols=[0, 1],   # A열과 B열만 읽기
        header=0,         # 첫 번째 행을 헤더로 사용
        engine="openpyxl"
    )
    # 컬럼명을 "Main"과 해당 시트 이름으로 변경
    df_temp.columns = ["Main", col_name]
    
    # 첫 시트는 그대로 할당, 이후 시트는 "Main" 기준 outer merge
    if df_merged is None:
        df_merged = df_temp
    else:
        df_merged = pd.merge(df_merged, df_temp, on="Main", how="outer")

##########################################
# 5) 각 시트의 금액 데이터 결측치(NaN)를 0으로 대체
##########################################
for col in sheet_map.values():
    df_merged[col] = df_merged[col].fillna(0)

################################################################
# 6) "Closing balance" 열 추가: 5개 시트의 금액 합산 (숫자 형식 유지)
################################################################
df_merged["Closing balance"] = df_merged[list(sheet_map.values())].sum(axis=1)

##########################################
# 7) 열 순서 정리 (Main, West, NE, MW, NSS, Direct LLC, Closing balance)
##########################################
final_cols = ["Main"] + list(sheet_map.values()) + ["Closing balance"]
df_merged = df_merged[final_cols]

##########################################
# 8) 기존 원본 시트들은 유지한 채, 새 Consolidated 시트 추가하여 새 파일에 저장
##########################################
with pd.ExcelWriter(new_file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df_merged.to_excel(writer, sheet_name="Consolidated", index=False)

print("완료! 원본 탭은 그대로 두고, 'Consolidated' 시트를 추가한 새 파일이 생성되었습니다:")
print(new_file_path)
