import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# GUI 창 숨기기
root = tk.Tk()
root.withdraw()

# 파일 선택
messagebox.showinfo("안내", "먼저 '건별매출조회' 엑셀 파일을 선택해주세요.")
file1 = filedialog.askopenfilename(title="건별매출조회 파일 선택", filetypes=[("Excel files", "*.xls;*.xlsx")])

messagebox.showinfo("안내", "다음으로 'KIS 정산보고' 엑셀 파일을 선택해주세요.")
file2 = filedialog.askopenfilename(title="KIS 정산보고 파일 선택", filetypes=[("Excel files", "*.xls;*.xlsx")])

# 파일 선택 안했을 경우 종료
if not file1 or not file2:
    messagebox.showerror("오류", "두 개의 엑셀 파일을 모두 선택해야 합니다.")
    exit()

try:
    # 엑셀 불러오기
    df1 = pd.read_excel(file1, engine='xlrd')
    df2 = pd.read_excel(file2, engine='xlrd')

    # 병합
    merged = df1.merge(df2[['승인번호', '결제수수료', 'VAT']], on='승인번호', how='left')

    # 매출금액 컬럼 찾기
    매출컬럼 = [col for col in merged.columns if '매출' in col and '금액' in col]
    if not 매출컬럼:
        raise KeyError("매출금액 컬럼을 찾을 수 없습니다.")
    매출금액_컬럼명 = 매출컬럼[0]

    # 컬럼 초기화
    merged['수수료합'] = 0
    merged['라운드로빈'] = 0

    # 계산 함수
    def 수수료계산(group):
        group = group.copy()

        수수료_계산값 = [round(row[매출금액_컬럼명] * 0.011 * 1.1) for _, row in group.iterrows()]
        group['라운드로빈'] = 수수료_계산값

        결제수수료 = group['결제수수료'].fillna(0).iloc[0]
        vat = group['VAT'].fillna(0).iloc[0]
        총수수료 = 결제수수료 + vat
        수수료합계 = sum(수수료_계산값)
        차이 = 총수수료 - 수수료합계

        group.iloc[0, group.columns.get_loc('수수료합')] = 총수수료

        if 차이 != 0 and len(group) > 0:
            idx = group['라운드로빈'].idxmin()
            group.loc[idx, '라운드로빈'] += 차이

        if len(group) > 1:
            group.iloc[1:, group.columns.get_loc('결제수수료')] = None
            group.iloc[1:, group.columns.get_loc('VAT')] = None
            group.iloc[1:, group.columns.get_loc('수수료합')] = None

        return group

    # 그룹별 계산 적용
    merged = merged.groupby('승인번호', group_keys=False).apply(수수료계산)

    # 승인번호 문자열 보정
    merged['승인번호'] = merged['승인번호'].astype(str).apply(lambda x: x.zfill(8))

    # 저장
    output_dir = os.path.dirname(file1)
    output_path = os.path.join(output_dir, '결과_수수료합.xlsx')
    merged.to_excel(output_path, index=False)

    messagebox.showinfo("완료", f"✅ 저장 완료:\n{output_path}")

except Exception as e:
    messagebox.showerror("에러 발생", str(e))
