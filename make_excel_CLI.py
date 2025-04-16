import os
import pandas as pd

# 파일 경로 입력 받기
file1 = input("관리자페이지_건별매출조회 파일 경로를 입력하세요: ").strip('"')
file2 = input("KIS_정산보고 파일 경로를 입력하세요: ").strip('"')

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

# 저장 경로
output_dir = os.path.dirname(file1)
os.makedirs(output_dir, exist_ok=True)
output_path = os.path.join(output_dir, '결과_수수료합.xlsx')

# 저장
merged.to_excel(output_path, index=False)
print("✅ 저장 완료:", output_path)