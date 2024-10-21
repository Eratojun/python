import pandas as pd

from openpyxl import load_workbook

# 엑셀 파일 읽어오기
mon_df = pd.read_excel('C:\dev\myproj01\금상동\결과폴더\월별결산부2023.xlsx')  # 메인 엑셀 파일


mon_df['수입'] = mon_df['수입'].fillna(0)
mon_df['지출'] = mon_df['지출'].fillna(0)

# '구분1', '수입', '지출' 열에 NaN 값이 있는 행 제거
#mon_df = mon_df.dropna(subset=['구분1', '수입', '지출'])
mon_df['금액2']=mon_df['금액2']=0


# 금액 계산을 위한 함수 정의
mon_df['지출']=mon_df['지출'].astype(int)
mon_df['수입']=mon_df['수입'].astype(int)









mon_df['코드1'] = None
mon_df['구분1']=None
mon_df['구분2']=None
mon_df.info()



#mon_df.info()

# 두 문자열 칼럼을 합쳐서 새로운 칼럼에 저장
#mon_df['적요명'] = mon_df['적       요'].astype(str) + mon_df['봉안실'].astype(str)

import pandas as pd

# 1. 엑셀 파일에서 회계 데이터를 불러오기
#mon_df = pd.read_excel('accounting_data.xlsx')









account_code_df = pd.read_excel('C:\dev\myproj01\금상동\결과폴더\금상동계정코드.xlsx')  # 계정코드 엑셀 파일

account_code_df.info()

# '코드1'이라는 새로운 열 추가

# 조건에 따른 값을 계산하는 함수 정의
def calculate_gubun(row):
    E = str(row['납부방법'])  # E3608
    D = str(row['계정과목']) # D3608
    G =str(row['회계계정'])  # G3608
    H = row['수입']  # H3608
    I = row['지출']  # I3608

    if "현금" in E and H > 0:
        return 2
    elif ("대체" in E or "이체" in E) and "해약" in D and I > 0:
        return 4
    elif (
        ("이체" in E and H > 0) or
        ("이체" in E and "일반관리선수" in G and I > 0) or
        ("이체" in E and "회전관리" in G and I > 0) or
        ("이체" in E and "사용수입" in G and I > 0) or
        ("대체" in E and "사용수입" in G and H > 0) or
        ("대체" in E and "관리선수" in G and H > 0) or
        ("대체" in E and "관리선수" in G and I > 0) or
        ("대체" in E and "사용수입" in G and H > 0) or
        ("신용" in E and H > 0)
    ):
        return 4
    else:
        return 0

# 각 행에 대해 함수 적용
mon_df['구분1'] = mon_df.apply(calculate_gubun, axis=1)

mon_df['구분1']=mon_df['구분1'].astype(float)



mon_df['구분1'] = mon_df['구분1'].fillna(0)













# 조건에 따른 값 설정
for index, row in mon_df.iterrows():
    # 첫번째 조건
    if "해약" in str(row['계정과목']) and row['납부방법'] == "이체" and (row['회계계정'] in ["관리선수", "사용선수"]) and row['지출'] > 0:
        mon_df.at[index, '코드1'] = 103
    # 두번째 조건
    elif "해약" in str(row['계정과목']) and row['회계계정'] == "사용수입" and row['지출'] > 0:
        mon_df.at[index, '코드1'] = 420
    # 세번째 조건
    elif "관리비" in str(row['계정과목']) and "관리수입" in str(row['회계계정']) and "일반관리선수" in str(mon_df.at[index+1, '회계계정']) and row['수입'] > 0:
        mon_df.at[index, '코드1'] = 276
    # 네번째 조건
    elif "관리비" in str(row['계정과목']) and "회전관리선수" in str(mon_df.at[index+1, '회계계정']) and row['수입'] > 0:
        mon_df.at[index, '코드1'] = 275
    # 다섯번째 조건
    elif "대체" in str(row['납부방법']) and "관리선수" in str(row['회계계정']) and row['지출'] > 0:
        mon_df.at[index, '코드1'] = 421
    # 여섯번째 조건
    elif "이체" in str(row['납부방법']) and row['회계계정'] == "관리수입":
        mon_df.at[index, '코드1'] = 274
    # 일곱번째 조건
    elif row['수입'] == 0:
        mon_df.at[index, '코드1'] = 0
    # VLOOKUP 유사 작업
    else:
        lookup_value = str(row['회계계정'])[:6]  # G열의 첫 6글자 추출
        code_row = account_code_df[account_code_df['계정코드'].fillna('').str.startswith(lookup_value)]
        
        
        
        #code_row = account_code_df[account_code_df['계정'].str.startswith(lookup_value)]
        if not code_row.empty:
            mon_df.at[index, '코드1'] = code_row.iloc[0]['코드1']
        else:
            mon_df.at[index, '코드1'] = 0  # 일치하는 값이 없으면 0 처리
    
     
    
    
    
    
    
# 결과를 엑셀 파일로 저장




# 2. VLOOKUP 유사 기능: '코드1' 열을 기준으로 lookup_df의 4번째 열과 병합
# lookup_df의 4번째 열의 이름을 임시로 설정
lookup_df_column_name = account_code_df.columns[5]  # 4번째 열의 이름 가져오기

# 병합 수행
merged_df = pd.merge(mon_df, account_code_df[['코드2','계정코드2']], how='left',left_on='코드1',right_on='코드2')






# 3. 병합된 결과를 엑셀 파일로 저장
#mon_df.to_excel('병합결과_엑셀파일.xlsx', index=False)


# 3. 불필요한 '코드2' 열 제거 (선택 사항)
merged_df.drop(columns=['코드2'], inplace=True)

#mon_df=mon_df['적       요'].fillna(method='ffill') # 최근일 월일없는행은 전일자로 처리




merged_df.iloc[:, 3] = merged_df.iloc[:, 3].fillna(method='ffill')


# 두 문자열 칼럼을 합쳐서 새로운 칼럼에 저장
merged_df['적요명'] = merged_df.iloc[:, 3].astype(str) + merged_df.iloc[:, 6].astype(str)

merged_df.info()
# 금액 계산을 위한 함수 정의
def calculate_amount(row):
    if row['구분1'] == 0:
        return (row['수입'] + row['지출']) * float(row['구분1'])
    elif (
        '해약' in str(row['계정과목']) and
        '이체' in row['납부방법'] and
        (
            '관리선수' in row['회계계정'] or
            '사용선수' in row['회계계정']
        ) and
        row['지출'] > 0
    ):
        return row['수입'] + row['지출'] * 1
    elif (
        '해약' in str(row['계정과목']) and
        row['지출'] > 0
    ):
        return row['수입'] + row['지출'] * -1
    else:
        return row['수입'] + row['지출']
merged_df['금액'] = mon_df.apply(calculate_amount, axis=1)
# 데이터프레임에 '금액' 열 추가
#mon_df['금액'] = mon_df.apply(calculate_amount, axis=1)

merged_df['지출'] =merged_df['지출'].fillna(0)

merged_df['수입'] =merged_df['수입'].fillna(0)

merged_df['금액'] = merged_df['금액'].fillna(0)

merged_df['금액']=merged_df['금액'].astype(int)






# 코드3 값을 계산하는 함수 정의
def calculate_code2(row):
    if row['구분1'] == 0:
        return 0
    elif (
        ('이체' in str(row['납부방법']) and '일반관리' in str(row['회계계정']) and row['지출'] > 0) or
        ('이체' in str(row['납부방법']) and '회전관리선수' in str(row['회계계정']) and row['지출'] > 0) or
        ('이체' in str(row['납부방법']) and '회전' in str(row['회계계정']) and row['지출'] > 0) or
        ('대체' in str(row['납부방법']) and '관리선수' in str(row['회계계정']) and row['지출'] > 0) or
        ('대체' in str(row['납부방법']) and '관리수입' in str(row['회계계정']) and row['수입'] > 0)
    ):
        if '이체' in str(row['납부방법']) and '일반관리' in str(row['회계계정']) and row['지출'] > 0:
            return 276
        elif '이체' in str(row['납부방법']) and '회전관리선수' in str(row['회계계정']) and row['지출'] > 0:
            return 274
        elif '이체' in str(row['납부방법']) and '회전' in str(row['회계계정']) and row['지출'] > 0:
            return 275
        else:
            return 274
    elif (
        ('이체' in str(row['납부방법']) and '사용수입' in str(row['회계계정']) and row['지출'] > 0) or
        ('대체' in str(row['납부방법']) and '사용수입' in str(row['회계계정']))
    ):
        return 273
    elif '이체' in str(row['납부방법']) and row['수입'] > 0:
        return 108  # C50 외상매출금 셀의 값을 참조해야 하므로
    elif (
        ('대체' in str(row['납부방법']) and '해약' in str(row['계정과목'])) or
        ('대체' in str(row['납부방법']) and '사용선수' in str(row['회계계정'])) or
        ('이체' in str(row['납부방법']) and '해약' in str(row['계정과목']))
    ):
        return row['코드1']
    elif '신용' in str(row['납부방법']) and row['수입'] > 0:
        return 120  # C20 미수금 셀의 값을 참조해야 하므로
    elif '현금' in str(row['납부방법']) and row['수입'] > 0:
        return 103  # C73 현금 셀의 값을 참조해야 하므로
    else:
        return 0

# '코드3' 열을 계산하여 추가
merged_df['코드2'] = merged_df.apply(calculate_code2, axis=1)

# 엑셀 파일 불러오기

code_df = pd.read_excel('C:\dev\myproj01\금상동\결과폴더\금상동계정코드.xlsx')

# 금상동 계정코드 파일을 딕셔너리로 변환
code_dict = code_df.set_index(code_df.columns[0]).to_dict()[code_df.columns[1]]



# 구분2 값을 계산하는 함수 정의
def calculate_gubun2(row):
    if (
        ('대체' in str(row['납부방법']) and '해약' in str(row['계정과목'])) or
        ('대체' in str(row['납부방법']) and '사용수입' in str(row['회계계정']) and row['지출'] > 0) or
        ('대체' in str(row['납부방법']) and '관리선수' in str(row['회계계정']) and row['지출'] > 0) or
        ('대체' in str(row['납부방법']) and '사용선수' in str(row['회계계정'])) or
        ('이체' in str(row['납부방법']) and '해약' in str(row['계정과목']))
    ):
        return 3
    elif row['코드1'] == 0:
        return 0
    elif row['구분1'] == 2:
        return 0
    else:
        return row['구분1'] - 1






merged_df['구분2'] = merged_df.apply(calculate_gubun2, axis=1)











# 결과를 엑셀 파일로 저장
#merged_df.to_excel('C:\\dev\\myproj01\\금상동\\결과폴더\\금상동월별결산부2024업로드.xlsx', index=False)


code_df1=code_df[['코드2','계정코드2']]

# VLOOKUP 함수 구현
def vlookup(lookup_value, code_df1, col_index, range_lookup=False):
    # 'range_lookup'에 따라 정확히 일치하는 값 또는 근사치를 찾을 수 있지만,
    # 현재는 정확한 일치를 기준으로 구현
    result = code_df1[code_df1.iloc[:, 0] == lookup_value]
    if not result.empty:
        return result.iloc[0, col_index - 1]  # col_index는 1부터 시작
    return 0 if not range_lookup else "Approximate Not Found"



# VLOOKUP을 적용할 데이터
lookup_col = '코드2'  # VLOOKUP에 사용할 열 이름
value_col = '계정코드2'      # 반환할 열 이름
result_col = '계정코드3'  # 결과를 저장할 열 이름

# VLOOKUP 함수 적용
merged_df[result_col] = merged_df[lookup_col].apply(lambda x: vlookup(x, code_df1, 2))

#merged_df['금액2']=None
merged_df['금액2']=merged_df['금액'].astype(int)


# 금액2 열을 맨 끝으로 이동시킵니다
# 현재 열 순서를 얻습니다
columns = list(merged_df.columns)

# 금액2가 맨 끝으로 오도록 열 순서를 재배열합니다
new_columns_order = [col for col in columns if col != '금액2'] + ['금액2']

# DataFrame을 새로운 열 순서로 재정렬합니다
merged_df=merged_df[new_columns_order]





merged_df['번호']=merged_df['Unnamed: 0']


print(merged_df)
#merged_df.info()

merged_df.to_excel('C:\dev\myproj01\금상동\결과폴더\금상동월별결산부2023업로드.xlsx', index=False)
upload_df = pd.DataFrame()
dae_df=merged_df[['번호','월일','구분1','코드1','계정코드2','적요명','금액']]
cha_df=merged_df[['번호','월일','구분2','코드2','계정코드3','적요명','금액2']]

cha_df['구분1']=cha_df['구분2']
cha_df['코드1']=cha_df['코드2']
cha_df['계정코드2']=cha_df['계정코드3']
cha_df['금액']=cha_df['금액2']


#merged_df = pd.merge(mon_df, account_code_df[['코드2','계정코드2']], how='left',left_on='코드1',right_on='코드2')

#upload_df = pd.merge(cha_df, dae_df, how='left',left_on='구분1',right_on='구분1')



cha_df2=cha_df[['번호','월일','구분1','코드1','계정코드2','적요명','금액']]







#dae_df.info()
cha_df2.to_excel('C:\dev\myproj01\금상동\결과폴더\차변업로드.xlsx', index=False)
dae_df.to_excel('C:\dev\myproj01\금상동\결과폴더\대변업로드.xlsx', index=False)
#upload_df = pd.merge(cha_df, dae_df[['월일','구분2','코드2','계정코드3']], how='left',left_on='월일',right_on='코드2')

#merged_df = pd.merge(mon_df, account_code_df[['코드2','계정코드2']], how='left',left_on='코드1',right_on='코드2')




# 데이터프레임 1과 2의 행 수가 같은지 확인합니다
#assert len(cha_df) == len(dae_df), "두 데이터프레임의 행 수가 다릅니다."

# 빈 데이터프레임을 만듭니다
upload_df = pd.DataFrame()

# 교차로 행을 추가합니다
for i in range(len(cha_df)):
    upload_df= pd.concat([upload_df, cha_df.iloc[[i]], dae_df.iloc[[i]]], ignore_index=True)

#upload_df['차대']=cha_df['구분1']+dae_df['구분2']
#upload_df=upload_df.drop(columns=['구분1'], inplace=True)
#upload_df.info()



#mon_df['구분1'] > 1) & (mon_df['코드1'] > 0)


#mon_df = mon_df[mon_df['구분1'] >= 1]
#upload_df['구분1']=upload_df['구분1'].astype(int)

upload_df=upload_df[['월일','구분1','코드1','계정코드2','적요명','금액']]


#upload_df2.info()

upload_df2= upload_df[(upload_df['구분1'] >= 1) & (upload_df['코드1'] > 0)] #구분1 값이 1보다 큰값은 남기고 코드1값이 0 이면 삭제하기위함


#upload_df2= upload_df[upload_df['구분1'] >= 1]


upload_df2.info()

upload_df2.insert(4, '코드4', '')      # 6th column, index 5
upload_df2.insert(5, '거래처명', '')   # 7th column, index 6
upload_df2.insert(6, '코드5', '')      # 8th column, index 7








upload_df2.to_excel('C:\dev\myproj01\금상동\결과폴더\금상동세무사랑2024업로드.xlsx', index=False)


# 1. 엑셀 파일 불러오기
input_file = 'C:\dev\myproj01\금상동\결과폴더\금상동세무사랑2024업로드.xlsx'  # 여기에 불러올 엑셀 파일의 경로를 입력하세요.
df7 = pd.read_excel(input_file)

df7.rename(columns={'월일' : '년도월일'}, inplace=False)
df7.rename(columns={'계정코드2' : '계정과목'}, inplace=False)
df7.rename(columns={'코드4' : '코드'}, inplace=False)







# 2. 데이터프레임을 엑셀 파일로 저장 (임시 저장)
output_file = 'C:\dev\myproj01\금상동\결과폴더\금상동세무사랑기타업로드.xlsx'  # 여기에 저장할 엑셀 파일의 경로를 입력하세요.
df7.to_excel(output_file, index=False)

# 3. openpyxl을 사용하여 엑셀 파일 열기
wb = load_workbook(output_file)
ws = wb.active

# 4. 열 너비 설정
ws.column_dimensions['D'].width = 20  # '구분1' 열 (가정: A열에 위치)
ws.column_dimensions['H'].width = 25  # '적요명' 열 (가정: B열에 위치)

# 5. 엑셀 파일 저장
wb.save(output_file)