import pandas as pd
from datetime import datetime

# 여러 개의 데이터프레임을 엑셀 파일로 내보내기 함수
def export_multiple_dfs_to_excel(dfs, sheet_names, file_name):
    """
    여러 개의 데이터프레임을 각각 다른 시트에 엑셀 파일로 내보내는 함수.
    : param dfs: 데이터프레임 리스트 
    : param sheet_names: 시트 이름 리스트 
    : param file_name: 저장할 엑셀 파일 이름
   """ 
    
    if len(dfs) != len(sheet_names) :

      raise ValueError("데이터프레임 리스트와 시트 이름 리스트의 길이가 같아야 합니다. ")


    with pd. ExcelWriter (file_name) as writer:
     for df, sheet_name in zip(dfs, sheet_names) :
        df. to_excel(writer, sheet_name=sheet_name,index=False)
    #Print(f"{file_name} 파일이 생성되었습니다.")


# 엑셀 파일 읽기
file_path = 'C:\dev\myproj01\회사등록_엑셀간편저장_241115.xls'  # 여기에 원본 엑셀 파일 경로를 입력하세요.
df = pd.read_excel(file_path)

df['회사코드2']=df['회사코드']
df['상호2']=df['상호']
df['사업자등록번호2']=df['사업자등록번호']


df5=df
df5.info()
# 대표자명 열과 사업자번호 열이 존재한다고 가정합니다. 그렇지 않으면 아래에서 열 이름을 변경하세요.
representative_col = '대표자명'
business_number_col = '사업자등록번호'

# 대표자명을 기준으로 내림차순 정렬
df = df.sort_values(by=representative_col, ascending=False)




# 대표자명이 같은 경우, 대표자명1, 대표자명2, 대표자명3 등으로 변경

df[representative_col] = df[representative_col].astype(str)

df[representative_col] = df.groupby(representative_col).cumcount().add(1).astype(str).radd(df[representative_col])









# 정렬된 데이터프레임을 엑셀 파일로 저장
sorted_by_representative_path = 'sorted_by_representative.xlsx'
df.to_excel(sorted_by_representative_path, index=False)

# 사업자번호 열을 기준으로 내림차순 정렬
df = df.sort_values(by=business_number_col, ascending=True)


# 대표자명으로 내림차순 정렬
#df2 = df.sort_values(by=representative_col, ascending=True)




# 최종 정렬된 데이터프레임을 엑셀 파일로 저장
final_sorted_path = 'C:\dev\myproj01\회사목록20241115.xlsx'
#df.to_excel(final_sorted_path, sheet_name='사업자번호순',index=False)
#df2.to_excel(final_sorted_path, sheet_name='대표자명순',index=False)



df1 =df.sort_values(by=business_number_col, ascending=True)
df2 =df.sort_values(by=representative_col, ascending=True)


df3= df1.dropna(subset=["사업자등록번호"]) #아래 누적액,월계액을 없애기위해 회계계정없는부분을 삭제함

df4= df2.dropna(subset=["대표자주민등록번호"]) #아래 누적액,월계액을 없애기위해 회계계정없는부분을 삭제함





#df3 =pd.DataFrame (data3)
# 데이터프레임과 시트 이름 리스트
dfs = [df3, df4,df5]
sheet_names = ['사업자번호순', '대표자명순','세무사랑원본']
# 함수 호출

today = datetime.today().strftime('%Y-%m-%d')

output_file = f'C:\dev\myproj01\회사목록-{today}.xlsx'


#'C:\dev\myproj01\회사목록20241031.xlsx'

export_multiple_dfs_to_excel(dfs, sheet_names,output_file)