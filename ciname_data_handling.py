import pandas as pd
from datetime import date


df_movie_target = pd.read_excel(r'C:\Users\hanbi01\Desktop\한빛누리\제일기획\CAMS 타깃별.xlsx')
df_cams_week = pd.read_excel(r'C:\Users\hanbi01\Desktop\한빛누리\제일기획\CAMS 광고주별.xlsx', sheet_name='MF1359')
today = date.today().strftime('%Y%m%d')
excel_writer = pd.ExcelWriter(r'C:\Users\hanbi01\Desktop\한빛누리\제일기획\(극장) 광고비 및 광고노출 데이터_%s.xlsx' % today, engine='xlsxwriter')


# preprocessing in target
df_movie_target['성연령'].astype(str)
df_target = df_movie_target[['월','성연령','성연령비','모집단','AF']]


# count the number of targets (monthly)
monthly_target_count = df_target.pivot_table('성연령비', columns=['월'], aggfunc='count', fill_value=0)
target_count = monthly_target_count.values.tolist()[0]  # dataframe 내 데이터를 list로 변환


# grouping target by month
df_target_1 = df_target.iloc[:target_count[0]]
df_target_2 = df_target.iloc[target_count[0]:target_count[0]+target_count[1]]
df_target_3 = df_target.iloc[target_count[0]+target_count[1]:]

df_target_0 = [df_target_1, df_target_2, df_target_3]  # 리스트에 담아서 for 문에서 하나씩 꺼내감


# preprocessing in advertiser 
df_movie_advertiser = df_cams_week[['상품별광고주','상품별', '날짜', '광고주', 'cost', '광고시청자','Share','week weight']]


# count the number of advertisers (monthly)
monthly_count = df_movie_advertiser.pivot_table('광고주', columns=['날짜'], aggfunc='count', fill_value=0)
advertiser_count = monthly_count.values.tolist()[0]  # dataframe to list


# grouping advertisers by month
df_advertiser_1 = df_movie_advertiser.iloc[:advertiser_count[0]]
df_advertiser_2 = df_movie_advertiser.iloc[advertiser_count[0]:advertiser_count[0]+advertiser_count[1]]
df_advertiser_3 = df_movie_advertiser.iloc[advertiser_count[0]+advertiser_count[1]:]

df_advertiser_0 = [df_advertiser_1, df_advertiser_2, df_advertiser_3]  # 리스트에 담아서 for 문에서 하나씩 꺼내감


for j in range(1):
    df_ad = pd.concat([df_advertiser_0[j]] * 45, axis=0, ignore_index=True)
    df_ad = df_ad.sort_values(['상품별광고주'], ascending=[True])
    df_ad = df_ad.reset_index(drop=True)

    df_age = pd.concat([df_target_0[j]] * (advertiser_count[j]), axis=0, ignore_index=True)
    df_test_1 = pd.concat([df_ad, df_age], axis=1)


for j in range(2):
    df_ad = pd.concat([df_advertiser_0[j]] * 45, axis=0, ignore_index=True)
    df_ad = df_ad.sort_values(['상품별광고주'], ascending=[True])
    df_ad = df_ad.reset_index(drop=True)

    df_age = pd.concat([df_target_0[j]] * (advertiser_count[j]), axis=0, ignore_index=True)
    df_test_2 = pd.concat([df_ad, df_age], axis=1)


for j in range(3):
    df_ad = pd.concat([df_advertiser_0[j]] * 45, axis=0)
    df_ad = df_ad.sort_values(['상품별광고주'], ascending=[True])
    df_ad = df_ad.reset_index(drop=True)

    df_age = pd.concat([df_target_0[j]] * (advertiser_count[j]), axis=0, ignore_index=True)
    df_test_3 = pd.concat([df_ad, df_age], axis=1)

df_final_data_set = pd.concat([df_test_1, df_test_2, df_test_3], axis=0, ignore_index=True)


# making new variables
df_final_data_set['신af'] = df_final_data_set['AF']* (df_final_data_set['Share']/100) * df_final_data_set['week weight']
df_final_data_set.loc[df_final_data_set['신af'] < 1, '신af'] = 1
df_final_data_set['impression'] = round(df_final_data_set['광고시청자']*df_final_data_set['성연령비'])
df_final_data_set['GRP'] = df_final_data_set['impression']/df_final_data_set['모집단']*100
df_final_data_set['Reach'] = df_final_data_set['GRP']/df_final_data_set['신af']

# select variables in new variables only included in excel file
df_final_data_set = df_final_data_set[['월', '상품별','광고주','cost','성연령','impression','GRP','Reach']]
df_final_data_set = df_final_data_set.sort_values(['성연령'], ascending=[True])  # 성연령 기준으로 sorting
df_final_data_set.to_excel(excel_writer, sheet_name='data', index=False, index_label=False)  # index 제거


# Excel style
workbook = excel_writer.book
worksheet = excel_writer.sheets['data']

style = workbook.add_format({
    'bold': False,
    'text_wrap': False})  # 볼드체/표그리기 X

for col_num, value in enumerate(df_final_data_set.columns.values):
    worksheet.write(0, col_num, value, style)

format1 = workbook.add_format({'num_format': '#,##0'})  # 소숫점 천의자리 구분자
worksheet.set_column('D:D', 15, format1)  # (column, column width, format style)
worksheet.set_column('F:F', 15, format1)
worksheet.set_column('B:B', 12)
worksheet.set_column('C:C', 25)


excel_writer.save()
