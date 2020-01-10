import pandas as pd
from openpyxl import load_workbook


# tv raw 데이터 pandas로 불러와서 엑셀로 가공
def tv_raw_data(read_data_1, read_data_2, sheet_name):
    global df_TV_target_1, df_TV_target_2, df_tv_target
    df_TV_target_1 = pd.read_excel(r'Q:\조희진\이노션 TAR\이노션db입력\TVraw\%s' % read_data_1, skiprows=2)
    df_TV_target_2 = pd.read_excel(r'Q:\조희진\이노션 TAR\이노션db입력\TVraw\%s' % read_data_2, skiprows=2)
    df_tv_target = pd.concat([df_TV_target_1, df_TV_target_2], axis=1)
    df_tv_target.to_excel(excel_writer, sheet_name, index=False)


# tv raw 데이터 복사해서 작업 파일에 붙여넣기
def copy_and_paste_data(sheet_number, sheet_name):
    global read_excel_file, worksheet_read, write_excel_file, worksheet_write, row, col, array, inner_array, i, j, value
    read_excel_file = load_workbook(r'Q:\조희진\이노션 TAR\이노션db입력\(이노션 TAR) 2S Usage_TV_raw.xlsx')
    worksheet_read = read_excel_file.worksheets[sheet_number]

    row = 2
    col = 1
    last_col = worksheet_read.max_column
    last_row = worksheet_read.max_row
    array = []
    inner_array = []
    for i in range(last_row-1):
        for j in range(last_col):
            value = worksheet_read.cell(row=row, column=col).value
            inner_array.append(value)
            col += 1

        array.append(inner_array)
        inner_array = []
        col = 1
        row += 1

    write_excel_file = load_workbook(filename=r'Q:\조희진\이노션 TAR\이노션db입력\(이노션 TAR) 2S Usage_TV.xlsx')
    worksheet_write = write_excel_file[sheet_name]
    row = 5
    col = 1
    for i in range(last_row-1):
        for j in range(last_col):
            test = worksheet_write.cell(row=row, column=col)
            test.value = array[i][j]
            col += 1
        col = 1
        row += 1

    write_excel_file.save(r'Q:\조희진\이노션 TAR\이노션db입력\(이노션 TAR) 2S Usage_TV.xlsx')


excel_writer = pd.ExcelWriter(r'Q:\조희진\이노션 TAR\이노션db입력\(이노션 TAR) 2S Usage_TV_raw.xlsx', engine='openpyxl')
tv_raw_data('1.xls','1_1.xls','1_타깃별')
tv_raw_data('1_2.xls','1_3.xls','2.채널별')
tv_raw_data('1_4.xls','1_5.xls','3.시간대별 ')
tv_raw_data('1_6.xls','1_7.xls','3.시간대별_ALL')
tv_raw_data('1_8.xls','1_9.xls','4.,주간별')
tv_raw_data('1_10.xls','1_11.xls','4.,주간별_all')
excel_writer.save()


copy_and_paste_data(0,'1_타깃별')
copy_and_paste_data(1,'2.채널별')
copy_and_paste_data(2,'3.시간대별 ')
copy_and_paste_data(3,'3.시간대별_ALL')
copy_and_paste_data(4,'4.,주간별')
copy_and_paste_data(5,'4.,주간별_all')



# 타겟 raw 만들기
df_digital_target_1 = pd.read_excel(r'Q:\조희진\이노션 TAR\이노션db입력\이노션 TAR 2S Usage_DGT.xlsx', sheet_name='1. TARGET', skiprows=2)
df_digital_target_2 = pd.read_excel(r'Q:\조희진\이노션 TAR\이노션db입력\이노션 TAR 2S Usage_DGT.xlsx', sheet_name='2. Digital Site', skiprows=2)
df_tv_target_1 = pd.read_excel(r'Q:\조희진\이노션 TAR\이노션db입력\(이노션 TAR) 2S Usage_TV.xlsx',
                                sheet_name='1_타깃별', skiprows=46, usecols='A:L')  # 이용할 데이터 (46 행부터, L열까지)
print(df_tv_target_1)
df_tv_target_2 = pd.read_excel(r'Q:\조희진\이노션 TAR\이노션db입력\(이노션 TAR) 2S Usage_TV.xlsx',
                                sheet_name='2.채널별', skiprows=81, usecols='B:M')  # 이용할 데이터 (81 행부터, M열까지)

# Digital 타겟 데이터 전처리
df_digital_target_2 = df_digital_target_2[df_digital_target_2['DEVICE'] =='Mobile∪PC']  # Mobile∪PC인 값만 불러오기
df_digital_target_2 = df_digital_target_2.reset_index(drop=True)
df_digital_target_2 = df_digital_target_2[['SITE_NAME', 'MONTHCODE', 'SEX',	'AGE', 'UV', 'AVG_DAILY_UV', 'TTS(MIN)',
                                           'AVG_DAILY_DT(MIN)', 'DAY_MONTH', 'AVG_DAILY_ATV', 'UNIVERSE',	'AVG_USER']]

df_digital_target_2.rename(columns={'SITE_NAME':'DEVICE'}, inplace=True)


df_target_final = pd.concat([df_digital_target_1, df_digital_target_2, df_tv_target_1, df_tv_target_2], ignore_index=True, sort=True)
df_target_final = df_target_final[['DEVICE', 'MONTHCODE', 'SEX', 'AGE', 'UV', 'AVG_DAILY_UV', 'TTS(MIN)',
                                    'AVG_DAILY_DT(MIN)', 'DAY_MONTH', 'AVG_DAILY_ATV', 'UNIVERSE',	'AVG_USER']]

round(df_target_final[['AVG_DAILY_DT(MIN)', 'AVG_DAILY_ATV', 'AVG_USER']], 5)

df_target_final.to_csv(r'Q:\조희진\이노션 TAR\이노션db입력\bytarget.csv', mode='w', index=False, encoding='ms949')


# 시간대 raw 만들기

df_digital_time = pd.read_excel(r'Q:\조희진\이노션 TAR\이노션db입력\이노션 TAR 2S Usage_DGT.xlsx', sheet_name='3. Time', skiprows=2)
df_tv_time_1 = pd.read_excel(r'Q:\조희진\이노션 TAR\이노션db입력\(이노션 TAR) 2S Usage_TV.xlsx',
                                sheet_name='3.시간대별 ', skiprows=1736, usecols='A:N')  # 이용할 데이터 (1737 행부터, N열까지)
df_tv_time_2 = pd.read_excel(r'Q:\조희진\이노션 TAR\이노션db입력\(이노션 TAR) 2S Usage_TV.xlsx',
                                sheet_name='3.시간대별_ALL', skiprows=872, usecols='A:N')  # 이용할 데이터 (873 행부터, M열까지)

df_time_final = pd.concat([df_digital_time, df_tv_time_1, df_tv_time_2], ignore_index=True, sort=True)
df_time_final = df_time_final[['DEVICE', 'MONTHCODE', 'TIME_CD', 'WEEKDAY',	'SEX', 'AGE', 'UV', 'AVG_DAILY_UV', 'TTS(MIN)',
                               'AVG_DAILY_DT(MIN)', 'DAY_MONTH', 'AVG_DAILY_ATV', 'UNIVERSE','AVG_USER']]
round(df_time_final[['AVG_DAILY_DT(MIN)', 'AVG_DAILY_ATV', 'AVG_USER']], 5)
df_time_final.to_csv(r'Q:\조희진\이노션 TAR\이노션db입력\bytime.csv', mode='w', index=False, encoding='ms949')


# 주간별 raw 만들기

df_digital_day = pd.read_excel(r'Q:\조희진\이노션 TAR\이노션db입력\이노션 TAR 2S Usage_DGT.xlsx', sheet_name='4. DAY', skiprows=2)
df_tv_day_1 = pd.read_excel(r'Q:\조희진\이노션 TAR\이노션db입력\(이노션 TAR) 2S Usage_TV.xlsx',
                            sheet_name='4.,주간별', skiprows=81, usecols='A:M')  # 이용할 데이터 (82 행부터, M열까지)
df_tv_day_2 = pd.read_excel(r'Q:\조희진\이노션 TAR\이노션db입력\(이노션 TAR) 2S Usage_TV.xlsx',
                            sheet_name='4.,주간별_all', skiprows=45, usecols='A:M')  # 이용할 데이터 (46 행부터, M열까지)


df_day_final = pd.concat([df_digital_day, df_tv_day_1, df_tv_day_2], ignore_index=True, sort=True)
df_day_final = df_day_final[['DEVICE', 'MONTHCODE', 'WEEKDAY',	'SEX', 'AGE', 'UV', 'AVG_DAILY_UV', 'TTS(MIN)',
                              'AVG_DAILY_DT(MIN)', 'DAY_MONTH', 'AVG_DAILY_ATV', 'UNIVERSE','AVG_USER']]
round(df_day_final[['AVG_DAILY_DT(MIN)', 'AVG_DAILY_ATV', 'AVG_USER']], 5)
df_day_final.to_csv(r'Q:\조희진\이노션 TAR\이노션db입력\bydaytype.csv', mode='w', index=False, encoding='ms949')


# Digital 시간대/주간별 raw 만들기

df_digital_time_video = pd.read_excel(r'Q:\조희진\이노션 TAR\이노션db입력\이노션 TAR 2S Usage_DGT.xlsx', sheet_name='5. Time_Video', skiprows=2)
df_digital_day_video = pd.read_excel(r'Q:\조희진\이노션 TAR\이노션db입력\이노션 TAR 2S Usage_DGT.xlsx', sheet_name='6. DAY_Video', skiprows=2)

df_digital_day_video = df_digital_day_video[['DEVICE', 'MONTHCODE', 'WEEKDAY',	'SEX', 'AGE', 'UV', 'AVG_DAILY_UV',
                                              'TTS(MIN)', 'AVG_DAILY_DT(MIN)', 'DAY_MONTH', 'AVG_DAILY_ATV', 'UNIVERSE','AVG_USER']]

df_digital_time_video = df_digital_time_video[['DEVICE', 'MONTHCODE', 'TIME_CD', 'WEEKDAY',	'SEX', 'AGE', 'UV',
                                               'AVG_DAILY_UV', 'TTS(MIN)', 'AVG_DAILY_DT(MIN)', 'DAY_MONTH', 'AVG_DAILY_ATV', 'UNIVERSE', 'AVG_USER']]

round(df_digital_day_video[['AVG_DAILY_DT(MIN)', 'AVG_DAILY_ATV', 'AVG_USER']], 5)
df_digital_day_video.to_csv(r'Q:\조희진\이노션 TAR\이노션db입력\bydaytype_video.csv', mode='w', index=False, encoding='ms949')
round(df_digital_time_video[['AVG_DAILY_DT(MIN)', 'AVG_DAILY_ATV', 'AVG_USER']], 5)
df_digital_time_video.to_csv(r'Q:\조희진\이노션 TAR\이노션db입력\bytime_video.csv', mode='w', index=False, encoding='ms949')
