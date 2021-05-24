import pandas as pd

path = r"data_output\\snmp_location_compiled.xlsx"
writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
room_df_ = pd.read_excel('src\\Sites list_room.xlsx', sheet_name='Site List with Address')
room_df__ = pd.read_excel('src\\SITE ADDRESS.xlsx', sheet_name='Sheet1')
room_df = pd.merge(room_df_,room_df__, how='left', on=["SITE NAME", "SITE NAME"])
room_df = room_df.dropna(subset=['SITE NAME'])
print(room_df.columns)
for site in room_df['SITE NAME']:
    # try:
    room = room_df.loc[room_df['SITE NAME'] == site, 'Room'].iloc[0]
    print(room)
    address = room_df.loc[room_df['SITE NAME'] == site, 'SITE ADDRESS'].iloc[0]
    print(address)
    REGION_x = room_df.loc[room_df['SITE NAME'] == site, 'REGION_x'].iloc[0]
    print(REGION_x)
    PROVINCE_x = room_df.loc[room_df['SITE NAME'] == site, 'PROVINCE_x'].iloc[0]
    print(PROVINCE_x)
    long = room_df.loc[room_df['SITE NAME'] == site, 'LONG'].iloc[0]
    lat = room_df.loc[room_df['SITE NAME'] == site, 'LAT'].iloc[0]
    room_df.loc[room_df['SITE NAME'] == site, ['SNMP_LOC']] = str(room) + \
        "/" + str(address) + "/" + str(PROVINCE_x) + "/" + str(REGION_x) + "(" + \
        str(long) + "," + str(lat) + ")"
    # except:
    #     continue
print(room_df.head(100))
write_df =room_df[["SITE NAME","SNMP_LOC"]]
pd.DataFrame(write_df).to_excel(writer, sheet_name="compiled",index=False)
writer.save()