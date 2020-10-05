import pandas as pd 
import os
import xlsxwriter
from datetime import date

def rename_file():
	path='./'
	dirs = os.listdir(path)
	for file in dirs:
		if os.path.isfile(os.path.join(path,file)) and 'Tokopedia' in file:
			os.rename(file,"zzTokopedia.xlsx")
		elif os.path.isfile(os.path.join(path,file)) and 'transaction' in file:
			os.rename(file,"zzBukalapak.xlsx")
		elif os.path.isfile(os.path.join(path,file)) and 'Order.' in file:
			os.rename(file,"zzShopee.xls")

rename_file()

today = date.today()
today = today.strftime("%d %B %Y")
tanggal = pd.DataFrame({'A':['Tanggal :'+ today]})

template_column = list(['Kurir','Cek', 'Pembeli','Ongkir','Status','OrderID','Kode'])
all_data = pd.DataFrame()

#Excel Data into dataframe 
TP_file = r'zzTokopedia.xlsx'
TP_df = pd.read_excel(TP_file)

BL_file = r'zzBukalapak.xlsx'
BL_df = pd.read_excel(BL_file)

S_file = r'zzShopee.xls'
S_df = pd.read_excel(S_file)

#rename the file to delete it 
os.rename(TP_file,"Processed_TP.xlsx")
os.rename(BL_file,"Processed_BL.xlsx")
os.rename(S_file,"Processed_S.xls")


# Remove Duplicate data and useless columns TOKOPEDIA
# Total Kolom Tokopedia = 26
# Important index column :
# 1: Order ID
# 4: Order Status
# 13 :Recipient 
# 16: Courier
# 17: Ongkir 

# use iloc to remove unnecasry data 
TP_df = TP_df.iloc[3:]

TP_total_col = 26
list_col_keep_TP = [1,4,13,16,17]
list_col_del_TP = []

for index in range(TP_total_col):
	if index not in list_col_keep_TP:
		list_col_del_TP.append(index)

#drop useless column
TP_df.drop(TP_df.columns[list_col_del_TP], axis = 1 , inplace = True)
TP_df.columns = ['OrderID','Status','Pembeli','Kurir','Ongkir']
TP_df.drop_duplicates(subset="OrderID",keep = 'first',inplace = True)

#filter the data with 'Status Pesanan' = perlu dikirim 
criteria = TP_df[(TP_df['Status'] != 'Pemesanan sedang diproses oleh penjual.')].index
TP_df = TP_df.drop(criteria)
criteria = TP_df[
	(TP_df.Kurir != 'JNE(Reguler)') & 
	(TP_df.Kurir != 'JNE(YES)') & 
	(TP_df.Kurir !='SiCepat(Regular Package)') & 
	(TP_df.Kurir != 'SiCepat(BEST)') & 
	(TP_df.Kurir != 'J&T(Reguler)') 
].index
TP_df = TP_df.drop(criteria)

#add new column 'Kode'
TP_df['Kode'] = 'TP'
TP_df['Cek']=''
TP_df = TP_df[template_column]
#add this to all_data
all_data = all_data.append(TP_df)[TP_df.columns.tolist()]

# Remove Duplicate data and useless column BUKALAPAK
# Total kolom Bukalapak = 24
# 1: ID Transaksi 
# 6: Pembeli 
# 16:Biaya Pengiriman 
# 22: Kurir
# 24: Status 

BL_total_col = 25
list_col_keep_BL = [1,6,16,23,25]
list_col_del_BL = []

for index in range(BL_total_col):
	if index not in list_col_keep_BL:
		list_col_del_BL.append(index)

BL_df.drop_duplicates(subset="ID Transaksi",keep = 'first',inplace = True)
BL_df.drop(BL_df.columns[list_col_del_BL], axis = 1 , inplace = True)

#rename the column 
BL_df.columns = ['OrderID' , 'Pembeli', 'Ongkir','Kurir', 'Status']

#filter the data with 'Status Pesanan' = perlu dikirim 
criteria = BL_df[(BL_df['Status'] != 'Diproses Pelapak')].index
BL_df = BL_df.drop(criteria)
criteria = BL_df[
	(BL_df.Kurir != 'JNE REG') & 
	(BL_df.Kurir != 'JNE YES') & 
	(BL_df.Kurir !='SiCepat REG') & 
	(BL_df.Kurir != 'SiCepat BEST') & 
	(BL_df.Kurir != 'J&T REG')
].index
BL_df = BL_df.drop(criteria)

#add new column 'Kode'
BL_df['Kode'] = 'BL'
BL_df['Cek']=''
BL_df = BL_df[template_column]
#add this to all_data
all_data = all_data.append(BL_df)[BL_df.columns.tolist()]

# Remove Duplicate data and useless column SHOPEE
# 0: no pesanan 
# 1:status pesanan 
# 4: Opsi Pengiriman 
# 34: Perkiraan Ongkos Kirim 
# 38: Nama Penerima  
S_total_col = 44
list_col_keep_S = [0,1,4,34,38]
list_col_del_S = []

for index in range(S_total_col):
	if index not in list_col_keep_S:
		list_col_del_S.append(index)

S_df.drop_duplicates(subset="No. Pesanan", keep ='first', inplace = True)
S_df.drop(S_df.columns[list_col_del_S], axis = 1, inplace=True)

S_df.columns = ['OrderID','Status','Kurir','Ongkir','Pembeli']

criteria = S_df[(S_df['Status'] !='Perlu Dikirim')].index
S_df = S_df.drop(criteria)
criteria = S_df[
	(S_df.Kurir != 'J&T Express') & 
	(S_df.Kurir != 'JNE REG') & 
	(S_df.Kurir != 'JNE YES') & 
	(S_df.Kurir != 'SiCepat REG') &
	(S_df.Kurir != 'JNE Reguler (Cashless)')
].index
S_df = S_df.drop(criteria)

S_df['Kode'] = 'S'
S_df['Cek']=''
S_df = S_df[template_column]

all_data = all_data.append(S_df)[S_df.columns.tolist()]

#save the paper
all_data['Kurir'] = all_data['Kurir'].replace('SiCepat(Regular Package)','SiCepat')
all_data['Kurir'] = all_data['Kurir'].replace('SiCepat REG','SiCepat')
all_data['Kurir'] = all_data['Kurir'].replace('SiCepat BEST','SiCepat')
all_data['Kurir'] = all_data['Kurir'].replace('SiCepat(BEST)','SiCepat')
all_data['Kurir'] = all_data['Kurir'].replace('J&T Express','J&T')
all_data['Kurir'] = all_data['Kurir'].replace('J&T REG','J&T')
all_data['Kurir'] = all_data['Kurir'].replace('J&T(Reguler)','J&T')
all_data['Kurir'] = all_data['Kurir'].replace('JNE REG','JNE')
all_data['Kurir'] = all_data['Kurir'].replace('JNE YES','JNE')
all_data['Kurir'] = all_data['Kurir'].replace('JNE(YES)','JNE')
all_data['Kurir'] = all_data['Kurir'].replace('JNE(Reguler)','JNE')
all_data['Kurir'] = all_data['Kurir'].replace('JNE Reguler (Cashless)', 'JNE')

all_data['Pembeli.lower'] = all_data['Pembeli'].str.lower()
all_data = all_data.sort_values(by=['Kurir','Pembeli.lower','Kode'])
all_data = all_data.drop('Pembeli.lower',1)
all_data = all_data.reset_index(drop=True)
all_data.drop(['OrderID','Status'], axis = 1, inplace=True)

# GROUPING 
test = all_data
grouped = test.groupby('Kurir', as_index = False)
grouped_index = grouped.apply(lambda x: x.reset_index(drop = True)).reset_index()
# have 2 level and drop 1 column
all_data=grouped_index
all_data.drop(['level_0'],axis=1 , inplace= True)

all_data.columns = ['No','Kurir' ,'Cek','Pembeli', 'Ongkir','Kode']
all_data.No+=1

# print(all_data)

# Dataframe to excel
filename = './order-list_'+ today+'.xlsx'

writer = pd.ExcelWriter(filename,engine ='xlsxwriter')
tanggal.to_excel(writer , sheet_name='sheetTest',header = False,startrow= 0, startcol=3,index = False)
all_data.to_excel(writer,sheet_name='sheetTest',startrow = 1,index = False)

workbook = writer.book
worksheet = writer.sheets['sheetTest']

#KOLOM 
# A:No
# B:Kurir
# C:Cek 
# D:Pembeli
# E:Ongkir 
# F:Kode 

# Set Column width
worksheet.set_column('A:A',5)
worksheet.set_column('B:B',8)
worksheet.set_column('C:C',5)
worksheet.set_column('D:D',35)
worksheet.set_column('E:E',10)
worksheet.set_column('F:F',5)

#Excel Border Format 
border_format = workbook.add_format({
		'border':1,
	})
thick_border_format = workbook.add_format({
		'border':2,
	})

# applying the border format 
worksheet.conditional_format(1,0,1,5,{'type':'no_blanks', 'format':thick_border_format})
worksheet.conditional_format(2,0,len(all_data)+1,5,{'type':'no_blanks', 'format':border_format})
worksheet.conditional_format(2,0,len(all_data)+1,5,{'type':'blanks', 'format':border_format})

writer.save()

#remove unnecessary file 
deletePath=  './'
TP_file = os.path.join(deletePath, "Processed_TP.xlsx")
BL_file = os.path.join(deletePath, "Processed_BL.xlsx")
S_file = os.path.join(deletePath, "Processed_S.xls")
os.remove(TP_file)
os.remove(BL_file)
os.remove(S_file)