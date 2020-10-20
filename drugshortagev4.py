import requests,json, credentials,datetime, openpyxl,time, os

LOGIN_URL = 'https://www.drugshortagescanada.ca/api/v1/login'
SEARCH_URL =  'https://www.drugshortagescanada.ca/api/v1/search?din='

class Session:
    def __init__(self,email,password):
        self.email = email
        self.password = password
        self.back_order = []
        self.din_list = []
        self.num = 1
        self.row_num = 2

        self.today = datetime.datetime.now().strftime('%Y-%m-%d')
        self.workbook = openpyxl.Workbook()
        self.sheet = self.workbook.active
        headings = ['Date','DIN','Médicament','ID du rapport', 'Date initiale', 'MAJ', 'Date estimée de fin', 'Raison rupture']
        font = openpyxl.styles.Font(size = 14, bold = True)
        for i in range(len(headings)):
            self.sheet.cell(row = 1, column = i+1).value = headings[i]
            self.sheet.cell(row = 1, column = i+1).font = font

    #logs in the website and returns and auth-token which will be used later on
    def login(self):
        self.response = requests.post(LOGIN_URL, data = {'email': self.email, 'password': self.password})
        self.response.raise_for_status()
        self.data = self.response.json()
        self.auth_token = self.response.headers['auth-token']

    #gets the DIN from the .xlsx provided list. Puts then in self.din_list
    def get_din(self,filename):
        wb = openpyxl.load_workbook(filename)
        ws = wb.active
        #find column containing DIN and get all DINs
        maxcol = ws.max_column
        maxrow = ws.max_row
        for num in range (1, maxcol+1):
             if 'DIN' in ws.cell(row = 1, column = num).value:
                     col = openpyxl.utils.cell.get_column_letter(num)
        din_range = ws[f'{col}2':f'{col}{str(maxrow)}'] #first cell of range is name of column so i'm not including it
        partial_din_list = [din[0].value for din in din_range if din[0].value != '' if type(din[0].value) == str] #there appears to be some nonetypes

        #some DINs may not be well entered so fix these lines are meant to fix the list
        din_list_len8 = [din for din in partial_din_list if len(din) == 8]
        din_list_lenunder8 = ['0'*(8-len(din))+din for din in partial_din_list if len(din) < 8]
        self.din_list = din_list_len8 + din_list_lenunder8

        return self.din_list

        #requests info from api and puts them in sel.back_orders which will be iterated through in writexlsx().
    def searchandwrite(self,din):
        headers = {'auth-token' : self.auth_token}
        FULL_SEARCH_URL = SEARCH_URL + din + '&filter_status=active_confirmed'
        self.req = requests.get(FULL_SEARCH_URL, headers = headers)
        self.req.raise_for_status()
        self.search_data = self.req.json()
        print(f'Getting info for DIN #{self.num} {din}...')

        # self.back_order = [item for item in self.search_data['data'] if item['resolved'] == False if item['status'] != 'discontinued']
        self.back_order = [item for item in self.search_data['data']] # veut-on exclure les rx 'resolved' = True?
        if len(self.back_order) == 1:
            print(f'{len(self.back_order)} backorder found for this drug')
        elif len(self.back_order) > 1:
            print(f'{len(self.back_order)} backorders found for this drug')

        for bo in self.back_order:
            print(f"Writing info for backorder #{self.row_num-1} - DIN {bo['din']}...")
            #figuring out a way to handle KeyError
            bo_keys = ['din','fr_drug_brand_name','drug_strength','drug_dosage_form_fr','drug_package_quantity',
                        'id','created_date','updated_date','estimated_end_date','shortage_reason','resolved','status']
            values_dict = {}
            for bo_key in bo_keys:
                if bo_key in bo.keys():
                    values_dict[bo_key] = bo[bo_key]
                else :
                    if bo_key == 'estimated_end_date':
                        if bo['unknown_estimated_end_date'] == True:
                            values_dict['estimated_end_date'] = 'InconnueT' #see values_list to understand why the end T is there
                        else:
                            values_dict['estimated_end_date'] = 'ERROR'
                    elif bo_key == 'updated_date':
                        values_dict['updated_date'] = 'Pas encore mis à jourT'
                    elif bo_key == 'shortage_reason':
                        values_dict['shortage_reason'] = {'fr_reason':'Non inclus'}
                    elif bo_key == 'drug_strength':
                        values_dict['drug_strength'] = '*Dose non fournie*'
                    else:
                        values_dict[bo_key] = 'ERROR'

            values_list = [self.today, values_dict['din'],f"{values_dict['fr_drug_brand_name']} {values_dict['drug_strength']} {values_dict['drug_dosage_form_fr']} {values_dict['drug_package_quantity']}",
                            values_dict['id'], values_dict['created_date'].split('T')[0], values_dict['updated_date'].split('T')[0],values_dict['estimated_end_date'].split('T')[0],
                            values_dict['shortage_reason']['fr_reason']]

            print(values_dict)

            #writes to excel
            for col_num in range(len(values_list)):
                self.sheet.cell(row = self.row_num, column = col_num + 1).value = values_list[col_num]
            self.row_num += 1

        filename = f'Ruptures du {self.today} try.xlsx'
        if self.num % 100 == 0 :
            print('Saving worksheet.')
            self.workbook.save(filename)
        if self.num % 900 == 0: #call limit is set at 1000/hour so setting a in case other people are making requests.
            print('Taking an hour break as call limit is set at 1000 per hour.')
            start = time.time()
            end = start + 3600
            while time.time() < end: #will leave the s at the end of the statement when remaining time goes < 10 minutes (so 2 digits). Don't know how to fix it yet.
                message = f'Restarting in {round((end-time.time())/60)-1} minutes'
                print(message, end = '')
                time.sleep(60)
                print('\b'* len(message), end = '', flush = True)
        if self.num == len(self.din_list):
            print('Saving worksheet for last time.')
            self.workbook.save(filename)
        self.num += 1

user = Session(credentials.email,credentials.password)
user.login()
for din in user.get_din('list.xlsx')[:100]:
    user.searchandwrite(din)
print('Done.')
