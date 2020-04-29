import pandas as pd
import sys
from babel.numbers import format_decimal
import json
from os import path
import os
import numpy as np


class ProcessData:
    def __init__(self, excel_file, json_file, dir_to_save, aida_export):
        self.excel_file = excel_file
        self.json_file = json_file
        self.dir_to_save = dir_to_save
        self.AIDAExport = aida_export
        self.type = "Dettagliato"
        self.col1 = None
        self.col2 = None
        self.col3 = None
                
    def convert_to_str(self, number):
        """
        :param number: the number that needs to be converted to string
        if the number is already in the format of str just returns it
        if the number is nan returns "n.a."
        if AIDAExport is enabled returns the same number with the string type in Italian format
        if AIDAExport is not enables returns the number/1000 with the string type
        :return: the string format of number
        """
        # if number is already a string returns it
        if isinstance(number, str):
            return number

        # if number is nan returns string "n.a."
        if np.isnan(number):
            return 'n.a.'

        # checks to set the proper output of string
        if self.AIDAExport:
            if float(number).is_integer():
                res_string = format_decimal(number, locale='it_IT') + ",00"
            else:
                print(number)
                res_string = format_decimal(float("{0:.2f}".format(number)), locale='it_IT', decimal_quantization=False)

        else:
            res_string = str(float("{0:.2f}".format(number / 1000)))

        return res_string
    
    def first_read_excel(self):
        """
        reads the excel file to see if it is Abbreviato or Dettagliato
        :return: the array of 3 sheet in excel file
        """
        # to get the sheet names of excel file ad set the type of survey
        xls = pd.ExcelFile(self.excel_file)
        try:
            if "bbreviato" in xls.sheet_names[1]:
                self.type = "Abbreviato"

        except Exception as e:
            print(e)
            sys.exit(1)

        # reads excel file with all sheets inside for later check
        df = pd.read_excel(self.excel_file, sheet_name=None)

        # depending the type of survey if it is Abbreviato or Dettagliato
        # reads the proper sheet_names
        if self.type == 'Abbreviato':
            first_df = df['SP_Civilistico_Abbreviato']
            second_df = df['CE_Civilistico_Abbreviato']
            third_df = df['CASH_FLOW']
        else:
            first_df = df['SP_Civilistico_Ordinario']
            second_df = df['CE_Civilistico_Ordinario']
            third_df = df['CASH_FLOW']

        return [first_df, second_df, third_df]

    def convert_data(self):
        """
        first calls first_read_excel to see create the df_arr
        that contains 3 sheets of data in the excel file
        :return: "Done" if every thing works fine!
        """
        try:
            df_arr = self.first_read_excel()
            if self.type == 'Abbreviato':
                self.read_data_abbreviato(df_arr)
            else:
                self.data_read_dettagliato(df_arr)
            # self.save_json()

            return 'Done'
        except Exception as e:
            print(e)
            return 'There is a problem'

    def save_json(self):
        """
        creates the directory first if doesn't exist and save the json file in the directory
        :return: saves the self.json_file in the result.json
        """
        self.create_dir()
        with open(self.dir_to_save+'/result_'+self.type+'_'+str(self.col1)+'.json', 'w') as json_file1:
            json.dump(self.json_file, json_file1, indent=2)

    def create_dir(self):
        """
        To save the json file we first check the destination directory where the file needs to be saved
        :return: creates dir_to_save if it doesn't exist
        """
        if not path.exists(self.dir_to_save):
            try:
                os.mkdir(self.dir_to_save)
            except OSError as excp:
                print("Creation of the directory '%s' failed" % self.dir_to_save)
                print(excp)
            else:
                print("Successfully created '%s' directory" % self.dir_to_save)
        
    def read_data_abbreviato(self, df_arr):
        """
        reads data and save it to self.json_file if
        the type of survey is Abbreviato
        :param df_arr: array that contains 3 sheets of survey
        :return: sets the new values in self.json_file and saves the json_file
        """
        data = self.json_file

        # for all the columns that contains year's data iterate
        # columns of years start from the 3rd column
        for i in range(len(df_arr[0].columns[2:])):
            self.col1 = df_arr[0].columns[i+2]
            SPATTLABELACREDSOCITOT = df_arr[0][self.col1][3]
            SPATTLABELBIMMOBIMMATTOT = df_arr[0][self.col1][5]
            SPATTLABELBIMMOBMATTOT = df_arr[0][self.col1][6]
            SPATTLABELBIMMOBFINANZTOT = df_arr[0][self.col1][7]
            SPATTLABELBIMMOBTOT = df_arr[0][self.col1][8]

            SPACITOT = df_arr[0][self.col1][10]
            SPACIITOT = df_arr[0][self.col1][11]
            SPACIIOLTRETOT = df_arr[0][self.col1][12]
            SPACIIITOT = df_arr[0][self.col1][13]
            SPACIVTOT = df_arr[0][self.col1][14]
            SPACTOT = df_arr[0][self.col1][15]

            SPATTRATEIRISCONTITOT = df_arr[0][self.col1][18]
            SPATTTOT = df_arr[0][self.col1][19]

            SPPA1 = df_arr[0][self.col1][22]
            SPPA2 = df_arr[0][self.col1][23]
            SPPA3 = df_arr[0][self.col1][24]
            SPPA4 = df_arr[0][self.col1][25]
            SPPA5 = df_arr[0][self.col1][26]
            SPPA6 = df_arr[0][self.col1][27]
            SPPA7 = df_arr[0][self.col1][28]
            SPPA8 = df_arr[0][self.col1][29]
            SPPA9 = df_arr[0][self.col1][30]
            SPPATOT = df_arr[0][self.col1][31]

            SPPBTOT = df_arr[0][self.col1][34]

            SPPCTOT = df_arr[0][self.col1][37]

            SPPDTOTENTRO = df_arr[0][self.col1][39]
            SPPDTOTOLTRE = df_arr[0][self.col1][40]
            SPPDTOT = df_arr[0][self.col1][41]

            SPPETOT = df_arr[0][self.col1][44]
            SPPTOT = df_arr[0][self.col1][45]

            self.col2 = df_arr[1].columns[i+2]
            CEA1 = df_arr[1][self.col2][2]
            CEA2 = df_arr[1][self.col2][3]
            CEA3 = df_arr[1][self.col2][4]
            CEA23TOT = df_arr[1][self.col2][5]
            CEA4 = df_arr[1][self.col2][6]
            CEA5 = df_arr[1][self.col2][7]
            CEATOT = df_arr[1][self.col2][8]
            print(df_arr[1][self.col2][8])

            CEB1 = df_arr[1][self.col2][10]
            CEB2 = df_arr[1][self.col2][11]
            CEB3 = df_arr[1][self.col2][12]

            CEB4a = df_arr[1][self.col2][13]
            CEB4b = df_arr[1][self.col2][14]
            CEB4c = df_arr[1][self.col2][15]
            CEB4d = df_arr[1][self.col2][16]
            CEB4e = df_arr[1][self.col2][17]
            CEB4cde = df_arr[1][self.col2][18]

            CEB5a = df_arr[1][self.col2][20]
            CEB5b = df_arr[1][self.col2][21]
            CEB5c = df_arr[1][self.col2][22]
            CEB5abc = df_arr[1][self.col2][23]
            CEB5d = df_arr[1][self.col2][24]
            CEB6 = df_arr[1][self.col2][25]
            CEB7 = df_arr[1][self.col2][26]
            CEB8 = df_arr[1][self.col2][27]
            CEB9 = df_arr[1][self.col2][28]
            CEBTOT = df_arr[1][self.col2][29]
            CEBDIFFVAL = df_arr[1][self.col2][30]

            CEC1 = df_arr[1][self.col2][32]
            CEC2 = df_arr[1][self.col2][33]
            CEC3 = df_arr[1][self.col2][34]
            CEC4 = df_arr[1][self.col2][35]
            CECTOT = df_arr[1][self.col2][36]

            CED1Tot = df_arr[1][self.col2][38]
            CED1a = df_arr[1][self.col2][39]
            CED1b = df_arr[1][self.col2][40]
            CED1c = df_arr[1][self.col2][41]
            CED2Tot = df_arr[1][self.col2][42]
            CED2a = df_arr[1][self.col2][43]
            CED2b = df_arr[1][self.col2][44]
            CED2c = df_arr[1][self.col2][45]
            CEDTot = df_arr[1][self.col2][46]

            CEE1 = df_arr[1][self.col2][48]
            CEE2 = df_arr[1][self.col2][49]
            CEPROVSTRA = df_arr[1][self.col2][50]
            CERISANTEIMP = df_arr[1][self.col2][51]
            CEIMPOSTE = df_arr[1][self.col2][52]
            CEUTILE = df_arr[1][self.col2][53]

            self.col3 = df_arr[2].columns[i+2]
            RFA1 = df_arr[2][self.col3][0]
            RFA2 = df_arr[2][self.col3][1]
            RFA3 = df_arr[2][self.col3][2]
            RFA4 = df_arr[2][self.col3][3]
            RFA5 = df_arr[2][self.col3][4]
            RFA6 = df_arr[2][self.col3][5]
            RFA7 = df_arr[2][self.col3][6]
            RFA8 = df_arr[2][self.col3][7]

            RFB1 = df_arr[2][self.col3][9]
            RFB2 = df_arr[2][self.col3][10]
            RFB3 = df_arr[2][self.col3][11]

            RFB4 = df_arr[2][self.col3][13]

            data[0]['value'] = "P.IVA"
            data[1]['value'] = "value"
            data[2]['value'] = "1"

            data[4]['fixedRows'][0]['values'][0]['value'] = 'n.a.'
            data[4]['fixedRows'][1]['values'][0]['value'] = 'n.a.'
            data[4]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(SPATTLABELACREDSOCITOT)

            data[5]['fixedRows'][0]['values'][0]['value'] = 'n.a.'
            data[5]['fixedRows'][1]['values'][0]['value'] = 'n.a.'
            data[5]['fixedRows'][2]['values'][0]['value'] = 'n.a.'
            data[5]['fixedRows'][3]['values'][0]['value'] = 'n.a.'
            data[5]['fixedRows'][4]['values'][0]['value'] = 'n.a.'
            data[5]['fixedRows'][5]['values'][0]['value'] = 'n.a.'
            data[5]['fixedRows'][6]['values'][0]['value'] = 'n.a.'
            data[5]['fixedRows'][7]['values'][0]['value'] = self.convert_to_str(SPATTLABELBIMMOBIMMATTOT)

            data[6]['fixedRows'][0]['values'][0]['value'] = 'n.a.'
            data[6]['fixedRows'][1]['values'][0]['value'] = 'n.a.'
            data[6]['fixedRows'][2]['values'][0]['value'] = 'n.a.'
            data[6]['fixedRows'][3]['values'][0]['value'] = 'n.a.'
            data[6]['fixedRows'][4]['values'][0]['value'] = 'n.a.'
            data[6]['fixedRows'][5]['values'][0]['value'] = self.convert_to_str(SPATTLABELBIMMOBMATTOT)

            data[7]['fixedRows'][0]['values'][0]['value'] = 'n.a.'
            data[7]['fixedRows'][1]['values'][0]['value'] = 'n.a.'
            data[7]['fixedRows'][2]['values'][0]['value'] = 'n.a.'
            data[7]['fixedRows'][3]['values'][0]['value'] = 'n.a.'
            data[7]['fixedRows'][4]['values'][0]['value'] = 'n.a.'
            data[7]['fixedRows'][5]['values'][0]['value'] = 'n.a.'
            data[7]['fixedRows'][6]['values'][0]['value'] = 'n.a.'
            data[7]['fixedRows'][7]['values'][0]['value'] = 'n.a.'
            data[7]['fixedRows'][8]['values'][0]['value'] = 'n.a.'
            data[7]['fixedRows'][9]['values'][0]['value'] = 'n.a.'
            data[7]['fixedRows'][10]['values'][0]['value'] = self.convert_to_str(SPATTLABELBIMMOBFINANZTOT)
            data[7]['fixedRows'][11]['values'][0]['value'] = self.convert_to_str(SPATTLABELBIMMOBTOT)

            data[8]['fixedRows'][0]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][1]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][2]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][3]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][4]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][5]['values'][0]['value'] = self.convert_to_str(SPACITOT)
            data[8]['fixedRows'][6]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][7]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][8]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][9]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][10]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][11]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][12]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][13]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][14]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][15]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][16]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][17]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][18]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][19]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][20]['values'][0]['value'] = self.convert_to_str(SPACIITOT)
            data[8]['fixedRows'][21]['values'][0]['value'] = self.convert_to_str(SPACIIOLTRETOT)
            data[8]['fixedRows'][22]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][23]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][24]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][25]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][26]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][27]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][28]['values'][0]['value'] = self.convert_to_str(SPACIIITOT)
            data[8]['fixedRows'][29]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][30]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][31]['values'][0]['value'] = 'n.a.'
            data[8]['fixedRows'][32]['values'][0]['value'] = self.convert_to_str(SPACIVTOT)
            data[8]['fixedRows'][33]['values'][0]['value'] = self.convert_to_str(SPACTOT)

            data[9]['fixedRows'][0]['values'][0]['value'] = 'n.a.'
            data[9]['fixedRows'][1]['values'][0]['value'] = 'n.a.'
            data[9]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(SPATTRATEIRISCONTITOT)
            data[9]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(SPATTTOT)

            data[11]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(SPPA1)
            data[11]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(SPPA2)
            data[11]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(SPPA3)
            data[11]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(SPPA4)
            data[11]['fixedRows'][4]['values'][0]['value'] = self.convert_to_str(SPPA5)
            data[11]['fixedRows'][5]['values'][0]['value'] = self.convert_to_str(SPPA6)
            data[11]['fixedRows'][6]['values'][0]['value'] = self.convert_to_str(SPPA7)
            data[11]['fixedRows'][7]['values'][0]['value'] = self.convert_to_str(SPPA8)
            data[11]['fixedRows'][8]['values'][0]['value'] = self.convert_to_str(SPPA9)
            data[11]['fixedRows'][9]['values'][0]['value'] = self.convert_to_str(SPPATOT)

            data[12]['fixedRows'][0]['values'][0]['value'] = 'n.a.'
            data[12]['fixedRows'][1]['values'][0]['value'] = 'n.a.'
            data[12]['fixedRows'][2]['values'][0]['value'] = 'n.a.'
            data[12]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(SPPBTOT)

            data[13]['fixedRows'][0]['values'][0]['value'] = 'n.a.'
            data[13]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(SPPCTOT)

            data[14]['fixedRows'][0]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][1]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][2]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][3]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][4]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][5]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][6]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][7]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][8]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][9]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][10]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][11]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][12]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][13]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][14]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][15]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][16]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][17]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][18]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][19]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][20]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][21]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][22]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][23]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][24]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][25]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][26]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][27]['values'][0]['value'] = 'n.a.'
            data[14]['fixedRows'][28]['values'][0]['value'] = self.convert_to_str(SPPDTOTENTRO)
            data[14]['fixedRows'][29]['values'][0]['value'] = self.convert_to_str(SPPDTOTOLTRE)
            data[14]['fixedRows'][30]['values'][0]['value'] = self.convert_to_str(SPPDTOT)

            data[15]['fixedRows'][0]['values'][0]['value'] = 'n.a.'
            data[15]['fixedRows'][1]['values'][0]['value'] = 'n.a.'
            data[15]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(SPPETOT)
            data[15]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(SPPTOT)

            data[17]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(CEA1)
            data[17]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(CEA2)
            data[17]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(CEA3)
            data[17]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(CEA23TOT)
            data[17]['fixedRows'][4]['values'][0]['value'] = self.convert_to_str(CEA4)
            data[17]['fixedRows'][5]['values'][0]['value'] = self.convert_to_str(CEA5)
            data[17]['fixedRows'][6]['values'][0]['value'] = self.convert_to_str(CEATOT)

            data[18]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(CEB1)
            data[18]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(CEB2)
            data[18]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(CEB3)
            data[18]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(CEB4a)
            data[18]['fixedRows'][4]['values'][0]['value'] = self.convert_to_str(CEB4b)
            data[18]['fixedRows'][5]['values'][0]['value'] = self.convert_to_str(CEB4c)
            data[18]['fixedRows'][6]['values'][0]['value'] = self.convert_to_str(CEB4d)
            data[18]['fixedRows'][7]['values'][0]['value'] = self.convert_to_str(CEB4e)
            data[18]['fixedRows'][8]['values'][0]['value'] = self.convert_to_str(CEB4cde)
            data[18]['fixedRows'][9]['values'][0]['value'] = self.convert_to_str(CEB5a)
            data[18]['fixedRows'][10]['values'][0]['value'] = self.convert_to_str(CEB5b)
            data[18]['fixedRows'][11]['values'][0]['value'] = self.convert_to_str(CEB5c)
            data[18]['fixedRows'][12]['values'][0]['value'] = self.convert_to_str(CEB5abc)
            data[18]['fixedRows'][13]['values'][0]['value'] = self.convert_to_str(CEB5d)
            data[18]['fixedRows'][14]['values'][0]['value'] = self.convert_to_str(CEB6)
            data[18]['fixedRows'][15]['values'][0]['value'] = self.convert_to_str(CEB7)
            data[18]['fixedRows'][16]['values'][0]['value'] = self.convert_to_str(CEB8)
            data[18]['fixedRows'][17]['values'][0]['value'] = self.convert_to_str(CEB9)
            data[18]['fixedRows'][18]['values'][0]['value'] = self.convert_to_str(CEBTOT)
            data[18]['fixedRows'][19]['values'][0]['value'] = self.convert_to_str(CEBDIFFVAL)

            data[19]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(CEC1)
            data[19]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(CEC2)
            data[19]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(CEC3)
            data[19]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(CEC4)
            data[19]['fixedRows'][4]['values'][0]['value'] = self.convert_to_str(CECTOT)

            data[20]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(CED1a)
            data[20]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(CED1b)
            data[20]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(CED1c)
            data[20]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(CED1Tot)
            data[20]['fixedRows'][4]['values'][0]['value'] = self.convert_to_str(CED2a)
            data[20]['fixedRows'][5]['values'][0]['value'] = self.convert_to_str(CED2b)
            data[20]['fixedRows'][6]['values'][0]['value'] = self.convert_to_str(CED2c)
            data[20]['fixedRows'][7]['values'][0]['value'] = self.convert_to_str(CED2Tot)
            data[20]['fixedRows'][8]['values'][0]['value'] = self.convert_to_str(CEDTot)

            data[21]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(CEE1)
            data[21]['fixedRows'][1]['values'][0]['value'] = 'n.a.'
            data[21]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(CEE2)
            data[21]['fixedRows'][3]['values'][0]['value'] = 'n.a.'
            data[21]['fixedRows'][4]['values'][0]['value'] = self.convert_to_str(CEPROVSTRA)

            data[22]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(CERISANTEIMP)
            data[22]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(CEIMPOSTE)
            data[22]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(CEUTILE)

            data[24]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(RFA1)
            data[24]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(RFA2)
            data[24]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(RFA3)
            data[24]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(RFA4)
            data[24]['fixedRows'][4]['values'][0]['value'] = self.convert_to_str(RFA5)
            data[24]['fixedRows'][5]['values'][0]['value'] = self.convert_to_str(RFA6)
            data[24]['fixedRows'][6]['values'][0]['value'] = self.convert_to_str(RFA7)
            data[24]['fixedRows'][7]['values'][0]['value'] = self.convert_to_str(RFA8)

            data[25]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(RFB1)
            data[25]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(RFB2)
            data[25]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(RFB3)
            data[25]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(RFB4)
        
            self.json_file = data
            self.save_json()

    def data_read_dettagliato(self, df_arr):
        """
        reads data and save it to self.json_file if
        the type of survey is Dettagliato
        :param df_arr: array that contains 3 sheets of survey
        :return: sets the new values in self.json_file and saves the json_file
        """
        data = self.json_file
        # for all the columns that contains year's data iterate
        # columns of years start from the 3rd column
        for i in range(len(df_arr[0].columns[2:])):
            self.col1 = df_arr[0].columns[i+2]
            SPAA1 = df_arr[0][self.col1][2]
            SPAA2 = df_arr[0][self.col1][3]
            SPATTLABELACREDSOCITOT = df_arr[0][self.col1][4]

            SPABI1 = df_arr[0][self.col1][7]
            SPABI2 = df_arr[0][self.col1][8]

            SPABI3 = df_arr[0][self.col1][9]
            SPABI4 = df_arr[0][self.col1][10]
            SPABI5 = df_arr[0][self.col1][11]
            SPABI6 = df_arr[0][self.col1][12]
            SPABI7 = df_arr[0][self.col1][13]
            SPATTLABELBIMMOBIMMATTOT = df_arr[0][self.col1][14]

            SPABII1 = df_arr[0][self.col1][16]
            SPABII2 = df_arr[0][self.col1][17]
            SPABII3 = df_arr[0][self.col1][18]
            SPABII4 = df_arr[0][self.col1][19]
            SPABII5 = df_arr[0][self.col1][20]
            SPATTLABELBIMMOBMATTOT = df_arr[0][self.col1][21]

            SPABIII1A = df_arr[0][self.col1][24]
            SPABIII1B = df_arr[0][self.col1][25]
            SPABIII1C = df_arr[0][self.col1][26]
            SPABIII1D = df_arr[0][self.col1][27]

            SPABIII2A = df_arr[0][self.col1][29]
            SPABIII2B = df_arr[0][self.col1][30]
            SPABIII2C = df_arr[0][self.col1][31]
            SPABIII2D = df_arr[0][self.col1][32]
            SPABIII3 = df_arr[0][self.col1][33]
            SPABIII4 = df_arr[0][self.col1][34]
            SPATTLABELBIMMOBFINANZTOT = df_arr[0][self.col1][35]
            SPATTLABELBIMMOBTOT = df_arr[0][self.col1][36]

            SPACI1 = df_arr[0][self.col1][39]
            SPACI2 = df_arr[0][self.col1][40]
            SPACI3 = df_arr[0][self.col1][41]
            SPACI4 = df_arr[0][self.col1][42]
            SPACI5 = df_arr[0][self.col1][43]
            SPACITOT = df_arr[0][self.col1][44]

            SPACII1 = df_arr[0][self.col1][46]
            SPACII1OLTRE = df_arr[0][self.col1][47]
            SPACII2 = df_arr[0][self.col1][48]
            SPACII2OLTRE = df_arr[0][self.col1][49]
            SPACII3 = df_arr[0][self.col1][50]
            SPACII3OLTRE = df_arr[0][self.col1][51]
            SPACII4 = df_arr[0][self.col1][52]
            SPACII4OLTRE = df_arr[0][self.col1][53]
            SPACII5 = df_arr[0][self.col1][54]
            SPACII5OLTRE = df_arr[0][self.col1][55]
            SPACII6 = df_arr[0][self.col1][56]
            SPACII6OLTRE = df_arr[0][self.col1][57]
            SPACII7 = df_arr[0][self.col1][58]
            SPACII7OLTRE = df_arr[0][self.col1][59]
            SPACIITOT = df_arr[0][self.col1][60]
            SPACIIOLTRETOT = df_arr[0][self.col1][61]

            SPACIII1 = df_arr[0][self.col1][63]
            SPACIII2 = df_arr[0][self.col1][64]
            SPACIII3 = df_arr[0][self.col1][65]
            SPACIII4 = df_arr[0][self.col1][66]
            SPACIII5 = df_arr[0][self.col1][67]
            SPACIII6 = df_arr[0][self.col1][68]
            SPACIIITOT = df_arr[0][self.col1][69]

            SPACIV1 = df_arr[0][self.col1][71]
            SPACIV2 = df_arr[0][self.col1][72]
            SPACIV3 = df_arr[0][self.col1][73]
            SPACIVTOT = df_arr[0][self.col1][74]
            SPACTOT = df_arr[0][self.col1][75]

            SPAD1 = df_arr[0][self.col1][77]
            SPAD2 = df_arr[0][self.col1][78]
            SPATTRATEIRISCONTITOT = df_arr[0][self.col1][79]
            SPATTTOT = df_arr[0][self.col1][80]

            SPPA1 = df_arr[0][self.col1][83]
            SPPA2 = df_arr[0][self.col1][84]
            SPPA3 = df_arr[0][self.col1][85]
            SPPA4 = df_arr[0][self.col1][86]
            SPPA5 = df_arr[0][self.col1][87]
            SPPA6 = df_arr[0][self.col1][88]
            SPPA7 = df_arr[0][self.col1][89]
            SPPA8 = df_arr[0][self.col1][90]
            SPPA9 = df_arr[0][self.col1][91]
            SPPATOT = df_arr[0][self.col1][92]

            SPPB1 = df_arr[0][self.col1][94]
            SPPB2 = df_arr[0][self.col1][95]
            SPPB3 = df_arr[0][self.col1][96]
            SPPBTOT = df_arr[0][self.col1][97]

            SPPC1 = df_arr[0][self.col1][99]
            SPPCTOT = df_arr[0][self.col1][100]

            SPPD1 = df_arr[0][self.col1][102]
            SPPD1OLTRE = df_arr[0][self.col1][103]
            SPPD2 = df_arr[0][self.col1][104]
            SPPD2OLTRE = df_arr[0][self.col1][105]
            SPPD3 = df_arr[0][self.col1][106]
            SPPD3OLTRE = df_arr[0][self.col1][107]
            SPPD4 = df_arr[0][self.col1][108]
            SPPD4OLTRE = df_arr[0][self.col1][109]
            SPPD5 = df_arr[0][self.col1][110]
            SPPD5OLTRE = df_arr[0][self.col1][111]
            SPPD6 = df_arr[0][self.col1][112]
            SPPD6OLTRE = df_arr[0][self.col1][113]
            SPPD7 = df_arr[0][self.col1][114]
            SPPD7OLTRE = df_arr[0][self.col1][115]
            SPPD8 = df_arr[0][self.col1][116]
            SPPD8OLTRE = df_arr[0][self.col1][117]
            SPPD9 = df_arr[0][self.col1][118]
            SPPD9OLTRE = df_arr[0][self.col1][119]
            SPPD10 = df_arr[0][self.col1][120]
            SPPD10OLTRE = df_arr[0][self.col1][121]
            SPPD11 = df_arr[0][self.col1][122]
            SPPD11OLTRE = df_arr[0][self.col1][123]
            SPPD12 = df_arr[0][self.col1][124]
            SPPD12OLTRE = df_arr[0][self.col1][125]
            SPPD13 = df_arr[0][self.col1][126]
            SPPD13OLTRE = df_arr[0][self.col1][127]
            SPPD14 = df_arr[0][self.col1][128]
            SPPD14OLTRE = df_arr[0][self.col1][129]
            SPPDTOTENTRO = df_arr[0][self.col1][130]
            SPPDTOTOLTRE = df_arr[0][self.col1][131]
            SPPDTOT = df_arr[0][self.col1][132]

            SPPE1 = df_arr[0][self.col1][134]
            SPPE2 = df_arr[0][self.col1][135]
            SPPETOT = df_arr[0][self.col1][136]
            SPPTOT = df_arr[0][self.col1][137]
            self.col2 = df_arr[1].columns[i+2]
            CEA1 = df_arr[1][self.col2][1]
            CEA2 = df_arr[1][self.col2][2]
            CEA3 = df_arr[1][self.col2][3]
            CEA23TOT = df_arr[1][self.col2][4]
            CEA4 = df_arr[1][self.col2][5]
            CEA5 = df_arr[1][self.col2][6]
            CEATOT = df_arr[1][self.col2][7]

            CEB1 = df_arr[1][self.col2][9]
            CEB2 = df_arr[1][self.col2][10]
            CEB3 = df_arr[1][self.col2][11]

            CEB4a = df_arr[1][self.col2][13]
            CEB4b = df_arr[1][self.col2][14]
            CEB4c = df_arr[1][self.col2][15]
            CEB4d = df_arr[1][self.col2][16]
            CEB4e = df_arr[1][self.col2][17]
            CEB4cde = df_arr[1][self.col2][18]

            CEB5a = df_arr[1][self.col2][20]
            CEB5b = df_arr[1][self.col2][21]
            CEB5c = df_arr[1][self.col2][22]
            CEB5abc = df_arr[1][self.col2][23]
            CEB5d = df_arr[1][self.col2][24]
            CEB6 = df_arr[1][self.col2][25]
            CEB7 = df_arr[1][self.col2][26]
            CEB8 = df_arr[1][self.col2][27]
            CEB9 = df_arr[1][self.col2][28]
            CEBTOT = df_arr[1][self.col2][29]
            CEBDIFFVAL = df_arr[1][self.col2][30]

            CEC1 = df_arr[1][self.col2][32]
            CEC2 = df_arr[1][self.col2][33]
            CEC3 = df_arr[1][self.col2][34]
            CEC4 = df_arr[1][self.col2][35]
            CECTOT = df_arr[1][self.col2][36]

            CED1Tot = df_arr[1][self.col2][38]
            CED1a = df_arr[1][self.col2][39]
            CED1b = df_arr[1][self.col2][40]
            CED1c = df_arr[1][self.col2][41]
            CED2Tot = df_arr[1][self.col2][42]
            CED2a = df_arr[1][self.col2][43]
            CED2b = df_arr[1][self.col2][44]
            CED2c = df_arr[1][self.col2][45]
            CEDTot = df_arr[1][self.col2][46]

            CEE1 = df_arr[1][self.col2][48]
            CEE1PLUS = df_arr[1][self.col2][49]
            CEE2 = df_arr[1][self.col2][50]
            CEE2MINUS = df_arr[1][self.col2][51]
            CEPROVSTRA = df_arr[1][self.col2][52]
            CERISANTEIMP = df_arr[1][self.col2][53]
            CEIMPOSTE = df_arr[1][self.col2][54]
            CEUTILE = df_arr[1][self.col2][55]

            self.col3 = df_arr[2].columns[i+2]
            RFA1 = df_arr[2][self.col3][0]
            RFA2 = df_arr[2][self.col3][1]
            RFA3 = df_arr[2][self.col3][2]
            RFA4 = df_arr[2][self.col3][3]
            RFA5 = df_arr[2][self.col3][4]
            RFA6 = df_arr[2][self.col3][5]
            RFA7 = df_arr[2][self.col3][6]
            RFA8 = df_arr[2][self.col3][7]

            RFB1 = df_arr[2][self.col3][9]
            RFB2 = df_arr[2][self.col3][10]
            RFB3 = df_arr[2][self.col3][11]

            RFB4 = df_arr[2][self.col3][13]

            data[0]['value'] = "P.IVA"
            data[1]['value'] = "value"
            data[2]['value'] = "3"

            data[4]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(SPAA1)
            data[4]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(SPAA2)
            data[4]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(SPATTLABELACREDSOCITOT)

            data[5]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(SPABI1)
            data[5]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(SPABI2)
            data[5]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(SPABI3)
            data[5]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(SPABI4)
            data[5]['fixedRows'][4]['values'][0]['value'] = self.convert_to_str(SPABI5)
            data[5]['fixedRows'][5]['values'][0]['value'] = self.convert_to_str(SPABI6)
            data[5]['fixedRows'][6]['values'][0]['value'] = self.convert_to_str(SPABI7)
            data[5]['fixedRows'][7]['values'][0]['value'] = self.convert_to_str(SPATTLABELBIMMOBIMMATTOT)

            data[6]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(SPABII1)
            data[6]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(SPABII2)
            data[6]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(SPABII3)
            data[6]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(SPABII4)
            data[6]['fixedRows'][4]['values'][0]['value'] = self.convert_to_str(SPABII5)
            data[6]['fixedRows'][5]['values'][0]['value'] = self.convert_to_str(SPATTLABELBIMMOBMATTOT)

            data[7]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(SPABIII1A)
            data[7]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(SPABIII1B)
            data[7]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(SPABIII1C)
            data[7]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(SPABIII1D)
            data[7]['fixedRows'][4]['values'][0]['value'] = self.convert_to_str(SPABIII2A)
            data[7]['fixedRows'][5]['values'][0]['value'] = self.convert_to_str(SPABIII2B)
            data[7]['fixedRows'][6]['values'][0]['value'] = self.convert_to_str(SPABIII2C)
            data[7]['fixedRows'][7]['values'][0]['value'] = self.convert_to_str(SPABIII2D)
            data[7]['fixedRows'][8]['values'][0]['value'] = self.convert_to_str(SPABIII3)
            data[7]['fixedRows'][9]['values'][0]['value'] = self.convert_to_str(SPABIII4)
            data[7]['fixedRows'][10]['values'][0]['value'] = self.convert_to_str(SPATTLABELBIMMOBFINANZTOT)
            data[7]['fixedRows'][11]['values'][0]['value'] = self.convert_to_str(SPATTLABELBIMMOBTOT)

            data[8]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(SPACI1)
            data[8]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(SPACI2)
            data[8]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(SPACI3)
            data[8]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(SPACI4)
            data[8]['fixedRows'][4]['values'][0]['value'] = self.convert_to_str(SPACI5)
            data[8]['fixedRows'][5]['values'][0]['value'] = self.convert_to_str(SPACITOT)
            data[8]['fixedRows'][6]['values'][0]['value'] = self.convert_to_str(SPACII1)
            data[8]['fixedRows'][7]['values'][0]['value'] = self.convert_to_str(SPACII1OLTRE)
            data[8]['fixedRows'][8]['values'][0]['value'] = self.convert_to_str(SPACII2)
            data[8]['fixedRows'][9]['values'][0]['value'] = self.convert_to_str(SPACII2OLTRE)
            data[8]['fixedRows'][10]['values'][0]['value'] = self.convert_to_str(SPACII3)
            data[8]['fixedRows'][11]['values'][0]['value'] = self.convert_to_str(SPACII3OLTRE)
            data[8]['fixedRows'][12]['values'][0]['value'] = self.convert_to_str(SPACII4)
            data[8]['fixedRows'][13]['values'][0]['value'] = self.convert_to_str(SPACII4OLTRE)
            data[8]['fixedRows'][14]['values'][0]['value'] = self.convert_to_str(SPACII5)
            data[8]['fixedRows'][15]['values'][0]['value'] = self.convert_to_str(SPACII5OLTRE)
            data[8]['fixedRows'][16]['values'][0]['value'] = self.convert_to_str(SPACII6)
            data[8]['fixedRows'][17]['values'][0]['value'] = self.convert_to_str(SPACII6OLTRE)
            data[8]['fixedRows'][18]['values'][0]['value'] = self.convert_to_str(SPACII7)
            data[8]['fixedRows'][19]['values'][0]['value'] = self.convert_to_str(SPACII7OLTRE)
            data[8]['fixedRows'][20]['values'][0]['value'] = self.convert_to_str(SPACIITOT)
            data[8]['fixedRows'][21]['values'][0]['value'] = self.convert_to_str(SPACIIOLTRETOT)
            data[8]['fixedRows'][22]['values'][0]['value'] = self.convert_to_str(SPACIII1)
            data[8]['fixedRows'][23]['values'][0]['value'] = self.convert_to_str(SPACIII2)
            data[8]['fixedRows'][24]['values'][0]['value'] = self.convert_to_str(SPACIII3)
            data[8]['fixedRows'][25]['values'][0]['value'] = self.convert_to_str(SPACIII4)
            data[8]['fixedRows'][26]['values'][0]['value'] = self.convert_to_str(SPACIII5)
            data[8]['fixedRows'][27]['values'][0]['value'] = self.convert_to_str(SPACIII6)
            data[8]['fixedRows'][28]['values'][0]['value'] = self.convert_to_str(SPACIIITOT)
            data[8]['fixedRows'][29]['values'][0]['value'] = self.convert_to_str(SPACIV1)
            data[8]['fixedRows'][30]['values'][0]['value'] = self.convert_to_str(SPACIV2)
            data[8]['fixedRows'][31]['values'][0]['value'] = self.convert_to_str(SPACIV3)
            data[8]['fixedRows'][32]['values'][0]['value'] = self.convert_to_str(SPACIVTOT)
            data[8]['fixedRows'][33]['values'][0]['value'] = self.convert_to_str(SPACTOT)

            data[9]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(SPAD1)
            data[9]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(SPAD2)
            data[9]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(SPATTRATEIRISCONTITOT)
            data[9]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(SPATTTOT)

            data[11]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(SPPA1)
            data[11]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(SPPA2)
            data[11]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(SPPA3)
            data[11]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(SPPA4)
            data[11]['fixedRows'][4]['values'][0]['value'] = self.convert_to_str(SPPA5)
            data[11]['fixedRows'][5]['values'][0]['value'] = self.convert_to_str(SPPA6)
            data[11]['fixedRows'][6]['values'][0]['value'] = self.convert_to_str(SPPA7)
            data[11]['fixedRows'][7]['values'][0]['value'] = self.convert_to_str(SPPA8)
            data[11]['fixedRows'][8]['values'][0]['value'] = self.convert_to_str(SPPA9)
            data[11]['fixedRows'][9]['values'][0]['value'] = self.convert_to_str(SPPATOT)

            data[12]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(SPPB1)
            data[12]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(SPPB2)
            data[12]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(SPPB3)
            data[12]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(SPPBTOT)

            data[13]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(SPPC1)
            data[13]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(SPPCTOT)

            data[14]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(SPPD1)
            data[14]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(SPPD1OLTRE)
            data[14]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(SPPD2)
            data[14]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(SPPD2OLTRE)
            data[14]['fixedRows'][4]['values'][0]['value'] = self.convert_to_str(SPPD3)
            data[14]['fixedRows'][5]['values'][0]['value'] = self.convert_to_str(SPPD3OLTRE)
            data[14]['fixedRows'][6]['values'][0]['value'] = self.convert_to_str(SPPD4)
            data[14]['fixedRows'][7]['values'][0]['value'] = self.convert_to_str(SPPD4OLTRE)
            data[14]['fixedRows'][8]['values'][0]['value'] = self.convert_to_str(SPPD5)
            data[14]['fixedRows'][9]['values'][0]['value'] = self.convert_to_str(SPPD5OLTRE)
            data[14]['fixedRows'][10]['values'][0]['value'] = self.convert_to_str(SPPD6)
            data[14]['fixedRows'][11]['values'][0]['value'] = self.convert_to_str(SPPD6OLTRE)
            data[14]['fixedRows'][12]['values'][0]['value'] = self.convert_to_str(SPPD7)
            data[14]['fixedRows'][13]['values'][0]['value'] = self.convert_to_str(SPPD7OLTRE)
            data[14]['fixedRows'][14]['values'][0]['value'] = self.convert_to_str(SPPD8)
            data[14]['fixedRows'][15]['values'][0]['value'] = self.convert_to_str(SPPD8OLTRE)
            data[14]['fixedRows'][16]['values'][0]['value'] = self.convert_to_str(SPPD9)
            data[14]['fixedRows'][17]['values'][0]['value'] = self.convert_to_str(SPPD9OLTRE)
            data[14]['fixedRows'][18]['values'][0]['value'] = self.convert_to_str(SPPD10)
            data[14]['fixedRows'][19]['values'][0]['value'] = self.convert_to_str(SPPD10OLTRE)
            data[14]['fixedRows'][20]['values'][0]['value'] = self.convert_to_str(SPPD11)
            data[14]['fixedRows'][21]['values'][0]['value'] = self.convert_to_str(SPPD11OLTRE)
            data[14]['fixedRows'][22]['values'][0]['value'] = self.convert_to_str(SPPD12)
            data[14]['fixedRows'][23]['values'][0]['value'] = self.convert_to_str(SPPD12OLTRE)
            data[14]['fixedRows'][24]['values'][0]['value'] = self.convert_to_str(SPPD13)
            data[14]['fixedRows'][25]['values'][0]['value'] = self.convert_to_str(SPPD13OLTRE)
            data[14]['fixedRows'][26]['values'][0]['value'] = self.convert_to_str(SPPD14)
            data[14]['fixedRows'][27]['values'][0]['value'] = self.convert_to_str(SPPD14OLTRE)
            data[14]['fixedRows'][28]['values'][0]['value'] = self.convert_to_str(SPPDTOTENTRO)
            data[14]['fixedRows'][29]['values'][0]['value'] = self.convert_to_str(SPPDTOTOLTRE)
            data[14]['fixedRows'][30]['values'][0]['value'] = self.convert_to_str(SPPDTOT)

            data[15]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(SPPE1)
            data[15]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(SPPE2)
            data[15]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(SPPETOT)
            data[15]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(SPPTOT)

            data[17]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(CEA1)
            data[17]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(CEA2)
            data[17]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(CEA3)
            data[17]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(CEA23TOT)
            data[17]['fixedRows'][4]['values'][0]['value'] = self.convert_to_str(CEA4)
            data[17]['fixedRows'][5]['values'][0]['value'] = self.convert_to_str(CEA5)
            data[17]['fixedRows'][6]['values'][0]['value'] = self.convert_to_str(CEATOT)

            data[18]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(CEB1)
            data[18]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(CEB2)
            data[18]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(CEB3)
            data[18]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(CEB4a)
            data[18]['fixedRows'][4]['values'][0]['value'] = self.convert_to_str(CEB4b)
            data[18]['fixedRows'][5]['values'][0]['value'] = self.convert_to_str(CEB4c)
            data[18]['fixedRows'][6]['values'][0]['value'] = self.convert_to_str(CEB4d)
            data[18]['fixedRows'][7]['values'][0]['value'] = self.convert_to_str(CEB4e)
            data[18]['fixedRows'][8]['values'][0]['value'] = self.convert_to_str(CEB4cde)
            data[18]['fixedRows'][9]['values'][0]['value'] = self.convert_to_str(CEB5a)
            data[18]['fixedRows'][10]['values'][0]['value'] = self.convert_to_str(CEB5b)
            data[18]['fixedRows'][11]['values'][0]['value'] = self.convert_to_str(CEB5c)
            data[18]['fixedRows'][12]['values'][0]['value'] = self.convert_to_str(CEB5abc)
            data[18]['fixedRows'][13]['values'][0]['value'] = self.convert_to_str(CEB5d)
            data[18]['fixedRows'][14]['values'][0]['value'] = self.convert_to_str(CEB6)
            data[18]['fixedRows'][15]['values'][0]['value'] = self.convert_to_str(CEB7)
            data[18]['fixedRows'][16]['values'][0]['value'] = self.convert_to_str(CEB8)
            data[18]['fixedRows'][17]['values'][0]['value'] = self.convert_to_str(CEB9)
            data[18]['fixedRows'][18]['values'][0]['value'] = self.convert_to_str(CEBTOT)
            data[18]['fixedRows'][19]['values'][0]['value'] = self.convert_to_str(CEBDIFFVAL)

            data[19]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(CEC1)
            data[19]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(CEC2)
            data[19]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(CEC3)
            data[19]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(CEC4)
            data[19]['fixedRows'][4]['values'][0]['value'] = self.convert_to_str(CECTOT)

            data[20]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(CED1a)
            data[20]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(CED1b)
            data[20]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(CED1c)
            data[20]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(CED1Tot)
            data[20]['fixedRows'][4]['values'][0]['value'] = self.convert_to_str(CED2a)
            data[20]['fixedRows'][5]['values'][0]['value'] = self.convert_to_str(CED2b)
            data[20]['fixedRows'][6]['values'][0]['value'] = self.convert_to_str(CED2c)
            data[20]['fixedRows'][7]['values'][0]['value'] = self.convert_to_str(CED2Tot)
            data[20]['fixedRows'][8]['values'][0]['value'] = self.convert_to_str(CEDTot)

            data[21]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(CEE1)
            data[21]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(CEE1PLUS)
            data[21]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(CEE2)
            data[21]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(CEE2MINUS)
            data[21]['fixedRows'][4]['values'][0]['value'] = self.convert_to_str(CEPROVSTRA)

            data[22]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(CERISANTEIMP)
            data[22]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(CEIMPOSTE)
            data[22]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(CEUTILE)

            data[24]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(RFA1)
            data[24]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(RFA2)
            data[24]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(RFA3)
            data[24]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(RFA4)
            data[24]['fixedRows'][4]['values'][0]['value'] = self.convert_to_str(RFA5)
            data[24]['fixedRows'][5]['values'][0]['value'] = self.convert_to_str(RFA6)
            data[24]['fixedRows'][6]['values'][0]['value'] = self.convert_to_str(RFA7)
            data[24]['fixedRows'][7]['values'][0]['value'] = self.convert_to_str(RFA8)

            data[25]['fixedRows'][0]['values'][0]['value'] = self.convert_to_str(RFB1)
            data[25]['fixedRows'][1]['values'][0]['value'] = self.convert_to_str(RFB2)
            data[25]['fixedRows'][2]['values'][0]['value'] = self.convert_to_str(RFB3)
            data[25]['fixedRows'][3]['values'][0]['value'] = self.convert_to_str(RFB4)

            self.json_file = data
            self.save_json()
