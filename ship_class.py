import openpyxl

class ship:
    def __init__(self,current_shipping_lane,vessel_code,three_digit_code, index, read_worksheet, write_worksheet, write_workbook, report_path):
        self.current_shipping_lane = current_shipping_lane
        self.vessel_code = vessel_code
        self.three_digit_code = three_digit_code
        self.data_code = "" 
        self.start_index = 2
        self.read_worksheet = read_worksheet
        self.write_worksheet = write_worksheet
        self.write_workbook = write_workbook
        self.report_path = report_path
        self.shipping_data = {
            ## Put HAL in VAN
            "VAN": {
                "PRR": 0,
                "VAN": 0, 
                "RF": 0
            }, 
            "MTR": {
                "PRR": 0,
                "VAN": 0, 
                "RF": 0
            }, 
            "TOR": {
                "PRR": 0,
                "VAN": 0, 
                "RF": 0
            },
            "EBI": {
                "PRR": 0,
                "VAN": 0, 
                "RF": 0
            } 
        }
        self.index = index

    def create_name(self): 
        str = self.current_shipping_lane + "-" + self.vessel_code + '-' + self.three_digit_code
        self.data_code = str

    def get_data(self):
        for i in range(self.start_index, self.read_worksheet.max_row):
            svvd = self.read_worksheet['D'+str(i)].value.split(" ") 

            if svvd[0] == self.data_code:
                booking_office_code = self.read_worksheet['A'+str(i)].value
                pol = self.read_worksheet['C' + str(i)].value
                emedia = self.read_worksheet['F' + str(i)].value

                bkg_teg = self.read_worksheet['G' + str(i)].value
                bkg_teg = int(bkg_teg) 

                rf = self.read_worksheet['L' + str(i)].value
                rf = int(rf)


                if (booking_office_code == "EBI"):
                    if (pol == "MTR" or pol == "HAL"):
                        self.shipping_data["EBI"]["VAN"] += bkg_teg
                        self.shipping_data["EBI"]["RF"] += rf
                    else:
                        self.shipping_data["EBI"][pol] += bkg_teg
                        self.shipping_data["EBI"]["RF"] += rf

                else:
                    if (emedia == 'SynconHub - Contract'):
                        if (pol == "MTR" or pol == "HAL"):
                            self.shipping_data["EBI"]["VAN"] += bkg_teg
                            self.shipping_data["EBI"]["RF"] += rf
                        else:                            
                            self.shipping_data["EBI"][pol] += bkg_teg
                            self.shipping_data["EBI"]["RF"] += rf
                    else: 
                        if (pol == "MTR" or pol == "HAL"):
                            self.shipping_data[booking_office_code]["VAN"] += bkg_teg
                            self.shipping_data[booking_office_code]["RF"] += rf
                        else:
                            self.shipping_data[booking_office_code][pol] += bkg_teg
                            self.shipping_data[booking_office_code]["RF"] += rf
        
        #print(str(self.data_code) + " " + str(self.shipping_data) + "\n")
                        
    def paste_data(self):
        letters = ['L', 'N', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y']


        #Vancouver Office
        self.write_worksheet['L' + str(self.index)].value = self.shipping_data["VAN"]["PRR"]
        self.write_worksheet['N' + str(self.index)].value = self.shipping_data["VAN"]["VAN"]
        self.write_worksheet['P' + str(self.index)].value = self.shipping_data["VAN"]["RF"]

        #Montreal Office
        self.write_worksheet['Q' + str(self.index)].value = self.shipping_data["MTR"]["PRR"]
        self.write_worksheet['R' + str(self.index)].value = self.shipping_data["MTR"]["VAN"]
        self.write_worksheet['S' + str(self.index)].value = self.shipping_data["MTR"]["RF"]

        #Toronto Office
        self.write_worksheet['T' + str(self.index)].value = self.shipping_data["TOR"]["PRR"]
        self.write_worksheet['U' + str(self.index)].value = self.shipping_data["TOR"]["VAN"]
        self.write_worksheet['V' + str(self.index)].value = self.shipping_data["TOR"]["RF"]

        #EBI 
        self.write_worksheet['W' + str(self.index)].value = self.shipping_data["EBI"]["PRR"]
        self.write_worksheet['X' + str(self.index)].value = self.shipping_data["EBI"]["VAN"]
        self.write_worksheet['Y' + str(self.index)].value = self.shipping_data["EBI"]["RF"]

        
        for letter in letters:
            if self.write_worksheet[letter + str(self.index)].fill.fgColor.tint < 0:
                self.write_worksheet[letter + str(self.index)].value = "" 
        
    def save_data(self): 
        self.write_workbook.save(self.report_path) 




                
                    
                    
                    


                

    

    

    
        

