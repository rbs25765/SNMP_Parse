import re
import os
import xlsxwriter
from collections import OrderedDict
import logging


class SnmpCheck:

    def __init__(self):
        self.host_pat = re.compile(r'(?:hostname\s)(.+)')
        self.acl_pat = re.compile(r'(?:access-list\s)(?:11\s)(?:permit\s)(.+)')
        self.commu_pat = re.compile(r'(?:snmp-server.+)(Btcpe2niab#\s)(?:RO 11)')
        self.int_desc = re.compile(r'(?: desc.+)(BTCO(\s+|\S+))')
        self.master_dict = {}
        self.file_list = []
        logging.basicConfig(level=logging.INFO)

    def site_data(self, file_input):
        file = open(file_input)
        device_list = []
        device_dict = OrderedDict()
        commu_flag = False
        desc_flag = False
        host_flag = False
        logging.info("Reading file to extract snmp info"+"."*30)
        for line in file:
            if self.host_pat.match(line):
                device_dict['Hostname'] = self.host_pat.match(line).group(1)
                host_flag = True

            elif not host_flag and not self.host_pat.match(line):
                device_dict['Hostname'] = "No Running Config"

            if self.commu_pat.match(line):
                device_dict['Community'] = self.commu_pat.match(line).group(1)
                commu_flag = True

            elif not commu_flag and not self.commu_pat.match(line):
                device_dict['Community'] = "Not Configured"

            if self.acl_pat.match(line):
                device_list.append(self.acl_pat.match(line).group(1))

            if self.int_desc.match(line):
                device_dict['Description'] = self.int_desc.match(line).group(1)
                desc_flag = True
            elif not desc_flag and not self.int_desc.match(line):
                device_dict['Description'] = "Desc Not Configured "

        if not device_list == []:
            if device_list[0] and device_list[1]:
                device_dict['Access_List'] = "ACL Configured"
        else:
            device_dict['Access_List'] = "ACL Not Configured"
        return device_dict

    def file_list_extract (self):
        for file in os.listdir('./Input'):
            self.file_list.append(os.path.join(os.path.abspath('./Input'), file))
        return self.file_list

    def master_dict_extract(self, files_list):

        master_dict = {}
        for file_item in files_list:
            logging.info("Opening file {}".format(os.path.basename(file_item))+"."*35)
            master_dict[os.path.basename(file_item)] = self.site_data(file_item)
        logging.info("Data Extract completed for {} files".format(len(files_list)))
        return master_dict

    @staticmethod
    def excel_writer(dicto):

        workbook = xlsxwriter.Workbook('SNMP_Report.xlsx')
        worksheet = workbook.add_worksheet()
        row = 1
        worksheet.write(0, 0, "Site_Name")
        worksheet.write(0, 1, "Hostname")
        worksheet.write(0, 2, "Community")
        worksheet.write(0, 3, "Description")
        worksheet.write(0, 4, "Access_List")
        for key, value in dicto.items():
            worksheet.write(row, 0, key)
            col_data = 1
            for k1, v1 in value.items():
                worksheet.write(row, col_data, v1)
                col_data += 1
            row += 1
        workbook.close()
        logging.info("Data Written to Excel File.. Verify SNMP_Report.xlsx file")
        return


if __name__ == "__main__":
    snmp = SnmpCheck()
    flist = snmp.file_list_extract()
    my_master_dict = snmp.master_dict_extract(flist)
    snmp.excel_writer(my_master_dict)
    input("Press Enter to Exit")