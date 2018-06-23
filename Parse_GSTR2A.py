import json
import zipfile
import  os
import xlsxwriter


folder_path = './data' # folder path containing the GSTR2A zip files
excel_file_name = 'GSTR2.xlsx'


workbook = xlsxwriter.Workbook(excel_file_name)
worksheet_b2b = workbook.add_worksheet('b2b')
row = 0
worksheet_b2b.write_row(row,0,['Filing Period','Supplier Filing Status','Supplier GSTIN','Supplier Name','Invoice Number',
                                'Invoice Date','Customer GSTIN','Place of Supply (State Code)','Reverse Charge',
                                'Taxable Amount','IGST Amount','CGST Amount','SGST Amount','Cess Amount','Invoice Value'])

zipfile_list = []

def main():
    global row
    for file in os.listdir(folder_path):
        if file[-3:] == 'zip':
            with zipfile.ZipFile(folder_path + '/' +file) as myzip:
                print(myzip.namelist()[0])
                if myzip.namelist()[0][-4:] == 'json':
                    with myzip.open(myzip.namelist()[0]) as myfile:
                        parsed_json = json.loads(myfile.read().decode('utf-8'))
                        print(parsed_json.keys())
                        # dict_keys(['gstin', 'cdn', 'b2b', 'fp'])
                        if 'b2b' in parsed_json:
                            for supplier in parsed_json['b2b']:
                                for inv in supplier['inv']:
                                    inv_values = calc_inv_value(inv['itms'])
                                    print(parsed_json['fp'],supplier['cfs'],supplier['ctin'],supplier['cname'],
                                    inv['inum'],inv['idt'],parsed_json['gstin'],parsed_json['gstin'][:2],
                                    inv['rchrg'],inv_values['txval'],inv_values['iamt'],inv_values['camt'],
                                    inv_values['samt'],inv_values['csamt'],inv['val'])
                                    row += 1
                                    worksheet_b2b.write_row(row,0,[parsed_json['fp'],supplier['cfs'],
                                    supplier['ctin'],supplier['cname'],
                                    inv['inum'],inv['idt'],parsed_json['gstin'],parsed_json['gstin'][:2],
                                    inv['rchrg'],inv_values['txval'],inv_values['iamt'],inv_values['camt'],
                                    inv_values['samt'],inv_values['csamt'],inv['val']])


def calc_inv_value(itms):
    total = {
        'txval':0,
        'iamt':0,
        'camt':0,
        'samt':0,
        'csamt':0,
        'val':0
    }

    for itm in itms:
        for key in total.keys():
            if key in itm['itm_det']:
                total[key] += itm['itm_det'][key]
                total['val'] += itm['itm_det'][key]

    return total


def cleanup():
  workbook.close()
  print("...bye")


if __name__ == '__main__':
  try:
    main()
  finally:
    cleanup()