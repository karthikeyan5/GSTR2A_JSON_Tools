import json
import zipfile
import  os
import xlsxwriter


folder_path = './data' # folder path containing the GSTR2A zip files
excel_file_name = './data/GSTR2.xlsx'


workbook = xlsxwriter.Workbook(excel_file_name)
worksheet_b2b = workbook.add_worksheet('b2b-and-b2ba')
row = 0
worksheet_b2b.write_row(row,0,['Filing Period','Supplier Filing Status','Supplier GSTIN','Supplier Name','Invoice Number',
                                'Invoice Date','GSTIN of the taxpayer','Place of Supply (State Code)','Reverse Charge',
                                'Taxable Amount','IGST Amount','CGST Amount','SGST Amount','Cess Amount','Invoice Value','Tax Rate','Old Invoice Number','Old Invoice Date','b2ba'])

def main():
    global row
    for file in os.listdir(folder_path):
        if file[-3:] == 'zip':
            with zipfile.ZipFile(folder_path + '/' +file) as myzip:
                # print(myzip.namelist()[0])
                if myzip.namelist()[0][-4:] == 'json':
                    with myzip.open(myzip.namelist()[0]) as myfile:
                        parsed_json = json.loads(myfile.read().decode('utf-8'))
                        # print(parsed_json.keys())
                        # dict_keys(['gstin', 'cdn', 'b2b', 'fp'])
                        print(parsed_json['fp'])
                        doc_type_list = ['b2b','b2ba']
                        for doc_type in doc_type_list:
                            if doc_type in parsed_json:
                                for supplier in parsed_json[doc_type]:
                                    for inv in supplier['inv']:
                                        inv_values = calc_inv_value(inv['itms'])
                                        if doc_type == 'b2ba':
                                            oinum,oidt,b2ba = (inv['oinum'],inv['oidt'],'Y')
                                        else:
                                            oinum,oidt,b2ba = ('','','')
                                        row += 1
                                        worksheet_b2b.write_row(row,0,[parsed_json['fp'],supplier['cfs'],
                                        supplier['ctin'],supplier['cname'] if 'cname' in supplier else '',
                                        inv['inum'],inv['idt'],parsed_json['gstin'],parsed_json['gstin'][:2],
                                        inv['rchrg'],inv_values['txval'],inv_values['iamt'],inv_values['camt'],
                                        inv_values['samt'],inv_values['csamt'],inv['val'],','.join(str(x) for x in inv_values['tax_rate']),oinum,oidt,b2ba])


def calc_inv_value(itms):
    total = {
        'txval':0,
        'iamt':0,
        'camt':0,
        'samt':0,
        'csamt':0,
        'val':0,
        'tax_rate':[]
    }

    for itm in itms:
        for key in total.keys():
            if key in itm['itm_det']:
                total[key] += itm['itm_det'][key]
                total['val'] += itm['itm_det'][key]
        if 'rt' in itm['itm_det']:
            if itm['itm_det']['rt'] not in total['tax_rate']:
                total['tax_rate'].append(itm['itm_det']['rt'])

    return total


def cleanup():
  workbook.close()
  print("...bye")


if __name__ == '__main__':
  try:
    main()
  finally:
    cleanup()