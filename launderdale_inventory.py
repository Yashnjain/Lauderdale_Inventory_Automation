import os
import re
import time
import glob
import logging
import bu_alerts
import pandas as pd
import xlwings as xw
import xlwings.constants as win32c
from datetime import date, timedelta


drive = r"J:\India"

def remove_existing_files(files_location):
    """_summary_

    Args:
        files_location (_type_): _description_

    Raises:
        e: _description_
    """           
    logging.info("Inside remove_existing_files function")
    try:
        files = os.listdir(files_location)
        if len(files) > 0:
            for file in files:
                os.remove(files_location + "\\" + file)
            logging.info("Existing files removed successfully")
        else:
            print("No existing files available to reomve")
        print("Pause")
    except Exception as e:
        logging.exception("Exception in: remove_existing_files()")
        logging.exception(e)
        raise e


def num_to_col_letters(num):
    try:
        letters = ''
        while num:
            mod = (num - 1) % 26
            letters += chr(mod + 65)
            num = (num - 1) // 26
        return ''.join(reversed(letters))
    except Exception as e:
        raise e
    

def xlOpner(inputFile):
    try:
        retry = 0
        while retry<10:
            try:
                input_wb = xw.Book(inputFile, update_links=False)
                return input_wb
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
                    raise e
    except Exception as e:
        print(f"Exception caught in xlOpner :{e}")
        logging.info(f"Exception caught in xlOpner :{e}")
        raise e
    

def remove_special_characters(my_pdf,column_list):
    try:
        # column_list = list(my_pdf.columns[[-5,-4,-3,-2]])
        logging.info("inside remove special characters")
        for values in column_list:
            my_pdf[values] = my_pdf[values].astype(str)
            my_pdf[values]  = [x[values].replace('$', '') for i, x in my_pdf.iterrows()]
            my_pdf[values]  = [x[values].replace('(', '-') for i, x in my_pdf.iterrows()]
            my_pdf[values]  = [x[values].replace(')', '') for i, x in my_pdf.iterrows()]
            my_pdf[values]  = [x[values].replace(',', '') for i, x in my_pdf.iterrows()]
            # my_pdf[values]  = [x[values].replace('0.0', '0.00') for i, x in my_pdf.iterrows()]
            my_pdf[values] = my_pdf[values].astype(float)
            # my_pdf[values]  = [x[values].replace('0.00', '0') for i, x in my_pdf.iterrows()]
            
        return  my_pdf   
    except Exception as e:
        raise e  


def insert_top1_btm2_borders(cellrange:str,working_sheet,working_workbook):
    try:
        working_sheet.activate()
        working_sheet.api.Range(cellrange).Select()
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlDiagonalDown).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlDiagonalUp).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeLeft).LineStyle = win32c.Constants.xlNone
        # linestylevalues=[win32c.BordersIndex.xlEdgeLeft,win32c.BordersIndex.xlEdgeTop,win32c.BordersIndex.xlEdgeBottom,win32c.BordersIndex.xlEdgeRight,win32c.BordersIndex.xlInsideVertical,win32c.BordersIndex.xlInsideHorizontal]
        # for values in linestylevalues:
        a=working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeTop)
        a.LineStyle = win32c.LineStyle.xlContinuous
        a.ColorIndex = 0
        a.TintAndShade = 0
        a.Weight = win32c.BorderWeight.xlThin
        b=working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeBottom)
        b.LineStyle = win32c.LineStyle.xlDouble
        b.ColorIndex = 0
        b.TintAndShade = 0
        b.Weight = win32c.BorderWeight.xlThick
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeRight).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlInsideVertical).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlInsideHorizontal).LineStyle = win32c.Constants.xlNone
    except Exception as e:
        raise e
    

def working(inventory_wb,sales_wb):
     try:
        sales_ws = sales_wb.sheets("Sheet1")
        inv_working_ws = inventory_wb.sheets("Working")
        last_row = sales_ws.range(f'B'+ str(sales_ws.cells.last_cell.row)).end('up').row
        curr_col_list = sales_ws.range("B6").expand('right').value
        sales_terminal = curr_col_list.index('Terminal')
        sales_ws.api.AutoFilterMode= False
        sales_ws.api.Range(f"B6:AR{last_row}").AutoFilter(Field:=f"{sales_terminal+1}", Criteria1="FT LAUDERDALE, FL - TLOAD")
        inv_working_ws.range('A1').expand('down').clear()
        inv_working_ws.range('B1').expand('down').clear()
        inv_working_ws.range('C1').expand('down').clear()
        inv_working_ws.range('D1').expand('down').clear()
        inv_working_ws.range('E1').expand('down').clear()
        inv_working_ws.api.Range("F:F").EntireColumn.Clear()
        inv_working_ws.range('G1').expand('down').clear()
        inv_working_ws.range('H1').expand('down').clear()
        ########## Particulars #########################
        sales_ws.range(f'B6:B{last_row}').copy()
        inv_working_ws.range('A1').paste()
        ########## Date ################################
        sales_ws.range(f'E6:E{last_row}').copy()
        inv_working_ws.range('B1').paste()
        ########## Customer Name #######################
        sales_ws.range(f'I6:I{last_row}').copy()
        inv_working_ws.range('C1').paste()
        ########## BOL Number ##########################
        sales_ws.range(f'L6:L{last_row}').copy()
        inv_working_ws.range('D1').paste()
        ########## BOL Date ############################
        sales_ws.range(f'M6:M{last_row}').copy()
        inv_working_ws.range('E1').paste()
        ########## Billed QTY ##########################
        sales_ws.range(f"AD6:AD{last_row}").copy()
        inv_working_ws.range('F1').paste()
        sales_ws.api.AutoFilterMode= False
        ########## applying formul for G and H column ############  
        lst_rw = inv_working_ws.range(f'A'+ str(inv_working_ws.cells.last_cell.row)).end('up').row
        inv_working_ws.range(f"G2:G{lst_rw}").api.Formula = f"=+VLOOKUP(D2,Outbound!E:J,6,0)"
        inv_working_ws.range(f"H2:H{lst_rw}").api.Formula ="=+G2+F2"
        working_total_rw = lst_rw+2
        inv_working_ws.range(f"F{lst_rw+2}").api.Formula =f"=SUBTOTAL(9,F2:F{lst_rw})"
        insert_top1_btm2_borders(cellrange=f"F{lst_rw+2}",working_sheet=inv_working_ws,working_workbook=inventory_wb)
        df = inv_working_ws.range(f"A2:A{lst_rw}").options(pd.DataFrame,header=False,index=False).value
        if df[0].str.contains("SRT").any():
            df = df.iloc[::-1]
            for i in range(len(df[0])-1,-1,-1):
                if "SRT" in df[0][i]:
                    inv_working_ws.range(f"A{i+2}").api.EntireRow.Delete()
                    print(df[0][i])
        return working_total_rw
     except Exception as e:
          raise e
     

def mrn(inventory_wb,mrn_wb):
     try:
        inventory_wb.activate()  
        inventory_mrndetail_ws = inventory_wb.sheets["MRN Detail"]
        inventory_mrndetail_ws.range("A1:AR1").expand('down').delete()
        mrn_ws = mrn_wb.sheets[0]
        last_row = mrn_ws.range(f'B'+ str(mrn_ws.cells.last_cell.row)).end('up').row
        curr_col_list = mrn_ws.range("B6").expand('right').value
        arrival_date_col = curr_col_list.index('Arrival Date')
        prev = date.today().replace(day=1) - timedelta(days=1)
        mrn_ws.api.Range(f"B6:AR{last_row}").AutoFilter(Field:=arrival_date_col+1,Criteria1:=f">={prev}",Operator:=2,Criteria2:=f"=")
        terminal_col_no = curr_col_list.index('Terminal')
        mrn_ws.api.Range(f"B6:AR{last_row}").AutoFilter(Field:=f"{terminal_col_no+1}", Criteria1="FT LAUDERDALE, FL - TLOAD")
        # sp_address = row_range_calc("B",mrn_ws)
        # mrn_ws.range(f"{sp_address}").copy()
        mrn_ws.range(f"B6:AR{last_row}").copy()
        inventory_mrndetail_ws.range("B1").paste()
        inventory_mrndetail_ws.range("I1").expand('down').copy()
        inventory_mrndetail_ws.range("A1").paste()
     except Exception as e:
          raise e


    
def in_out_inv(inv_path,inventory_wb):
    try:  
        outbound_inv =  inventory_wb.sheets['Outbound']  
        try:
            if len(glob.glob(inv_path+"\\*.xls"))>0:   
                for file in glob.glob(inv_path+"\\*.xls"):
                    path, file_name = os.path.split(file)
                    inout_inv_file_name = file_name
                    try:
                        outbound_wb = xlOpner(file)
                    except Exception as e:
                        logging.info(f"could not open workbook: {file}")
                        raise e  
                    try:
                        outbound_sheet = outbound_wb.sheets[f'{today_date.strftime("%b")}{small_yr} Inv.']
                    except:
                        logging.info(f"could not find sheet in file : {outbound_wb.name}, name of sheet : {today_date.strftime('%b')}{small_yr} Inv.")
                        raise e                             
                    outbound_sheet.activate()

                    ######### Clearing Sheet #####################
                    clr_rw = outbound_inv.api.UsedRange.Rows.Count
                    outbound_inv.range(f"A2:M{clr_rw}").clear_contents()
                    end_row_J = outbound_sheet.range(f'J'+ str(outbound_sheet.cells.last_cell.row)).end('up').row
                    outbound_sheet.activate()
                    outbound_sheet.api.AutoFilterMode=False

                    df = outbound_sheet.range(f"A4:J{end_row_J}").options(pd.DataFrame,header=False).value
                    if len(df)>0:
                        outdf = df[(~df[7].isnull())]
                        outbound_inv.range("B2").options(header=False).value = outdf

                        outbound_inv.range(f"L2").value = f"=+VLOOKUP(E2,Working!D:F,3,0)"
                        outbound_inv.range(f"M2").value = f"=J2+L2"

                        outbound_inv.range(f"L2:M2").copy(outbound_inv.range(f"L2:M{len(outdf)+1}"))                       
                        # =+VLOOKUP(E2,Working!D:F,3,0)
                        # =J2+L2
                    else:
                        logging.info(f"No outbound values to update, check :: {outbound_inv}")    
                    
                    # inventory_wb.api.AutoFilterMode=False
                    # inventory_wb.app.api.CutCopyMode=False

            else:
                logging.info(f"No outbound reports found, please check ::: {inv_path}")
        except Exception as e:
            logging.exception(f"Check {path}:::::{file_name}")
            logging.exception(str(e))
            print("Error while generating outbound sheet")
            raise e
        
        ################# Total Outbound sheet ##################
        check_row_J = outbound_inv.range(f'J'+ str(outbound_inv.cells.last_cell.row)).end('up').row
        outbound_total_rw = check_row_J + 2
        outbound_inv.range(f"J{outbound_total_rw}").value = f"=SUM(J2:J{check_row_J})"
        insert_top1_btm2_borders(cellrange=f"J{outbound_total_rw}",working_sheet=outbound_inv,working_workbook=inventory_wb)

        try:     
            print(f"inbound started")
            inbound_inv = inventory_wb.sheets['Inbound']

            ######### Clearing Sheet #####################
            clr_rwim = inbound_inv.api.UsedRange.Rows.Count
            inbound_inv.range(f"2:{clr_rwim}").api.EntireRow.Delete()
            end_row_J = outbound_sheet.range(f'J'+ str(outbound_sheet.cells.last_cell.row)).end('up').row
            outbound_sheet.activate()
            outbound_sheet.api.AutoFilterMode=False

            if len(df)>0:
                indf = df[(~df[5].isnull())]
                indf = indf.reset_index(drop=True)
                indf = indf.drop(columns =[0,2,3,6,7,8])
                indf = indf.reindex(columns=[4,5,1])

                inbound_inv.range("A2").options(header=False,index=False).value = indf
                inbound_inv.range(f"L2").value = f"=+VLOOKUP(E2,Working!D:F,3,0)"
                inbound_inv.range(f"M2").value = f"=J2+L2"

                inbound_inv.range(f"L2:M2").copy(inbound_inv.range(f"L2:M{len(indf)+1}"))                       
                # =+VLOOKUP(E2,Working!D:F,3,0)
                # =J2+L2

                inbound_inv.range(f"E:E").number_format='_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                inbound_inv.range(f"B:B").number_format='0'
                inbound_inv.range(f"F:F").number_format="m/d/yyyy"
                column_list = inbound_inv.range("A1").expand('right').value

                Diff_column_no = column_list.index('Diff')+1
                list2=[f"=+B2-AJ2",f"=+X2",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,2,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,3,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,4,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,5,0)",\
                        f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,6,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,7,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,8,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,9,0)",\
                        f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,10,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,11,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,12,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,13,0)",\
                        f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,14,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,15,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,16,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,17,0)",\
                        f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,18,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,19,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,20,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,21,0)",\
                        f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,22,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,23,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,24,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,25,0)",\
                        f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,26,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,27,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,28,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,29,0)",\
                        f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,30,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,31,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,32,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,33,0)",\
                        f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,34,0)",f"=VLOOKUP($A2,'MRN Detail'!$A:$AI,35,0)"] 
                last_row = inbound_inv.range(f'A'+ str(inbound_inv.cells.last_cell.row)).end('up').row
                for values in list2:
                    last_column_letter=num_to_col_letters(Diff_column_no)
                    inbound_inv.range(f"{last_column_letter}2").value = values
                    time.sleep(1)
                    inbound_inv.range(f"{last_column_letter}2").copy(inbound_inv.range(f"{last_column_letter}2:{last_column_letter}{last_row}"))
                    Diff_column_no+=1

                end_row_main = inbound_inv.range(f'B'+ str(inbound_inv.cells.last_cell.row)).end('up').row
                if inbound_inv.range(f"B2").value!=None:
                    inbound_total_rw = end_row_main + 2
                    inbound_inv.range(f"B{inbound_total_rw}").value = f"=SUM(B2:B{end_row_main})"
                    insert_top1_btm2_borders(cellrange=f"B{inbound_total_rw}",working_sheet=inbound_inv,working_workbook=inventory_wb)  
            else:
                logging.info(f"No outbound values to update, check :: {inbound_inv}") 

        except Exception as e:
            logging.exception(str(e))
            logging.exception(f"Check {path}:::::{file_name}")
            print("Error while generating Inbound sheet")
            raise e
        
        ############### updating summary tab ##################
        summary_inv = inventory_wb.sheets['Summary']
        summary_inv.activate()
        sum_end_rw = summary_inv.api.UsedRange.Rows.Count
        summary_inv.range(f"3:{sum_end_rw}").api.EntireRow.Delete()
        c_row_main = outbound_sheet.range(f'C'+ str(outbound_sheet.cells.last_cell.row)).end('up').row
        outbound_sheet.range(f"A4:J{c_row_main}").copy(summary_inv.range(f"A3"))


        sum_J_end = summary_inv.range(f'J'+ str(summary_inv.cells.last_cell.row)).end('up').row
        summary_inv.range(f"J{sum_J_end}").copy(summary_inv.range(f"J{sum_J_end}:J{sum_J_end+3}"))

        ############### updating costing tab ##################
        costing_inv = inventory_wb.sheets['Costing']
        costing_inv.activate()
        cos_end_rw = costing_inv.api.UsedRange.Rows.Count
        costing_inv.range(f"2:{cos_end_rw}").api.EntireRow.Delete()        
        if len(df)>0:
            col = 1
            list2=[f"=+Inbound!A2",f"=+Inbound!AJ2",f"=+Inbound!M2",f"=+Inbound!G2",f"=+Inbound!L2",f"=+Inbound!W2",\
                        f"=+Inbound!C2",f"=+Inbound!H2",f"=+Inbound!AL2",f"=+K2/B2","=+Inbound!AK2"] 
            last_row = inbound_inv.range(f'A'+ str(inbound_inv.cells.last_cell.row)).end('up').row
            for values in list2:
                    last_column_letter=num_to_col_letters(col)
                    costing_inv.range(f"{last_column_letter}2").value = values
                    time.sleep(1)
                    col+=1
        costing_inv.range(f"A2:K2").copy(costing_inv.range(f"A2:K{last_row}"))

        return inout_inv_file_name
    except Exception as e:
        raise e


if __name__ == "__main__":
    try:

        job_name="BIO_PAD01_Ft Lauderdale_INV_AUTOMATION"
        # log progress --
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)

        logfile = os.getcwd() + '\\' + 'logs' + '\\' + f'{job_name}.txt'

        logging.basicConfig(
            level=logging.INFO, 
            format='%(asctime)s [%(levelname)s] - %(message)s',
            filename=logfile)

        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        logging.info("Execution Started")

        locations_list = []
        # logging.info('setting receiver_email')
        # receiver_email = "yashn.jain@biourja.com"
        receiver_email = "yashn.jain@biourja.com,imam.khan@biourja.com,apoorva.kansara@biourja.com, accounts@biourja.com, rini.gohil@biourja.com,itdevsupport@biourja.com"


        time_start=time.time()
        today_date=date.today()
        inv_path = r'J:\India\Inv Rpt\IT_INVENTORY\flows\Ft Lauderdale'
        if len(glob.glob(inv_path+"\\*.xls"))>0:
            inv_file = glob.glob(inv_path+"\\*.xls")[0]    
            pathinv, file_name_inv = os.path.split(inv_file)
            year = today_date.year
            pre_month = int(re.findall("\d+",file_name_inv)[0]) - 1
            pre_date = today_date.replace(month=pre_month)
            today_date = today_date.replace(month=int(re.findall("\d+",file_name_inv)[0]))
            pre_date_fldr = pre_date.strftime("%m-%y")
            date_fldr = today_date.strftime("%m-%y")
            small_yr = today_date.strftime("%y")
        else:
            logging.info(f"inventory report not found ::: {inv_path}")   
            locations_list.append(logfile)
            bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name}',mail_body = f'{job_name} completed successfully,Inventory file not found here ::: {inv_path}',multiple_attachment_list = locations_list)
                 

        inventory_sheet = drive+rf'\{year}\{date_fldr}'+f'\\Ft Lauderdale Tload.xlsx'
        if not os.path.exists(inventory_sheet):

            logging.info(f"{inventory_sheet} Excel file not present")           

        mrn_sheet = drive+rf'\{year}\{date_fldr}'+f'\\MRN.xlsx'
        if not os.path.exists(mrn_sheet):
            logging.info(f"{mrn_sheet} Excel file not present")

        sales_sheet = drive+rf'\{year}\{date_fldr}'+f'\\Sales.xlsx'
        if not os.path.exists(sales_sheet):
            logging.info(f"{sales_sheet} Excel file not present")

        pre_month_sheet = drive+rf'\{year}\{pre_date_fldr}\Transfered'+f'\\Ft Lauderdale Tload.xlsx'
        if not os.path.exists(pre_month_sheet):
            logging.info(f"{pre_month_sheet} Excel file not present")


        try:
            inventory_wb = xlOpner(inventory_sheet)
        except Exception as e:
            logging.info(f"could not open workbook: {inventory_sheet}")
            raise e
        
        try:
            mrn_wb = xlOpner(mrn_sheet)
        except Exception as e:
            logging.info(f"could not open workbook: {mrn_sheet}")
            raise e   

        try:
            sales_wb = xlOpner(sales_sheet)
        except Exception as e:
            logging.info(f"could not open workbook: {sales_sheet}")
            raise e 
            
        inventory_wb.api.AutoFilterMode=False
        inventory_wb.app.api.CutCopyMode=False
        sales_wb.api.AutoFilterMode=False
        sales_wb.app.api.CutCopyMode=False 

        try:
            working_total_rw = working(inventory_wb,sales_wb)
        except Exception as e:
            logging.info(f"Sales Tab Failure : {e}")
            raise e  

        sales_wb.api.AutoFilterMode=False
        sales_wb.app.api.CutCopyMode=False         

        inventory_wb.api.AutoFilterMode=False
        inventory_wb.app.api.CutCopyMode=False
        mrn_wb.api.AutoFilterMode=False
        mrn_wb.app.api.CutCopyMode=False  

        try:
            mrn(inventory_wb,mrn_wb)
        except Exception as e:
            logging.info(f"Mrn Tab Failure : {e}")
            raise e 
        inventory_wb.api.AutoFilterMode=False
        inventory_wb.app.api.CutCopyMode=False
        mrn_wb.api.AutoFilterMode=False
        mrn_wb.app.api.CutCopyMode=False        
        print("sales and mrn done")

        try:
            inout_inv_file_name = in_out_inv(inv_path,inventory_wb)
        except Exception as e:
            logging.info(f"Inbound/Outbound Tab Failure : {e}")
            raise e        
        print("Done")
        
        output_location = rf'J:\India\Inv Rpt\IT_INVENTORY\Output\{year}\{date_fldr}\Lauderdale'
        if not os.path.exists(output_location):
            os.makedirs(output_location)

        try:
            inventory_wb.save(f"{output_location}\\Ft Lauderdale Tload.xlsx")
            print(f"inventory done and saved in {output_location}")
            wb_name = inventory_wb.name
            inventory_wb.app.quit()
        except Exception as e:
            logging.info(f"could not save or kill ::: {output_location}")
            raise e 



        time.sleep(2)
        remove_existing_files(inv_path)
        logging.info(f"files succesfully removed from folder :::: {inv_path}")
        locations_list.append(logfile)
        locations_list.append(f"{output_location}\\Ft Lauderdale Tload.xlsx")
        nl = '<br>'
        body = ''
        body = (f'{nl}<strong>{wb_name}</strong> {nl}{nl} <strong>{wb_name}</strong> successfully created from reports <strong>{inout_inv_file_name}</strong>, {nl} Attached path for the excel=<u>{output_location}</u>{nl}')
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name}',mail_body = f'{body}{job_name} completed successfully, Attached Logs and Excel',multiple_attachment_list = locations_list)
        logging.info("Process completed")
        print("process completed")

    except Exception as e:
        logging.exception(str(e))
        try:
            inventory_wb.app.quit()
        except:
            pass    
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB FAILED -{job_name}',mail_body = f'{job_name} failed in __main__, Attached logs',attachment_location = logfile)

