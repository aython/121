import zipfile
import os
from zipfile import BadZipfile
import xml.etree.ElementTree as ET
import pandas as pd
from pandas import ExcelWriter
from os.path import basename
import test


class Tabauthj:
    def signinj(self):
        if not os.path.exists(os.path.join('wb','twb')):
            os.makedirs(os.path.join('wb','twb'))
        tableau_auth = TSC.TableauAuth('userid', 'Password',site_id='',user_id_to_impersonate=None)
        server = TSC.Server('https://tableauserverurl.com')
        with server.auth.sign_in(tableau_auth):
            all_datasources, pagination_item = server.datasources.get()
            print("\nThere are {} datasources on site: ".format(pagination_item.total_available))
            print([datasource.name for datasource in all_datasources])
            print('------------------')
            all_sites, pagination_item = server.sites.get()
            for site in all_sites:
                print(site.id, site.name, site.content_url, site.state)
            all_workbooks, pagination_item = server.workbooks.get()
            print([workbook.id for workbook in all_workbooks])
            for workbook in all_workbooks:
            # sample workbook - d20704a2-d1f6-42da-acc2-dc876f3af165
                file_path = server.workbooks.download(workbook.id,filepath=os.path.join('wb',workbook.name+'.twbx'))
                print("\nDownloaded the file to {0}.".format(file_path))
                try:
                    with zipfile.ZipFile(os.path.join('wb',workbook.name+'.twbx')) as zf:
                        for file in zf.namelist():
                            if file.endswith('.twb'):
                                zf.extract(file, os.path.join('wb','twb'))
                except BadZipfile:
                    print("Does not work ")


class Extract:
    def worksheet(self):
        #gapath = input("Provide the path of twb file ")
            gapath = 'Users/aqhibjaveed/Downloads/wb/twb'
            os.chdir(r'/Users/aqhibjaveed/Downloads/wb/twb')
            tree = ET.parse('/Users/aqhibjaveed/Downloads/wb/twb/Superstore.twb')
            root = tree.getroot()
            ws = root.find('worksheets')
            callevent = root.find('datasources')
            dashboard = root.find('windows')
            actions = root.find('actions')
            savefilen = os.path.splitext(basename(gapath))[0]
            writer = ExcelWriter(savefilen + "-workbook-output" + '.xlsx')
            wksheet = []
            dashboardj = []
            allfields = []
            Calcfields = []
            don = []
            srg = []
            if dashboard is not None:
                for dash in dashboard.findall("window"):
                    for win in dash.getiterator():
                        if win.get('class') == 'dashboard':
                            for vp in win.findall("viewpoints"):
                                for vpn in vp.getiterator():
                                    if vpn.get('name',''):
                                        dashboardj.append({'Dashboard': win.get('name',''), 'SheetsLinked': vpn.get('name','')})
            Dashboardn = pd.DataFrame(dashboardj)
            Dashboardn.to_excel(writer, 'Dashboard', index=False)

            if ws is not None:
                for sheet in ws.findall("worksheet"):
                    storesheetname = sheet.get('name', '')
                    for sheets in sheet.getiterator():
                        if sheets is not None:
                            for dsdepen in sheets.findall("datasource-dependencies"):
                                for gdsdepen in dsdepen.getiterator():
                                    if gdsdepen is not None:
                                        for column in gdsdepen.findall('column'):
                                            wksheet.append({'WorksheetName': storesheetname, 'FieldName': column.get('name', ''),'caption': column.get('caption',''),
						                    'Aggregation': column.get('aggregation', ''), 'Datatype': column.get('datatype', ''),
						                    'Default Type': column.get('default-type', '')})
                wshetdg = pd.DataFrame(wksheet)
                wshetdg.to_excel(writer, 'WorkSheets and Fields', index=False,columns=['WorksheetName', 'FieldName', 'caption', 'Aggregation','Datatype','Default Type'])

            if callevent is not None:
                for Moc1 in callevent.findall("datasource"):
                    for node in Moc1.getiterator():
                        if node.tag == 'column':
                            allfields.append({'DatasourceName': Moc1.get('name', ''),'DatasourceCaptionName':Moc1.get('caption',''), 'FieldCaption': node.get('caption', ''),
                                      'FieldName': node.get('name', ''), 'DataType': node.get('datatype', ''),
                                      'Role': node.get('role', ''), 'Type': node.get('type', '')})
                            if node is not None:
                                for cf in node.findall("calculation"):
                                    Calcfields.append({'DatasourceName': Moc1.get('name', ''),'DatasourceCaptionName':Moc1.get('caption',''), 'FieldCaption': node.get('caption', ''),
                                               'FieldName': node.get('name', ''), 'DataType': node.get('datatype', ''),
                                               'Role': node.get('role', ''), 'Type': node.get('type', ''),
                                               'Formula': cf.get('formula', '')})


            if actions is not None:
                for action in actions.findall("action"):
                    don.append({'name': action.get('name',''), 'caption': action.get('caption','')})
                    for act in action.getiterator():
                        if act is not None:
                            for lnk in act.findall("link"):
                                if act is not None:
                                    for src in act.findall("source"):
                                        srg.append({'URL': lnk.get('expression',''), 'Source': src.get('dashboard','')})
        
            allfieldsn = pd.DataFrame(allfields)
            Calcfieldsn = pd.DataFrame(Calcfields)
            rt = pd.DataFrame(don)
            urld = pd.DataFrame(srg)
            allfieldsn.to_excel(writer, 'All Fields', index=False,columns=['DatasourceName','DatasourceCaptionName','FieldCaption','FieldName','DataType','Role','Type'])
            Calcfieldsn.to_excel(writer, 'Calculated Fields', index=False,columns=['DatasourceName','DatasourceCaptionName','FieldCaption','FieldName','DataType','Role','Type','Formula'])
            rt.to_excel(writer, 'Actions', index=False)
            urld.to_excel(writer, 'Action URLS', index=False)
            writer.save()
            print("Extraction completed successfully."+gapath+"  "+"Excel file generated in current directory ending with workbook output")

#q = Tabauthj()
#q.signinj()
p = Extract()
p.worksheet()


