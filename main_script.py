import pandas as pd
import openpyxl


#Removes duplicate entries based on hostname
def remove_duplicates(data_logs):
    data_logs.drop_duplicates(subset = "Name", keep = "first", inplace = True)


#Segregates the entries into server and workstation  
def segregate(data_logs):
    updated_server = 0
    outdated_server = 0
    updated_ws = 0
    outdated_ws = 0
    df_server = pd.DataFrame()
    df_workstation = pd.DataFrame()
    df_server = data_logs[data_logs['Platform'].str.contains("Server", na = False)]
    df_workstation = data_logs[~data_logs['Platform'].str.contains("Server", na = False)]
    compliance(df_server,df_workstation)


#Calculates compliance and stores it in a new excel document
def compliance(df_server,df_workstation):
    updated_server, updated_ws, outdated_server, outdated_ws = 0 ,0 ,0, 0
    server_comp = 0
    ws_copm = 0
    for status in df_server['Security Update Status']:
        if status == "Up-to-Date":
            updated_server += 1
    for status in df_workstation['Security Update Status']:
        if status == "Up-to-Date":
            updated_ws += 1
    outdated_server = len(df_server) - updated_server
    outdated_ws = len(df_workstation) - updated_ws
    server_comp = (updated_server/len(df_server)) * 100
    ws_comp = (updated_ws/len(df_workstation)) * 100
    output_df[' '] = ['Total', 'Up-to-Date', 'Outdated','Compliance(in %)']
    output_df['Server'] = [len(df_server), updated_server,outdated_server, server_comp]
    output_df['Workstation'] = [len(df_workstation), updated_ws,outdated_ws,ws_comp]
    output_df.to_excel("AV Complaince Report.xlsx", index = False)




if __name__ =="__main__":

    path = "dump.csv"

    data_logs = pd.read_csv(path)
    output_df = pd.DataFrame()
    remove_duplicates(data_logs)
    segregate(data_logs)
    
    



