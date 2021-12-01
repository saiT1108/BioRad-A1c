import pandas as pd
import sqlite3
## Parses through excel file to send the data to 

def getMachine(file):
    data = pd.read_excel(rf'{file}')
    machine_IDs = pd.DataFrame(data, columns=['Machine_ID'])
    types = pd.DataFrame(data, columns=['Machine_Type'])
    labs = pd.DataFrame(data, columns=['Lab_ID'])
    macs = machine_IDs.values.tolist()
    types_l = types.values.tolist()
    labs_l = labs.values.tolist()

    return {
        "IDs" : macs,
        "types" : types_l,
        "labs" : labs_l,
    }

def add_machine(memo):
    connection = sqlite3.connect(r"./test.db")
    for i in range(len(memo["IDs"])):
        entry = f"'{memo['IDs'][i][0]}', '{memo['types'][i][0]}', '{memo['labs'][i][0]}'"
        print(entry)
        try:
            with connection:
                sql = f"INSERT INTO biorad_machine (Machine_ID, Machine_Type, Lab_ID) VALUES ({entry})"
                cur = connection.cursor()
                cur.execute(sql)
                connection.commit()
        except:
            print("Machine already exists")
            continue

    connection.close()

def getSample(file):
    data = pd.read_excel(rf'{file}')
    date1 = pd.DataFrame(data, columns=['Date_Acq']).values.tolist()
    date2 = pd.DataFrame(data, columns=['Date_Sub']).values.tolist()
    sids = pd.DataFrame(data, columns=['Sample_ID']).values.tolist()
    lids = pd.DataFrame(data, columns=['Lab_ID']).values.tolist()

    return {
        "acq" : date1,
        "sub" : date2,
        "sids" : sids,
        "lids" : lids
    }

def add_sample(memo):
    connection = sqlite3.connect(r"./test.db")
    for i in range(len(memo["acq"])):
        entry = f"'{memo['acq'][i][0]}', '{memo['sub'][i][0]}', '{memo['sids'][i][0]}', '{memo['lids'][i][0]}'"
        print(entry)
        try:
            with connection:
                sql = f"INSERT INTO blood_sample (Date_Acquired, Date_Submitted, Sample_id, Lab_id) VALUES ({entry})"
                print(sql)
                cur = connection.cursor()
                cur.execute(sql)
                connection.commit()
        except:
            print("Sample already exists")
            continue
    
    connection.close()

def getClient(file):
    data = pd.read_excel(rf'{file}')
    cids = pd.DataFrame(data, columns=['Client_ID']).values.tolist()
    fn = pd.DataFrame(data, columns=['First_Name']).values.tolist()
    ln = pd.DataFrame(data, columns=['Last_Name']).values.tolist()
    phones = pd.DataFrame(data, columns=['Phone_Number']).values.tolist()

    return {
        "cids" : cids,
        "fn" : fn,
        "ln" : ln,
        "phones" : phones
    }

def add_client(memo):
    connection = sqlite3.connect(r"./test.db")
    for i in range(len(memo["cids"])):
        entry = f"'{memo['cids'][i][0]}', '{memo['fn'][i][0]}', '{memo['ln'][i][0]}', '{memo['phones'][i][0]}'"
        try:
            print(entry)
            with connection:
                sql = f"INSERT INTO client (Client_ID, First_Name, Last_Name, Phone_Number) VALUES ({entry})"
                print(sql)
                cur = connection.cursor()
                cur.execute(sql)
                connection.commit()
        except:
            print("Client already exists")
            continue
    
    connection.close()


add_sample(getSample('./samples.xlsx'))
add_machine(getMachine('./test_excel.xlsx'))
add_client((getClient('./client.xlsx')))