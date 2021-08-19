# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

from datetime import datetime, date
from logging import debug
from apis import api
from excel import Excel



def get_vuln_data():
    vuln_data = api.get_registries()
    # print('Vuln_data : ', vuln_data)
    if vuln_data == 'REGISTRY_NOT_FOUND':
        raise ValueError("REGISTRY_NOT_FOUND")
    return vuln_data

def get_instructions_from_history(history):
    instructions = dict()
    for step in history:
        instr = step.get('instruction')
        if instr:
            try:
                key, value = instr.split("=", 1)
                if instructions.get(key):
                    value = f'{instructions.get(key)};{value}'
                instructions.update({key: value})
            except:
                pass
    return instructions


def get_all_collections(vuln_data):
    collection = set()
    for coll in vuln_data:
        collection.update(coll['collections'])
    return collection


def is_valid_name(collection_name):
    digits_in_name = any(chr.isdigit() for chr in collection_name)
    return not digits_in_name


def filename_for_collection(collection_name):
    today = date.today().strftime('%Y-%m-%d')
    filename = today + "-" + collection_name.replace(" ", "_") + ".xlsx"
    return filename


# def row_from_collection()


def get_all_high_low_critical_findings_for_registries():

    try:
        vuln_data = get_vuln_data()
    except ValueError as e:
        return e

    collection = get_all_collections(vuln_data)

    for collection_name in collection:
        if is_valid_name(collection_name):
            filename = filename_for_collection(collection_name)
            excel = Excel(filename)
            print('=======================       writing to file', filename)
            collection_rows = list()
            for index in range(len(vuln_data)):    
                if collection_name in vuln_data[index]['collections']:
                    print(f'found matching collection - {collection_name}')
                    ID = vuln_data[index]['repoTag']['repo']
                    Registry = vuln_data[index]['repoTag']['registry']
                    Sha_id = vuln_data[index].get('topLayer') # using get as value is not always present 
                    Tag = vuln_data[index]['repoTag']['tag']
                    Distro = vuln_data[index]['distro']
                    instructions = get_instructions_from_history(vuln_data[index]['history'])
                    maintainer = instructions.get('LABEL maintainer', "None")
                    Application = instructions.get('LABEL applicationname', "None")
                    Cost = instructions.get('LABEL costcenter', "None")
                    Email = instructions.get('LABEL emaildistribution', "None")
                    AppID = instructions.get('LABEL applicationid', "None")
                    APM = instructions.get('LABEL apm', "None")
                    BIT = instructions.get('LABEL bit', "None")
                    result = vuln_data[index]['vulnerabilities']
                    repo = vuln_data[index]['repoTag']['repo']
                    vulnerabilities = vuln_data[index]['vulnerabilities']
                    if vulnerabilities:
                        # calculate count only if status of the issue is available
                        Critical = len(list(filter(lambda x: x.get('status') and x.get('severity') == 'critical', vulnerabilities)))
                        High = len(list(filter(lambda x: x.get('status') and x.get('severity') == 'high', vulnerabilities)))
                        Medium = len(list(filter(lambda x: x.get('status') and x.get('severity') == 'medium', vulnerabilities)))
                        Low = len(list(filter(lambda x: x.get('status') and x.get('severity') == 'low', vulnerabilities)))
                    else:
                        Critical = vuln_data[index]['vulnerabilityDistribution']['critical']
                        High = vuln_data[index]['vulnerabilityDistribution']['high']
                        Medium = vuln_data[index]['vulnerabilityDistribution']['medium']
                        Low = vuln_data[index]['vulnerabilityDistribution']['low']

                    Total = Critical + High + Medium + Low

                    row = [ID, Registry, Sha_id, Tag, Distro, Critical, High, Medium, Low, Total, maintainer, Application, Cost, Email, APM, BIT, result, repo, vulnerabilities]
                    collection_rows.append(row)
                    print(f'appended row - {len(collection_rows)}')
            excel.write_to_excel(collection_rows)

            excel._workbook.close()


if __name__ == '__main__':
    get_all_high_low_critical_findings_for_registries()

