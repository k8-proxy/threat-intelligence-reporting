import xmltodict, json
import pandas as pd
import os

files = []
# r=root, d=directories, f = files
for r, d, f in os.walk('bucketOriginal'):
    for file in f:
        if '.xml' in file:
            files.append(os.path.join(r, file))

TechnicalDescription1 = []
InstanceCount1 = []

TechnicalDescription2 = []
InstanceCount2 = []

for item in files:

    try:
        with open(item, 'r') as myfile:
            obj = xmltodict.parse(myfile.read())
    except FileNotFoundError:
        print("Wrong file or file path")

    filetype = obj['gw:GWallInfo']['gw:DocumentStatistics']['gw:DocumentSummary']['gw:FileType']

    if filetype == 'pdf':

        ll = obj['gw:GWallInfo']['gw:DocumentStatistics']['gw:ContentGroups']['gw:ContentGroup']

        for elem in ll:

            ################# REMEDIATION ######################################
            itemCount = int(elem['gw:RemedyItems']['@itemCount'])

            if itemCount != 0:

                lll = elem['gw:RemedyItems']['gw:RemedyItem']

                if itemCount == 1:  # lll is a dict

                    tech_des = lll['gw:TechnicalDescription']
                    ins_count = int(lll['gw:InstanceCount'])

                    if tech_des not in TechnicalDescription1:
                        TechnicalDescription1.append(tech_des)
                        InstanceCount1.append(ins_count)
                    else:
                        ind = TechnicalDescription1.index(tech_des)
                        InstanceCount1[ind] += ins_count

                elif itemCount > 1:  # lll is a list of dicts
                    for el in lll:

                        tech_des = el['gw:TechnicalDescription']
                        ins_count = int(el['gw:InstanceCount'])

                        if tech_des not in TechnicalDescription1:
                            TechnicalDescription1.append(tech_des)
                            InstanceCount1.append(ins_count)
                        else:
                            ind = TechnicalDescription1.index(tech_des)
                            InstanceCount1[ind] += ins_count

            ################# SANITISATION #####################################
            itemCount = int(elem['gw:SanitisationItems']['@itemCount'])

            if itemCount != 0:

                lll = elem['gw:SanitisationItems']['gw:SanitisationItem']

                if itemCount == 1:  # lll is a dict

                    tech_des = lll['gw:TechnicalDescription']
                    ins_count = int(lll['gw:InstanceCount'])

                    if tech_des not in TechnicalDescription2:
                        TechnicalDescription2.append(tech_des)
                        InstanceCount2.append(ins_count)
                    else:
                        ind = TechnicalDescription2.index(tech_des)
                        InstanceCount2[ind] += ins_count

                elif itemCount > 1:  # lll is a list of dicts
                    for el in lll:

                        tech_des = el['gw:TechnicalDescription']
                        ins_count = int(el['gw:InstanceCount'])

                        if tech_des not in TechnicalDescription2:
                            TechnicalDescription2.append(tech_des)
                            InstanceCount2.append(ins_count)
                        else:
                            ind = TechnicalDescription2.index(tech_des)
                            InstanceCount2[ind] += ins_count

df_remedy = pd.DataFrame()

df_remedy['Technical Description'] = TechnicalDescription1
df_remedy['Instances'] = InstanceCount1

df_remedy = df_remedy.sort_values(by='Instances', ascending=False)

df_remedy.to_csv('Remediations.csv', index=False)

df_sanitisation = pd.DataFrame()

df_sanitisation['Technical Description'] = TechnicalDescription2
df_sanitisation['Instances'] = InstanceCount2

df_sanitisation = df_sanitisation.sort_values(by='Instances', ascending=False)

df_sanitisation.to_csv('Sanitisations.csv', index=False)
