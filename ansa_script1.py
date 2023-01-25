import pandas as pd

input_file = pd.read_excel(r"C:\Users\Dhilitp\Desktop\Daass Automation\Material_data_190209_FUSO.xlsx", sheet_name='Steel sheet ')
print(input_file)
MID_list = input_file.iloc[7:,3].to_list()
material_list = input_file.iloc[7:,1].to_list()
thickness_list = input_file.iloc[7:,2].to_list()
Criteria_list = input_file.iloc[7:, 8].to_list()

material_thickness_list = []
for material, thickness in zip(material_list, thickness_list):
    join_string = str(material) +"--"+ (str(thickness))
    if join_string.replace('nan','').split('--')[1]:
        print(join_string)
        material_thickness_list.append(join_string)
    else:
        material_thickness_list.append(join_string.replace('nan','').replace('--',''))
        
text_list = ['! eCanter_T973X-07_ECELE6_EG8907', '   $ELRESRULE  NAME = st_fuso  TYPE = STRESS','!','    MAT_1000	   , 105.  :optional    !SEAM-PLUG_WELDS',\
             '    MAT_1001	   , 105.  :optional    !SPOT_WELDS','    MAT_1002	   , 119.  :optional	!Aluminum_WELDS' ]

for i in range(len(material_list)):
    text_list.append(f'    MAT_{MID_list[i]}	   , {int(round(Criteria_list[i]*0.7, 0))}.   :optional     !{material_thickness_list[i]}')

    
with open('newmatdb\static.txt', 'w',encoding="utf-8") as f:
    for line in text_list:
        f.write(line)
        f.write('\n')
        

    
    
    
        