import pandas as pd

input_file = pd.read_excel(r"C:\Users\Dhilitp\Desktop\Daass Automation\Material_data_190209_FUSO.xlsx", sheet_name='Aluminum ')


material_list = input_file.iloc[7:,1].to_list()
thickness_list = input_file.iloc[7:,2].to_list()
MID_list = input_file.iloc[7:,3].to_list()
young_list = input_file.iloc[7:,4].to_list()
poisson_list = input_file.iloc[7:,5].to_list()
density_list = input_file.iloc[7:,6].to_list()

material_thickness_list = []
for material, thickness in zip(material_list, thickness_list):
    join_string = str(material) +"--"+ (str(thickness))
    if join_string.replace('nan','').split('--')[1]:
        print(join_string)
        material_thickness_list.append(join_string)
    else:
        material_thickness_list.append(join_string.replace('nan','').replace('--',''))


text_list = ['$ENTER MATERIAL']
for i in range(len(material_list)):
    text_list.append(f'  $MATERIAL  NAME = MAT_{MID_list[i]}')
    text_list.append(f"$COMMENT '{material_thickness_list[i]}'")
    text_list.append("    $DENSITY")
    text_list.append(f'      {density_list[i]}E-9')
    text_list.append(f'    $ELASTIC')
    text_list.append(f'                        {young_list[i]}.      {poisson_list[i]}')
    text_list.append('    $DAMPING  INPUT=DATA')
    text_list.append('                              0.')
    text_list.append('  $END MATERIAL')
text_list.append('$EXIT MATERIAL')



with open('Aluminiummatdb.txt', 'w',encoding="utf-8") as f:
    for line in text_list:
        f.write(line)
        f.write('\n')




  


