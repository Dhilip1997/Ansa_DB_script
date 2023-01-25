from flask import Flask, render_template, request, session
import pandas as pd
from threading import Timer
import webbrowser
import os
app = Flask('__name__',template_folder='templates',static_url_path='/static')

app.secret_key = "27eduCBA089"
def column_reference(input_str):
    col_cell_reference = {
        'A' : 0,'B' : 1,'C' : 2,'D' : 3,'E' : 4,'F' : 5,'G' : 6,'H' : 7,'I' : 8,'J' : 9,'K' : 10,'L' : 11,'M' : 12,
        'N' : 13,'O' : 14,'P' : 15,'Q' : 16,'R' : 17,'S' : 18,'T' : 19,'U' : 20,'V' : 21,'W' : 22,'X' : 23,'Y' : 24,
        'Z' : 25,'AA' :26,'AB' : 27,'AC' : 28,'AD' : 29,'AE' : 30,'AF' : 31,'AG' : 32,'AH' : 33,'AI' : 34,"AJ" : 35,
        'AK' : 36,'AL' : 37,'AM' : 38,'AN' : 39
    }
    column_cell = col_cell_reference[input_str]
    return column_cell
file_input = {}
@app.route('/')
def display():
    session.clear()
    print(session)
    return render_template("index.html")

@app.route('/file_upload', methods = ["GET","POST"])
def file_upload():
    if request.method == 'POST':
        steel_input_file = request.files['Steel_sheet_input']
        steel_other_file = request.files['steel_other_input']
        aluminium_input_file = request.files['aluminium_input']
        file_input['steel_input_file'] = [steel_input_file.filename, False]
        file_input['steel_other_input_file'] = [steel_other_file.filename, False]
        file_input['aluminium_input_file'] = [aluminium_input_file.filename, False]
        if steel_input_file:
            session['steel_input'] =  file_input['steel_input_file'][0]
            file_input['steel_input_file'][1] = True
        if steel_other_file:
            session['steel_other_input'] = file_input['steel_other_input_file'][0]
            file_input['steel_other_input_file'][1] = True
        if aluminium_input_file:
            session['aluminium_input'] = file_input['aluminium_input_file'][0]
            file_input['aluminium_input_file'][1] = True
        print(file_input)
    return render_template("index.html", filename = file_input)    
            
@app.route('/output', methods = ["GET", "POST"])
def output():
    if request.method == "POST": 
        if 'steel_input' in session:
            steel_df = pd.read_excel(session.get('steel_input'), sheet_name = 'Steel sheet ')
            steel_file_name = request.form.get('steel_name')
            steel_row_number = int(request.form.get('Steel_row')) - 2
            steel_col_number = column_reference(request.form.get('steel_column').upper())
            output_file_input = request.form.get("file_path")
            print("out",output_file_input)
            MID_list = steel_df.iloc[7:,3].to_list()
            material_list = steel_df.iloc[7:,1].to_list()
            thickness_list = steel_df.iloc[7:,2].to_list()
            Criteria_list = steel_df.iloc[steel_row_number:, steel_col_number].to_list()
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

                
            with open(rf'{output_file_input}\{steel_file_name}.txt', 'w',encoding="utf-8") as f:
                for line in text_list:
                    f.write(line)
                    f.write('\n')
            return render_template("index.html", result_text = "Dat file Generated Successfully")
           
        elif 'steel_other_input' in session:
            steel_other_df = pd.read_excel(session.get('steel_other_input'), sheet_name = 'Steel other')
            steel_other_file_name = request.form.get('Steel_other')
            steel_other_row_number = int(request.form.get('steel_other_row')) - 2
            steel_other_col_number = column_reference(request.form.get('steel_other_col').upper())
            MID_list = steel_other_df.iloc[7:,3].to_list()
            material_list = steel_other_df.iloc[7:,1].to_list()
            thickness_list = steel_other_df.iloc[7:,2].to_list()
            Criteria_list = steel_other_df.iloc[steel_other_row_number:, steel_other_col_number].to_list()
            output_file_input = request.form.get("file_path")
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

                
            with open(rf'{output_file_input}\{steel_other_file_name}.txt', 'w',encoding="utf-8") as f:
                for line in text_list:
                    f.write(line)
                    f.write('\n')
            return render_template("index.html", result_text = "Dat file Generated Successfully")
            
        elif 'aluminium_input' in session:
            aluminium_input_df = pd.read_excel(session.get('aluminium_input'), sheet_name = 'Aluminum ' )
            aluminium_file_name = request.form.get('aluminium_name')
            aluminium_row_number = int(request.form.get('aluminium_row')) - 2
            aluminium_col_number = column_reference(request.form.get('aluminium_col').upper())
            MID_list = aluminium_input_df.iloc[7:,3].to_list()
            material_list = aluminium_input_df.iloc[7:,1].to_list()
            thickness_list = aluminium_input_df.iloc[7:,2].to_list()
            Criteria_list = aluminium_input_df.iloc[aluminium_row_number:, aluminium_col_number].to_list()
            output_file_input = request.form.get("file_path")
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

                
            with open(rf'{output_file_input}\{aluminium_file_name}.txt', 'w',encoding="utf-8") as f:
                for line in text_list:
                    f.write(line)
                    f.write('\n')
            return render_template("index.html", result_text = "Dat file Generated Successfully")


def main():
    if not os.environ.get("WERKZEUG_RUN_MAIN"):
        Timer(1, webbrowser.open_new('http://127.0.0.1:9000/')).start();
    app.run(debug=True, port=(9000))        


if __name__ == "__main__":
    main()


