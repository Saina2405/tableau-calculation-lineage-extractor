
import pandas as pd, os, re, string, webbrowser

from tableaudocumentapi import Workbook
from os.path import isfile, join

import Excelcreator as exg

pd.set_option('display.max_columns', None)


# ## File Handling


input_path = "inputs"
output_path = "outputs"

mypath = "./{}".format(input_path)   #./ points to "this path" as a relative path

#only gets files and not directories within the inputs folder 
input_files = [f for f in os.listdir(mypath) if isfile(join(mypath, f))] 
input_files



def removeSpecialCharFromStr(spstring):
    
#     """
#     input: string
#     output: new string, without any special char
#     """
    
    return ''.join(e for e in spstring if e.isalnum())


def removeSpecialCharFromStr_leaveSpaces(spstring):
  
    return ''.join(e for e in spstring if (e.isalnum() or e ==' '))


def remove_sp_char_then_turn_spaces_into_underscore(string_to_convert):
    filtered_string = re.sub(r'[^a-zA-Z0-9\s_]', '', string_to_convert).replace(' ', "_")
    return filtered_string


def remove_sp_char_leave_undescore_square_brackets(string_to_convert):
    # Clean special characters from strings while keeping underscores and square brackets
    filtered_string = re.sub(r'[^a-zA-Z0-9\s._\[\]]', '', string_to_convert).replace(' ', "_")
    return filtered_string

def find_tableau_file(inputfile):
    """Return the input filename if it is a .twb or .twbx, else return empty string.

    This does not rename files. It simply filters supported types.
    """

    if inputfile.lower().endswith('.twb') or inputfile.lower().endswith('.twbx'):
        return inputfile
    return ""


selected_file = None
for i in input_files:
    candidate = find_tableau_file(i)
    if candidate:
        selected_file = candidate
        break

if not selected_file:
    raise FileNotFoundError(f"No .twb or .twbx file found in '{input_path}'")

print('Selected Tableau file: ' + selected_file)

# substring to be used when naming the exported data (strip extension)
tableau_name_substring = os.path.splitext(selected_file)[0][:30]
print('\nOutput docs name: ' + tableau_name_substring)

packagedTableauFile_relPath = os.path.join(input_path, selected_file)

# # Doc API
# get all fields in workbook
# Attempt to open as a plain .twb first; if that fails and file is a .twbx, try to extract the .twb from it.
TWBX_Workbook = None
try:
    TWBX_Workbook = Workbook(packagedTableauFile_relPath)
except Exception:
    # If the file is a packaged workbook (.twbx), try to extract contained .twb
    if packagedTableauFile_relPath.lower().endswith('.twbx'):
        import zipfile, tempfile, shutil

        with zipfile.ZipFile(packagedTableauFile_relPath, 'r') as z:
            twb_name = next((n for n in z.namelist() if n.lower().endswith('.twb')), None)
            if twb_name is None:
                raise RuntimeError(f"No .twb found inside packaged workbook: {packagedTableauFile_relPath}")

            tmpdir = tempfile.mkdtemp(prefix='twbx_extract_')
            try:
                z.extract(twb_name, tmpdir)
                twb_path = os.path.join(tmpdir, twb_name)
                TWBX_Workbook = Workbook(twb_path)
            finally:
                # clean up the extracted files
                try:
                    shutil.rmtree(tmpdir)
                except Exception:
                    pass
    else:
        raise

collator = []
calcID = []
calcID2 = []
calcNames = []

c = 0

for datasource in TWBX_Workbook.datasources:
    datasource_name = datasource.name
    datasource_caption = datasource.caption if datasource.caption else datasource_name

    for count, field in enumerate(datasource.fields.values()):
        dict_temp = {
            'counter': c,
            'datasource_name': datasource_name,
            'datasource_caption': datasource_caption,
            'alias': field.alias,
            'field_calculation': field.calculation,
            'field_calculation_bk': field.calculation,
            'field_caption': field.caption,
            'field_datatype': field.datatype,
            'field_def_agg': field.default_aggregation,
            'field_desc': field.description,
            'field_hidden': field.hidden,
            'field_id': field.id,
            'field_is_nominal': field.is_nominal,
            'field_is_ordinal': field.is_ordinal,
            'field_is_quantitative': field.is_quantitative,
            'field_name': field.name,
            'field_role': field.role,
            'field_type': field.type,
            'field_worksheets': field.worksheets,
            'field_WHOLE': field
        }

        if field.calculation is not None:
            calcID.append(field.id)
            calcNames.append(field.name)

            f2 = field.id.replace(']', '').replace('[', '')
            calcID2.append(f2)

        c += 1
        collator.append(dict_temp)



def default_to_friendly_names2(formulaList,fieldToConvert, dictToUse):

    for i in formulaList:
        for tableauName, friendlyName in dictToUse.items():
            try:
                i[fieldToConvert] = (i[fieldToConvert]).replace(tableauName, friendlyName)
            except:
                a = 0
       
    return formulaList


def category_field_type(row):
    if row['datasource_name'] == 'Parameters':
        val = 'Parameters'
    elif row['field_calculation'] == None:
        val = 'Default_Field'
    else:
        val = 'Calculated_Field'
    return val

def compare_fields(row):
    if row['field_id'] == row['field_id2']:
        val = 0
    else:
        val = 1
    return val


calcDict = dict(zip(calcID, calcNames))
calcDict2 = dict(zip(calcID2, calcNames)) #raw fields without any []

collator = default_to_friendly_names2(collator,'field_calculation',calcDict2)

df_API_all = pd.DataFrame(collator)
df_API_all['field_type'] = df_API_all.apply(category_field_type, axis=1)

preference_list=['Parameters', 'Calculated_Field', 'Default_Field']
df_API_all["field_type"] = pd.Categorical(df_API_all["field_type"], categories=preference_list, ordered=True)

#get rid of duplicates for parameters, so only parameters from the explicit Parameters datasource are kept (as they are also listed again under the name of any other datasources)
df_API_all = df_API_all.sort_values(["field_id","field_type"]).drop_duplicates(["field_id", 'field_calculation']) 

df_API_all['field_id2'] = df_API_all['field_id'].str.replace(r'[\[\]]', '', regex=True)

df_API_all['comparison'] = df_API_all.apply(compare_fields, axis=1)
df_API_all = df_API_all[df_API_all['comparison'] == 1]

df_API_all = df_API_all.drop(['field_id2', 'comparison'], axis=1)
df_API_all.sort_values(['datasource_name', 'field_type', 'counter', 'field_name'])

df1 = df_API_all[[ 'field_name', 'field_datatype','field_type',  'field_calculation',   'field_id', 'datasource_caption', 'field_worksheets']].copy()

preference_list=[ 'Default_Field', 'Parameters', 'Calculated_Field']
df1["field_type"] = pd.Categorical(df1["field_type"], categories=preference_list, ordered=True)
df1 = df1.sort_values(['field_type'])

df1.columns = ['Field_Name', 'DataType', 'Type', 'Calculation', 'Field_ID', 'Datasource', 'Worksheets']

df1['Field_Name'] = df1['Field_Name'].str.replace(r'[\[\]]', '', regex=True)

# Clean the Worksheets column - convert list to comma-separated string without brackets
def format_worksheets(ws_list):
    if ws_list and len(ws_list) > 0:
        # Join worksheet names with comma and space
        return ', '.join(ws_list)
    return ''

df1['Worksheets'] = df1['Worksheets'].apply(format_worksheets)

# Add column to indicate if field is used in any worksheet
df1['Used_In_Report'] = df1['Worksheets'].apply(lambda x: 'Yes' if x else 'No')



# ## Generating an excel file from a df (so the excel rows/cols can be formatted), then turning the excel into a pdf

# Modify this part if you want to add more information/dfs to be saved as a separate sheet in excel
# Column widths optimized for: Field_Name(25), DataType(12), Type(18), Calculation(60), Field_ID(30), Datasource(25), Worksheets(40), Used_In_Report(15)

dfs_to_use = [{'excelSheetTitle': 'All fields extracted from DOC API', 'df_to_use':df1, 'mainColWidth':'' , 
               'normalColWidth': [25, 12, 18, 60, 30, 25, 40, 15], 'sheetName': 'GeneralDetails', 'footer': 'Data_1 (DOC API)', 'papersize':9, 'color': '#fff0b3'}                
             
             ]

#papersize: a3 = 8, a4 = 9

output_dir = os.path.join(os.getcwd(), 'outputs')
os.makedirs(output_dir, exist_ok=True)
path_excel_file_to_create = os.path.join(output_dir, f"{tableau_name_substring}_Calculations_table.xlsx")

exg.create_excel_from_dfs(dfs_to_use, path_excel_file_to_create)



# # Start of lineage diagram module

# Create abbreviated node IDs for the lineage diagram (AA, AB, AC, etc.)

def first_char_checker(cell_value):
    # Normalize field IDs by wrapping them with double underscores
    if cell_value[0] != '[':
        cell_value = '__' + cell_value + '__'
    else:
        cell_value = cell_value.replace('[', '__')
        cell_value = cell_value.replace(']', '__')
    return cell_value

# Define abc list to use during lineage diagram creation
abc = list(string.ascii_uppercase)
collated_abc = []

for i in abc:
    for j in abc:
        collated_abc.append(i+j)

# Map default fields to short abbreviations (AA, AB, etc.) for the diagram
# Filter to only include fields that are used in the report
def_fields = df1[(df1['Type'] == 'Default_Field') & (df1['Used_In_Report'] == 'Yes')]['Field_ID'].copy().apply(remove_sp_char_leave_undescore_square_brackets)

print(f"Default fields used in report: {len(def_fields)}")

abc_touse = collated_abc[0:len(def_fields)]

def_fields_final = pd.DataFrame(list(zip(def_fields.tolist(), abc_touse)))
def_fields_final['aa'] = def_fields_final.apply(lambda row: first_char_checker(row[0]), axis=1)

mapping_dict_friendly_names = dict(zip(def_fields_final[0].tolist(), abc_touse))
mapping_dict = dict(zip(def_fields_final['aa'].tolist(), abc_touse))

# Extract calculated fields and parameters, map them to abbreviated IDs (x___AA, x___AB, etc.)
# Filter to only include fields that are used in the report
created_calc = df_API_all[(df_API_all['field_type'] != 'Default_Field') & (df_API_all['field_worksheets'].apply(lambda x: x and len(x) > 0))]\
                [['field_name', 'field_id', 'field_calculation', 'field_calculation_bk']].copy()

print(f"Calculated fields used in report: {len(created_calc)}")

nlsi = ['x___' + i for i in collated_abc]
nlsi_to_use = nlsi[0:len(created_calc)]

created_calc['field_name'] = created_calc['field_name'].apply(remove_sp_char_leave_undescore_square_brackets)
created_calc['aa'] = created_calc.apply(lambda row: first_char_checker(row['field_id']), axis=1)
created_calc['field_calculation_bk'] = created_calc['field_calculation_bk'].str.replace(r'[\[\]]', '__', regex=True)

# Create mapping dictionary for calculated field abbreviations
calc_map_dict = dict(zip(created_calc['aa'].to_list(), nlsi_to_use))

# Add abbreviated IDs to calculated fields dataframe and sort
created_calc['shorthand_abc'] = created_calc['aa'].map(calc_map_dict)
created_calc.sort_values(by='shorthand_abc', inplace=True)

# Handle duplicate field names by adding numeric suffixes (e.g., Index, Index1, Index2)
def differentiate_duplicates(series):
    counts = series.groupby(series).cumcount() 
    return series + counts.astype(str).replace('0', '')

# differentiate field names that have duplicate values (eg. calc field Index appears twice in workbook, now it will be Index, Index1)
created_calc['field_name'] = differentiate_duplicates(created_calc['field_name'])

# Create final mapping of friendly field names to abbreviated IDs
calc_map_dict_friendly_names = dict(zip(created_calc['field_name'], created_calc['shorthand_abc']))

# Identify dependencies between fields by analyzing which fields are used in calculation formulas
def create_lineage_paths(df, field_type):
    c = 0
    t_collator = []

    for i in df['aa']:
        try:
            tlist = created_calc[created_calc['field_calculation_bk'].str.contains(i, regex=False) == True]['aa'].to_list()
        except:
            tlist = []

        if len(tlist) != 0:
            for x in tlist:
                newdict = {
                    'count': c,
                    'starting': i,
                    'ending': x,
                    'path_mermaid': i + " --> " + x
                }
                t_collator.append(newdict)
                c = c + 1
    
    return t_collator

# Find all dependencies for default fields
t_collator_def_fields = create_lineage_paths(def_fields_final, 'default_field')

# Find all dependencies for calculated fields
t_collator_calcs = create_lineage_paths(created_calc, 'calculation')

# Replace full field names in paths with abbreviated IDs to simplify the diagram
for default_field, mapping_letter in mapping_dict.items():
    for i in t_collator_def_fields:
        i['path_mermaid'] = i['path_mermaid'].replace(default_field, mapping_letter)

for default_field, mapping_letter in calc_map_dict.items():
    for i in t_collator_def_fields:
        i['path_mermaid'] = i['path_mermaid'].replace(default_field, mapping_letter)

for default_field, mapping_letter in mapping_dict.items():
    for i in t_collator_calcs:
        i['path_mermaid'] = i['path_mermaid'].replace(default_field, mapping_letter)

for default_field, mapping_letter in calc_map_dict.items():
    for i in t_collator_calcs:
        i['path_mermaid'] = i['path_mermaid'].replace(default_field, mapping_letter)

print(f"Processed {len(t_collator_def_fields)} default field paths")
print(f"Processed {len(t_collator_calcs)} calculated field paths")

# Build nodes (fields) and edges (dependencies) for the interactive lineage diagram
import json

nodes = []
edges = []
node_ids = set()

# Add default fields as nodes
for i, d in mapping_dict_friendly_names.items():
    if d not in node_ids:
        nodes.append({
            'id': d,
            'label': i,
            'group': 'default',
            'title': f'Default Field: {i}'
        })
        node_ids.add(d)

# Add calculated fields as nodes
for i, d in calc_map_dict_friendly_names.items():
    if d not in node_ids:
        calc_row = created_calc[created_calc['field_name'] == i]
        calc_formula = ''
        if not calc_row.empty and calc_row['field_calculation'].values[0]:
            calc_formula = str(calc_row['field_calculation'].values[0])
        
        nodes.append({
            'id': d,
            'label': i,
            'group': 'calculated',
            'title': f'{i}\n\nFormula:\n{calc_formula}'
        })
        node_ids.add(d)

# Build edges from the collators
for item in t_collator_def_fields:
    parts = item['path_mermaid'].split(' --> ')
    if len(parts) == 2:
        edges.append({
            'from': parts[0],
            'to': parts[1],
            'arrows': 'to'
        })

for item in t_collator_calcs:
    parts = item['path_mermaid'].split(' --> ')
    if len(parts) == 2:
        edges.append({
            'from': parts[0],
            'to': parts[1],
            'arrows': 'to'
        })

# Convert to JSON for JavaScript
nodes_json = json.dumps(nodes)
edges_json = json.dumps(edges)

print(f"Total nodes: {len(nodes)}")
print(f"Total edges: {len(edges)}")
print("Nodes and edges prepared for lineage diagram")

# Generate interactive HTML lineage diagram using Vis.js library
# Creates a hierarchical network visualization with tooltips showing formulas

html_base = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>""" + tableau_name_substring + """ Calculation Lineage</title>
    <script type="text/javascript" src="https://unpkg.com/vis-network/standalone/umd/vis-network.min.js"></script>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }
        h1 {
            color: #333;
            text-align: center;
        }
        #mynetwork {
            width: 100%;
            height: 800px;
            border: 1px solid #ddd;
            background-color: white;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .controls {
            text-align: center;
            margin: 20px 0;
            padding: 15px;
            background-color: white;
            border-radius: 5px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .controls button {
            margin: 0 5px;
            padding: 10px 20px;
            background-color: #4CAF50;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
        }
        .controls button:hover {
            background-color: #45a049;
        }
        .legend {
            margin-top: 20px;
            padding: 15px;
            background-color: white;
            border-radius: 5px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .legend-item {
            display: inline-block;
            margin-right: 20px;
        }
        .legend-color {
            display: inline-block;
            width: 20px;
            height: 20px;
            border-radius: 50%;
            margin-right: 5px;
            vertical-align: middle;
        }
    </style>
</head>
<body>
    <h1>""" + tableau_name_substring + """ Calculation Lineage</h1>
    
    <div class="controls">
        <button onclick="network.fit()">Fit to Screen</button>
        <button onclick="network.moveTo({scale: 1.0})">Reset Zoom</button>
    </div>
    
    <div id="mynetwork"></div>
    
    <div class="legend">
        <strong>Legend:</strong>
        <div class="legend-item">
            <span class="legend-color" style="background-color: #97C2FC;"></span>
            <span>Default Fields</span>
        </div>
        <div class="legend-item">
            <span class="legend-color" style="background-color: #FB7E81;"></span>
            <span>Calculated Fields</span>
        </div>
    </div>

    <script type="text/javascript">
        // Create nodes and edges data
        var nodes = new vis.DataSet(""" + nodes_json + """);
        var edges = new vis.DataSet(""" + edges_json + """);

        // Create network
        var container = document.getElementById('mynetwork');
        var data = {
            nodes: nodes,
            edges: edges
        };
        
        var options = {
            nodes: {
                shape: 'box',
                margin: 10,
                widthConstraint: {
                    maximum: 200
                },
                font: {
                    size: 14
                }
            },
            edges: {
                arrows: {
                    to: {
                        enabled: true,
                        scaleFactor: 0.5
                    }
                },
                smooth: {
                    type: 'cubicBezier',
                    forceDirection: 'horizontal'
                },
                color: {
                    color: '#848484',
                    highlight: '#2B7CE9'
                }
            },
            groups: {
                default: {
                    color: {
                        background: '#97C2FC',
                        border: '#2B7CE9',
                        highlight: {
                            background: '#D2E5FF',
                            border: '#2B7CE9'
                        }
                    }
                },
                calculated: {
                    color: {
                        background: '#FB7E81',
                        border: '#E92B36',
                        highlight: {
                            background: '#FFB5B8',
                            border: '#E92B36'
                        }
                    }
                }
            },
            layout: {
                hierarchical: {
                    enabled: true,
                    direction: 'LR',
                    sortMethod: 'directed',
                    levelSeparation: 200,
                    nodeSpacing: 150
                }
            },
            physics: {
                enabled: false
            },
            interaction: {
                hover: true,
                tooltipDelay: 100,
                navigationButtons: true,
                keyboard: true
            }
        };
        
        var network = new vis.Network(container, data, options);
        
        // Event listener for node clicks
        network.on("click", function(params) {
            if (params.nodes.length > 0) {
                var nodeId = params.nodes[0];
                var node = nodes.get(nodeId);
                console.log("Clicked node:", node);
            }
        });
        
        // Fit network when loaded
        network.once("stabilizationIterationsDone", function() {
            network.fit();
        });
    </script>
</body>
</html>
"""

print('\n ______________________________ START_OF_HTML ______________________________')
print(html_base[:500] + '...')
print('\n ______________________________ END_OF_HTML ______________________________')

# Output html string to a local file, then open it on the web browser

# Specify the file path
file_path = os.path.join('outputs', f'{tableau_name_substring}_lineage_diagram.html')

# Write the string to an HTML file with UTF-8 encoding
with open(file_path, 'w', encoding='utf-8') as file:
    file.write(html_base)

print("HTML content successfully written to {}".format(file_path))

# Open the HTML file in the default web browser
webbrowser.open('file://' + os.path.realpath(file_path))

print("\n✓ Interactive lineage diagram created and opened in browser")

