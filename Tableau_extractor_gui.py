import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import pandas as pd
from tableaudocumentapi import Workbook
import webbrowser
import Excelcreator as exg

class TableauExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Tableau Calculation & Lineage Extractor")
        self.root.geometry("450x350")
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Input section
        ttk.Label(main_frame, text="Select Tableau Workbook (.twb/.twbx):", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        # Input file frame
        input_frame = ttk.Frame(main_frame)
        input_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        self.input_path = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.input_path, width=50).grid(row=0, column=0, padx=5)
        ttk.Button(input_frame, text="Browse", command=self.browse_input).grid(row=0, column=1, padx=5)

        # Output directory section
        ttk.Label(main_frame, text="Select Output Directory:", font=("Arial", 10, "bold")).grid(row=2, column=0, sticky=tk.W, pady=5)
        
        # Output directory frame
        output_dir_frame = ttk.Frame(main_frame)
        output_dir_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=5)
        
        self.output_dir = tk.StringVar(value=os.path.join(os.getcwd(), "outputs"))
        ttk.Entry(output_dir_frame, textvariable=self.output_dir, width=50).grid(row=0, column=0, padx=5)
        ttk.Button(output_dir_frame, text="Browse", command=self.browse_output_dir).grid(row=0, column=1, padx=5)
        
        # Output options section
        ttk.Label(main_frame, text="Output Options:", font=("Arial", 10, "bold")).grid(row=4, column=0, sticky=tk.W, pady=5)
        
        # Checkboxes
        self.excel_var = tk.BooleanVar(value=True)
        self.mermaid_var = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(main_frame, text="Generate Excel", variable=self.excel_var).grid(row=5, column=0, sticky=tk.W)
        ttk.Checkbutton(main_frame, text="Generate Lineage Diagram", variable=self.mermaid_var).grid(row=6, column=0, sticky=tk.W)
        
        # Status
        self.status_var = tk.StringVar()
        ttk.Label(main_frame, textvariable=self.status_var, wraplength=500).grid(row=7, column=0, sticky=tk.W, pady=10)
        
        # Process button
        ttk.Button(main_frame, text="Process Workbook", command=self.process_workbook).grid(row=8, column=0, pady=20)

    def browse_input(self):
        filename = filedialog.askopenfilename(
            title="Select Tableau Workbook",
            filetypes=[("Tableau Workbook (.twb/.twbx)", "*.twb *.twbx"), ("All files", "*")]
        )
        if filename:
            self.input_path.set(filename)
            
    def browse_output_dir(self):
        directory = filedialog.askdirectory(
            title="Select Output Directory"
        )
        if directory:
            self.output_dir.set(directory)

    def process_workbook(self):
        if not self.input_path.get():
            messagebox.showerror("Error", "Please select a Tableau workbook file.")
            return
        
        try:
            self.status_var.set("Processing workbook...")
            self.root.update()
            
            # Process the workbook
            input_file = self.input_path.get()
            workbook = None
            tmpdir = None
            try:
                # Try opening directly (works for .twb)
                workbook = Workbook(input_file)
            except Exception:
                # If it's a packaged .twbx, try to extract contained .twb and open that
                if input_file.lower().endswith('.twbx'):
                    import zipfile, tempfile, shutil
                    z = zipfile.ZipFile(input_file, 'r')
                    twb_name = next((n for n in z.namelist() if n.lower().endswith('.twb')), None)
                    if twb_name is None:
                        z.close()
                        raise RuntimeError(f"No .twb found inside packaged workbook: {input_file}")
                    tmpdir = tempfile.mkdtemp(prefix='twbx_extract_')
                    try:
                        z.extract(twb_name, tmpdir)
                        extracted_path = os.path.join(tmpdir, twb_name)
                        workbook = Workbook(extracted_path)
                    finally:
                        z.close()
                else:
                    raise
            
            # Extract filename for output (strip extension)
            tableau_name = os.path.basename(input_file)
            tableau_name_substring = os.path.splitext(tableau_name)[0][:30]
            
            # Get output directory and create if it doesn't exist
            output_dir = self.output_dir.get()
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # Process calculations and create outputs
            df, df_API_all = self.extract_calculations(workbook)
            
            if self.excel_var.get():
                dfs_to_use = [{
                    'excelSheetTitle': 'All fields extracted from DOC API',
                    'df_to_use': df,
                    'mainColWidth': '',
                    'normalColWidth': [10, 15, 15, 50, 20, 25, 30, 12],
                    'sheetName': 'GeneralDetails',
                    'footer': 'Data_1 (DOC API)',
                    'papersize': 9,
                    'color': '#fff0b3'
                }]
                
                # Create file paths in the selected output directory
                excel_filename = f"{tableau_name_substring}_Calculations_table.xlsx"
                path_excel = os.path.join(output_dir, excel_filename)
                # Create the Excel file if requested
                if self.excel_var.get():
                    exg.create_excel_from_dfs(dfs_to_use, path_excel)
            
            if self.mermaid_var.get():
                self.generate_mermaid_diagram(df, df_API_all, tableau_name_substring, output_dir)
            
            self.status_var.set("Processing complete! Check the outputs folder for generated files.")
            messagebox.showinfo("Success", "Processing complete! Files have been generated in the outputs folder.")
            
        except Exception as e:
            self.status_var.set(f"Error: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
        finally:
            # Clean up any temporary extraction directory
            try:
                if 'tmpdir' in locals() and tmpdir is not None:
                    import shutil
                    shutil.rmtree(tmpdir)
            except Exception:
                pass

    def extract_calculations(self, workbook):
        def default_to_friendly_names2(formulaList, fieldToConvert, dictToUse):
            # Replace field IDs in formulas with friendly names for readability
            for i in formulaList:
                for tableauName, friendlyName in dictToUse.items():
                    try:
                        if i[fieldToConvert] is not None:
                            i[fieldToConvert] = i[fieldToConvert].replace(tableauName, friendlyName)
                    except:
                        pass
            return formulaList
        
        try:
            collator = []
            calcID = []
            calcNames = []
            c = 0
            calcID = []
            calcNames = []
            calcID2 = []
            
            for datasource in workbook.datasources:
                datasource_name = datasource.name
                datasource_caption = datasource.caption if datasource.caption else datasource_name
                
                for field in datasource.fields.values():
                    try:
                        field_id = field.id if field.id else f"[{field.name}]"
                        
                        # Store worksheets as a list (will be formatted later)
                        worksheet_names = field.worksheets if hasattr(field, 'worksheets') and field.worksheets else []
                        
                        dict_temp = {
                            'counter': c,
                            'datasource_name': datasource_name,
                            'datasource_caption': datasource_caption,
                            'field_calculation': field.calculation,
                            'field_calculation_bk': field.calculation,
                            'field_datatype': getattr(field, 'datatype', 'unknown'),
                            'field_id': field_id,
                            'field_name': field.name,
                            'field_worksheets': worksheet_names
                        }
                        
                        if field.calculation is not None:
                            calcID.append(field_id)
                            calcNames.append(field.name)
                            calcID2.append(field_id.replace(']', '').replace('[', ''))
                    
                        c += 1
                        collator.append(dict_temp)
                    except Exception as field_error:
                        self.status_var.set(f"Warning: Skipping field {field.name} due to error: {str(field_error)}")
                        continue
            
            # Create mapping dictionaries for field ID to field name replacement
            calcDict = dict(zip(calcID, calcNames))
            calcDict2 = dict(zip(calcID2, calcNames))  # raw fields without any []
            
            # Replace field IDs with friendly names in calculations
            collator = default_to_friendly_names2(collator, 'field_calculation', calcDict2)
            
            df_API_all = pd.DataFrame(collator)
            
            def category_field_type(row):
                if row['datasource_name'] == 'Parameters':
                    return 'Parameters'
                elif row['field_calculation'] is None:
                    return 'Default_Field'
                else:
                    return 'Calculated_Field'
            
            # Categorize fields
            df_API_all['field_type'] = df_API_all.apply(category_field_type, axis=1)
            
            # Sort by preference
            preference_list = ['Parameters', 'Calculated_Field', 'Default_Field']
            df_API_all["field_type"] = pd.Categorical(df_API_all["field_type"], categories=preference_list, ordered=True)
            
            # Remove duplicates for parameters
            df_API_all = df_API_all.sort_values(["field_id", "field_type"]).drop_duplicates(["field_id", 'field_calculation'])
            
            # Create final DataFrame
            df1 = pd.DataFrame({
                'Field_Name': df_API_all['field_name'],
                'DataType': df_API_all['field_datatype'],
                'Type': df_API_all['field_type'],
                'Calculation': df_API_all['field_calculation'],
                'Field_ID': df_API_all['field_id'],
                'Datasource': df_API_all['datasource_caption'],
                'Worksheets': df_API_all['field_worksheets']
            })
            
            # Clean up field names
            df1['Field_Name'] = df1['Field_Name'].str.replace(r'[\[\]]', '', regex=True)
            
            # Clean the Worksheets column - convert list to comma-separated string without brackets
            def format_worksheets(ws_list):
                if ws_list and isinstance(ws_list, list) and len(ws_list) > 0:
                    # Join worksheet names with comma and space
                    return ', '.join(ws_list)
                return ''
            
            df1['Worksheets'] = df1['Worksheets'].apply(format_worksheets)
            
            # Add column to indicate if field is used in any worksheet
            df1['Used_In_Report'] = df1['Worksheets'].apply(lambda x: 'Yes' if x else 'No')
            
            return df1, df_API_all
            
        except Exception as e:
            raise Exception(f"Error extracting calculations: {str(e)}")

    def generate_mermaid_diagram(self, df, df_API_all, tableau_name_substring, output_dir):
        import string
        import re
        import json

        def first_char_checker(cell_value):
            # Normalize field IDs by wrapping them with double underscores
            if cell_value[0] != '[':
                cell_value = '__' + cell_value + '__'
            else:
                cell_value = cell_value.replace('[', '__')
                cell_value = cell_value.replace(']', '__')
            return cell_value

        def remove_sp_char_leave_undescore_square_brackets(string_to_convert):
            filtered_string = re.sub(r'[^a-zA-Z0-9\s._\[\]]', '', string_to_convert).replace(' ', "_")
            return filtered_string

        try:
            # Create ABC list for node IDs
            abc = list(string.ascii_uppercase)
            collated_abc = []
            for i in abc:
                for j in abc:
                    collated_abc.append(i+j)
            
            # Process default fields - filter to only include fields used in the report
            def_fields_df = df[(df['Type'] == 'Default_Field') & (df['Used_In_Report'] == 'Yes')][['Field_ID', 'Field_Name']].copy()
            def_fields = def_fields_df['Field_ID'].apply(remove_sp_char_leave_undescore_square_brackets)
            def_fields_original_names = def_fields_df['Field_Name'].tolist()  # Keep original names for display
            abc_touse = collated_abc[0:len(def_fields)]
            
            def_fields_final = pd.DataFrame(list(zip(def_fields.tolist(), abc_touse, def_fields_original_names)), columns=['cleaned', 'abbrev', 'original'])
            def_fields_final['aa'] = def_fields_final['cleaned'].apply(lambda x: first_char_checker(x))
            
            mapping_dict_friendly_names = dict(zip(def_fields_final['cleaned'].tolist(), abc_touse))
            mapping_dict_original_names = dict(zip(abc_touse, def_fields_final['original'].tolist()))  # Map abbrev to original names
            mapping_dict = dict(zip(def_fields_final['aa'].tolist(), abc_touse))
            
            # Process calculated fields - filter to only include fields used in the report
            created_calc = df_API_all[(df_API_all['field_type'] != 'Default_Field') & (df_API_all['field_worksheets'].apply(lambda x: x and len(x) > 0))][['field_name', 'field_id', 'field_calculation', 'field_calculation_bk']].copy()
            
            nlsi = ['x___' + i for i in collated_abc]
            nlsi_to_use = nlsi[0:len(created_calc)]
            
            # Store original field names BEFORE cleaning
            created_calc['field_name_original'] = created_calc['field_name'].copy()
            
            created_calc['field_name'] = created_calc['field_name'].apply(remove_sp_char_leave_undescore_square_brackets)
            created_calc['aa'] = created_calc.apply(lambda row: first_char_checker(row['field_id']), axis=1)
            created_calc['field_calculation_bk'] = created_calc['field_calculation_bk'].str.replace(r'[\[\]]', '__', regex=True)
            
            calc_map_dict = dict(zip(created_calc['aa'].to_list(), nlsi_to_use))
            
            created_calc['shorthand_abc'] = created_calc['aa'].map(calc_map_dict)
            created_calc.sort_values(by='shorthand_abc', inplace=True)
            
            # Handle duplicate field names
            def differentiate_duplicates(series):
                counts = series.groupby(series).cumcount()
                return series + counts.astype(str).replace('0', '')
            
            created_calc['field_name'] = differentiate_duplicates(created_calc['field_name'])
            calc_map_dict_friendly_names = dict(zip(created_calc['field_name'], created_calc['shorthand_abc']))
            calc_map_dict_original_names = dict(zip(created_calc['shorthand_abc'], created_calc['field_name_original']))  # Map abbrev to original names
            
            # Create lineage paths
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
                            t_collator.append({
                                'count': c,
                                'starting': i,
                                'ending': x,
                                'path_mermaid': i + " --> " + x
                            })
                            c += 1
                return t_collator
            
            t_collator_def_fields = create_lineage_paths(def_fields_final, 'default_field')
            t_collator_calcs = create_lineage_paths(created_calc, 'calculation')
            
            # Replace field names with abbreviated IDs
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
            
            # Build nodes and edges for Vis.js
            nodes = []
            edges = []
            node_ids = set()
            
            # Add default fields as nodes (use original names for labels)
            for abbrev, original_name in mapping_dict_original_names.items():
                if abbrev not in node_ids:
                    nodes.append({
                        'id': abbrev,
                        'label': original_name,
                        'group': 'default',
                        'title': f'Default Field: {original_name}'
                    })
                    node_ids.add(abbrev)
            
            # Add calculated fields as nodes (use original names for labels)
            for abbrev, original_name in calc_map_dict_original_names.items():
                if abbrev not in node_ids:
                    calc_row = created_calc[created_calc['shorthand_abc'] == abbrev]
                    calc_formula = ''
                    if not calc_row.empty and calc_row['field_calculation'].values[0]:
                        calc_formula = str(calc_row['field_calculation'].values[0])
                    
                    nodes.append({
                        'id': abbrev,
                        'label': original_name,
                        'group': 'calculated',
                        'title': f'{original_name}\n\nFormula:\n{calc_formula}'
                    })
                    node_ids.add(abbrev)
            
            # Build edges
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
            
            # Convert to JSON
            nodes_json = json.dumps(nodes)
            edges_json = json.dumps(edges)
            
            # Generate interactive HTML lineage diagram using Vis.js
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
            
            # Save to file
            lineage_filename = f"{tableau_name_substring}_lineage_diagram.html"
            file_path = os.path.join(output_dir, lineage_filename)
            
            with open(file_path, 'w', encoding='utf-8') as file:
                file.write(html_base)
            
            # Open the diagram in browser
            webbrowser.open('file://' + os.path.realpath(file_path))
            
        except Exception as e:
            raise Exception(f"Error generating lineage diagram: {str(e)}")

def main():
    root = tk.Tk()
    app = TableauExtractorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()