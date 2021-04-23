# Import Libraries
#import pyNastran
import json
import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import os
import io

from PIL import Image

from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as Img
from openpyxl.styles import Font

from matplotlib.colors import ListedColormap

from pyNastran.op2.op2 import OP2
from pyNastran.bdf.bdf import BDF
from pyNastran.op2.data_in_material_coord import data_in_material_coord


# Edit file 'op2_reader.py' in the pyNastran library to add a new optistruct version at line 5777
# if the read_op2 fails with the message "unknown version";
# already added versions b'OS2017.0' b'OS2019.1', b'OS2019.2'

def read_mapping(mapping_path):
    ''' Reading a json file
        Returning a dictionary with mapping in numpy format'''

    # Reading the element mapping dictionary from json file
    with open(mapping_path, 'r') as f:
        mapping = json.load(f)

    # Transforming the dictionary of list into a dictionary of numpy darrays
    elm_mapping={}
    for key in mapping.keys():
        elm_mapping[key] = np.array(mapping[key])

    return elm_mapping

def save_in_excel(workbook, ws_counter, tag, data_list, img=None):
    ''' workbook is the path/name of an already created excel file
        tag is an identification aboute the data
        data is a numpy object
        img is an image (optional)
    '''
    book = load_workbook(workbook)
        
    ws_name = 'Sheet_' + str(ws_counter)
    
    writer = pd.ExcelWriter(workbook, engine = 'openpyxl')
    writer.book = book
    
    # Write line in Index_of_results worksheet
    # esta parte hace dar error al excel; además no parece colocarlo en su sitio
    ws_index = book['Index_of_results']
    cell_pos = 'A' + str(4 + ws_counter)
    ws_index[cell_pos] = ws_name + ' ----> ' + tag
    ws_index[cell_pos].font = Font(size=12, color="000000FF")
    
    # Creating hyperlink between index and new sheet
    link = '#' + ws_name + '!A1'
    ws_index[cell_pos].hyperlink = link
    
    row_idx = 0
    for data in data_list:
        # Create a Pandas dataframe from the data
        df = pd.DataFrame(data)
        df_height = df.shape[0]
        df.to_excel(writer, sheet_name = ws_name, startrow = (row_idx), startcol=0, index=False, header=False)
        row_idx += df_height + 4

    if img != None:
        sheet = writer.book[ws_name]
        img=Img(img)
        cell_pos = 'A' + str(row_idx)
        sheet.add_image(img, cell_pos)
   
    writer.save()
    writer.close()
    
    print('Results saved in excel workbook:', workbook)
    
    return

def fig2img(fig):
    """Convert a Matplotlib figure to a PIL Image and return it"""
    buf = io.BytesIO()
    fig.savefig(buf)
    buf.seek(0)
    img = Image.open(buf)
    
    return img


# Definition of the "Model" class definition and its methods
class Model:
    def __init__(self, op2_path, mapping_path):
        self.op2_path = op2_path
        self.mapping_path = mapping_path
        self.op2 = OP2(debug=False)         # instantiate self.op2
        self.elem_to_idx = {}
        self.load_cases = []
        self.workbook = os.path.splitext(op2_path)[0] + '.xlsx'   # instantiate excel workbook name
        self.ws_counter = int(0)    # instantiate counter for excel 
        
        # Plantearse crear un directorio dedicado para almacenar resultados
        
        # Creating a new Excel workbook for saving results
        wb = Workbook()        
        ws = wb.active
        ws.sheet_view.showGridLines = False         # grid lines off
        ws.title = 'Index_of_results'               # change name of active worksheet
        ws['A1'] = 'Index of results extracted with OP_Map from OP2 file: ' + op2_path
        ws['A1'].font = Font(size = 16, bold = True)
        ws['A2'] = 'OP_Map by Félix R. López M., version: beta'
        ws['A2'].font = Font(size = 8, italic = True)
        
        wb.save(self.workbook)
        print('Created excel workbook:', self.workbook)

    def r_op2_eforces(self):
        self.op2.set_results(('force.ctria3_force','force.cquad4_force'))
        self.op2.read_op2(self.op2_path);

        # Model statistics (set short to False for more detail)
        print('Loaded ctria3 and cquad4 element forces from op2 file:', self.op2_path)

        # Getting forces for all subcases  --> diccionary with key the LC and values the force values
        cq4_force = self.op2.cquad4_force
        tr3_force = self.op2.ctria3_force

        # Creating a list with all load cases
        self.load_cases = [lc for lc in cq4_force.keys()]
        print('Load cases in the op2 file:',len(self.load_cases))

        # Creating a list with all the elements ID and type
        lc = self.load_cases[0]
        cq4_elements = cq4_force[lc].element
        tr3_elements = tr3_force[lc].element
        elements = np.concatenate((cq4_elements,tr3_elements), axis=0).tolist()

        # Creating a dictionary from element to index starting with 1
        # first index associated to dummy element -666
        self.elem_to_idx[-666]=0
        for idx, elm in enumerate(elements,1):
            self.elem_to_idx[elm] = idx

        return

    def r_op2_eforces_matcoord(self, bdf_path):
        
        # Reading bdf file
        bdf = BDF(debug=False)  # instantiate bdf 
        bdf.read_bdf(bdf_path)
        
        # Creating a new op2 file with 2D results in material coordinates
        self.op2.set_results(('force.ctria3_force','force.cquad4_force'))
        self.op2.read_op2(self.op2_path);
        self.op2 = data_in_material_coord(bdf, self.op2)
        
        # Model statistics (set short to False for more detail)
        print('Loaded ctria3 and cquad4 element forces from op2 file:', self.op2_path)

        # Getting forces for all subcases  --> diccionary with key the LC and values the force values
        cq4_force = self.op2.cquad4_force
        tr3_force = self.op2.ctria3_force

        # Creating a list with all load cases
        self.load_cases = [lc for lc in cq4_force.keys()]
        print('Load cases in the op2 file:',len(self.load_cases))

        # Creating a list with all the elements ID and type
        lc = self.load_cases[0]
        cq4_elements = cq4_force[lc].element
        tr3_elements = tr3_force[lc].element
        elements = np.concatenate((cq4_elements,tr3_elements), axis=0).tolist()

        # Creating a dictionary from element to index starting with 1
        # first index associated to dummy element -666
        self.elem_to_idx[-666]=0
        for idx, elm in enumerate(elements,1):
            self.elem_to_idx[elm] = idx
            
        return    
    
    def list_lc(self, excel=False):
        ''' This function list the load cases of the op2 file considered'''
        print(self.load_cases)
        
        if excel:
            self.ws_counter += 1
            tag = 'List of load cases in the OP2 file'
            save_in_excel(self.workbook, self.ws_counter, tag, [self.load_cases])

        return

    def change_mapping(self, new_mapping_path):
        ''' This function change the mapping file for a new one'''
        self.mapping_path = new_mapping_path
        return

    def plot_max_eforces(self, component, value_field, excel = False):
        # Reading mapping from json file
        mapping = read_mapping(self.mapping_path)
        # Extracting elements mapping for selected component
        elm_mapping = mapping[component]
        # Extracting fishtail mapping dimensions
        n_dim = elm_mapping.shape[0]
        m_dim = elm_mapping.shape[1]

        # Matrix flattening and turning into a list
        elm_mapping_flt = elm_mapping.reshape(-1,).tolist()

        elm_mapping_flt_mask = [1 if elm != -666 else np.nan for elm in elm_mapping_flt]
        elm_mapping_mask = np.array(elm_mapping_flt_mask).reshape(n_dim, m_dim)

        # From element_mapping (flatten) to index_mapping
        # Note that element with id -666 is turn into index 0 (and later in nan when getting the results)
        idx_mapping = [self.elem_to_idx[elm] for elm in elm_mapping_flt]

        # Nunber of load load_cases
        d_dim = len(self.load_cases)

        # Initializing variable output_flt
        output_flt = np.zeros((d_dim, len(idx_mapping)))

        for idx, lc in enumerate(self.load_cases):
            # Accessing individual element forces for a given lC
            cq4_forces_lc = self.op2.cquad4_force[lc].data
            tr3_forces_lc = self.op2.ctria3_force[lc].data

            # Concatenating data from cquads and trias
            forces_lc = np.concatenate((cq4_forces_lc,tr3_forces_lc), axis=1)

            # Removing first dimension
            forces_lc = forces_lc.reshape(forces_lc.shape[1], forces_lc.shape[2])

            # Adding a first line of nan values associated to element index = 0 (dummy element-666)
            nones = np.repeat(np.nan, forces_lc.shape[1], axis=0).reshape(1, forces_lc.shape[1])
            forces_lc = np.concatenate((nones, forces_lc), axis=0)

            # Getting results for the idx_mapping
            output_flt[idx,:] = forces_lc[idx_mapping, (value_field-1)]

        # Deflattening
        output = output_flt.reshape(d_dim, n_dim, m_dim)

        # Maximum values
        output_max = np.max(output, axis=0)
        # Index in axis=0 of the maxim values (=lc index)
        output_max_lc_idx = np.argmax(output, axis=0)

        # From lc_index to lc id
        output_max_lc_flt = [self.load_cases[idx] for idx in output_max_lc_idx.reshape(-1).tolist()]
        output_max_lc = np.array(output_max_lc_flt).reshape(n_dim, m_dim)
        # Removing values from -666 elements
        output_max_lc = np.multiply(output_max_lc, elm_mapping_mask)

        # Plotting heatmap (fishtail shape)
        plt.figure(figsize=(40,20))
        plt.subplot(2, 1, 1)
        plot_1 = sns.heatmap(output_max, annot=True, fmt='.1f', annot_kws={"size": 20},
                             linewidths=2, cmap='coolwarm');
        plt.subplot(2, 1, 2)
        plot_2 = sns.heatmap(output_max_lc, annot=True, fmt='.0f', annot_kws={"size": 20},
                             cmap=ListedColormap(['whitesmoke']),
                             linewidths=1, linecolor='White');
        
        plot_img = fig2img(plot_1.get_figure())
        
        if excel:
            self.ws_counter += 1
            tag = 'Component: {}. Maximum element forces in dimension {} and corresponding load cases'.format(component, value_field)
            save_in_excel(self.workbook, self.ws_counter, tag, [output_max, output_max_lc], img = plot_img)
            
        return

    def plot_eforces(self, lc, component, value_field, excel = False):
        ''' This function plots the op2 values for the lc, component and value
            field passed in the function.
            Inputs:
                lc: load case to plot
                component: name of the component to plot acc. to mapping file
                value_field: force compoment acc. to F06 order
            Output: seaborn heatmap'''
        # Reading mapping from json file
        mapping = read_mapping(self.mapping_path)
        # Extracting elements mapping for selected component
        elm_mapping = mapping[component]
        # Extracting fishtail mapping dimensions
        n_dim = elm_mapping.shape[0]
        m_dim = elm_mapping.shape[1]

        # Matrix flattening and turning into a list
        elm_mapping_flt = elm_mapping.reshape(-1,).tolist()

        # From element_mapping (flatten) to index_mapping
        # Note that element with id -666 is turn into index 0 (and later in nan when getting the results)
        idx_mapping = [self.elem_to_idx[elm] for elm in elm_mapping_flt]
        # Accessing individual element forces for a given lC
        cq4_forces_lc = self.op2.cquad4_force[lc].data
        tr3_forces_lc = self.op2.ctria3_force[lc].data

        # Concatenating data from cquads and trias
        forces_lc = np.concatenate((cq4_forces_lc,tr3_forces_lc), axis=1)

        # Removing first dimension
        forces_lc = forces_lc.reshape(forces_lc.shape[1], forces_lc.shape[2])

        # Adding a first line of nan values associated to element index = 0 (dummy element-666)
        nones = np.repeat(np.nan, forces_lc.shape[1], axis=0).reshape(1, forces_lc.shape[1])
        forces_lc = np.concatenate((nones, forces_lc), axis=0)

        # Getting results for the idx_mapping
        output_flt = forces_lc[idx_mapping, (value_field-1)]

        # Deflattening
        output = output_flt.reshape(n_dim, m_dim)

        # Plotting heatmap (fishtail shape)
        plt.figure(figsize=(40,10))
        plot = sns.heatmap(output, annot=True, fmt='.1f',  annot_kws={"size": 20},
                           linewidths=2, cmap='coolwarm');
        plot_img = fig2img(plot.get_figure())
        
        if excel:
            self.ws_counter += 1
            tag = 'Component: {}. Element forces in dimension {} and for load case '.format(component, value_field) + str(lc)
            save_in_excel(self.workbook, self.ws_counter, tag, [output], img = plot_img)
            
        return

    def plot_component_mapping(self, component, excel = False):
        # Reading mapping from json file
        mapping = read_mapping(self.mapping_path)
        # Extracting elements mapping for selected component
        elm_mapping = mapping[component]
        # Extracting fishtail mapping dimensions
        n_dim = elm_mapping.shape[0]
        m_dim = elm_mapping.shape[1]

        # Matrix flattening and turning into a list
        elm_mapping_flt = elm_mapping.reshape(-1,).tolist()

        # From element_mapping (flatten) to index_mapping
        # Note that element with id -666 is turn into index 0 (and later in nan when getting the results)
        elm_mapping_flt = [elm if elm != -666 else np.nan for elm in elm_mapping_flt]

        # Deflattening
        elm_mapping = np.array(elm_mapping_flt).reshape(n_dim, m_dim)

        # Plotting heatmap (fishtail shape)
        plt.figure(figsize=(40,10))
        plot = sns.heatmap(elm_mapping, annot=True, fmt='.0f', annot_kws={"size": 20},
                           cmap=ListedColormap(['whitesmoke']),
                           linewidths=1, linecolor='white');
        plt.title('Element mapping of component: {}'.format(component), fontsize = 20)
        plt.xlabel('X-label', fontsize = 15)
        plt.ylabel('Y-label', fontsize = 15)
        
        plot_img = fig2img(plot.get_figure())
        
        if excel:
            self.ws_counter += 1
            tag = 'Component: {}. Mapping of elements'.format(component)
            save_in_excel(self.workbook, self.ws_counter, tag, [elm_mapping], img = plot_img)

        return
