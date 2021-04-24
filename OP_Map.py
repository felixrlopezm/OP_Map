# OP_Map library: auxiliary functions, model class and its methods
# by Félix Ramón López Martínez
# v0.6

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


########################################################################
# AUXILIARY FUNCTIONS
########################################################################

def read_mapping(mapping_path):
    ''' Reading a json file
        Returning a dictionary with three elements per key:
        + element mapping in numpy format
        + labels for x axis of the mapping
        + labels for the y axis of the mapping
        '''

    # Reading the element mapping dictionary from json file
    with open(mapping_path, 'r') as f:
        mapping = json.load(f)

    # Transforming mapping into numpy darrays
    elm_mapping={}
    for key in mapping.keys():
        elm_mapping[key] = {'mapping': np.array(mapping[key]['mapping']),
                            'x_labels': mapping[key]['x_labels'],
                            'y_labels': mapping[key]['y_labels']}

    return elm_mapping

def save_in_excel(workbook, ws_counter, tag, data_list, img=None):
    ''' + workbook is the path/name of an already created excel file
        + tag is an identification aboute the data
        + data is a numpy object
        + img is an image (optional)
    '''
    # Load excel workbook and create the writer with openpyxl
    book = load_workbook(workbook)
    writer = pd.ExcelWriter(workbook, engine = 'openpyxl')
    writer.book = book
    
    # Name for a new workbook sheet with a sequencial id
    ws_name = 'Sheet_' + str(ws_counter)
    
    # Write a line in the Index_of_results sheet of the workbook
    ws_index = book['Index_of_results']
    cell_pos = 'A' + str(4 + ws_counter)
    ws_index[cell_pos] = ws_name + ' ----> ' + tag
    ws_index[cell_pos].font = Font(size=12, color="000000FF")
    
    # Hyperlink from the new line in the index to its sheet
    link = '#' + ws_name + '!A1'
    ws_index[cell_pos].hyperlink = link
    
    # Writing the data in the corresponding new sheet
    row_idx = 0
    for data in data_list:
        df = pd.DataFrame(data)    # Create a Pandas dataframe from the data
        df_height = df.shape[0]
        df.to_excel(writer, sheet_name = ws_name, startrow = (row_idx),
                    startcol=0, index=False, header=False)
        row_idx += df_height + 4
        
    # Inserting the image in the new sheet
    if img != None:
        sheet = writer.book[ws_name]
        img=Img(img)
        cell_pos = 'A' + str(row_idx)
        sheet.add_image(img, cell_pos)
   
    # Save and close the workbook
    writer.save()
    writer.close()
    
    # Output message
    print('Results saved in excel workbook:', workbook)
    
    return


def fig2img(fig):
    """Convert a Matplotlib figure to a PIL Image and return it"""
    buf = io.BytesIO()
    fig.savefig(buf)
    buf.seek(0)
    img = Image.open(buf)
    
    return img


def mapping_extraction(component, mapping_path):
    # Reading mapping from json file
    mapping = read_mapping(mapping_path)
    
    # Extracting elements mapping for selected component
    elm_mapping = mapping[component]['mapping']
    x_labels = mapping[component]['x_labels']
    y_labels = mapping[component]['y_labels']
    
    # Extracting fishtail mapping dimensions
    n_dim = elm_mapping.shape[0]
    m_dim = elm_mapping.shape[1]
    # Matrix flattening and turning into a list
    elm_mapping_flt = elm_mapping.reshape(-1,).tolist()
        
    return elm_mapping_flt, x_labels, y_labels, n_dim, m_dim    


########################################################################
# MODEL CLASSS AND METHODS
########################################################################

class Model:
    def __init__(self, op2_path, mapping_path):
        self.op2_path = op2_path
        self.mapping_path = mapping_path
        self.op2 = OP2(debug=False)         # instantiate self.op2
        self.elem_to_idx = {}
        self.load_cases = []
        self.workbook = os.path.splitext(op2_path)[0] + '.xlsx'   # instantiate excel workbook name
        self.ws_counter = int(0)                             # instantiate counter for excel sheets
        
        # Creating a new Excel workbook and the Index_of_results sheet
        wb = Workbook()        
        ws = wb.active
        ws.sheet_view.showGridLines = False         # grid lines off
        ws.title = 'Index_of_results'               # change name of active worksheet
        ws['A1'] = 'Index of results extracted with OP_Map from OP2 file: ' + op2_path
        ws['A1'].font = Font(size = 16, bold = True)
        ws['A2'] = 'OP_Map by Félix R. López M., version: beta'
        ws['A2'].font = Font(size = 8, italic = True)
        
        # Saving the excel workbook
        wb.save(self.workbook)
        
        # Output message
        print('Created excel workbook:', self.workbook)

        
    def r_op2_eforces(self):
        ''' method for reading the element forces from the OP2 file
        '''
        
        self.op2.set_results(('force.ctria3_force','force.cquad4_force'))
        self.op2.read_op2(self.op2_path);

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
            
        # Output message
        print('Loaded ctria3 and cquad4 element forces in element coordinates from op2 file:', self.op2_path)

        return

    
    def r_op2_eforces_matcoord(self, bdf_path):
        ''' Method for reading the element forces contained in the OP2 file in material coordinates
            bdf_path is the path to the bdf file that corresponds to the OP2 file
        '''
        # Reading bdf file
        bdf = BDF(debug=False)  # instantiate bdf 
        bdf.read_bdf(bdf_path)
        
        # Creating a new op2 file with 2D results in material coordinates
        self.op2.set_results(('force.ctria3_force','force.cquad4_force'))
        self.op2.read_op2(self.op2_path);
        self.op2 = data_in_material_coord(bdf, self.op2)

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
        
        # Output message
        print('Loaded ctria3 and cquad4 element forces in material coordinates from op2 file:', self.op2_path)
        
        return    
 

    def list_lc(self, excel=False):
        ''' Method for listing the load cases in the OP2 file '''
        print(self.load_cases)
        
        if excel:
            self.ws_counter += 1
            tag = 'List of load cases in the OP2 file'
            save_in_excel(self.workbook, self.ws_counter, tag, [self.load_cases])

        return

    
    def change_mapping(self, new_mapping_path):
        ''' This method change the mapping file for a new one'''
        self.mapping_path = new_mapping_path
        return
 

    def plot_env_eforces(self, component, env_type, value_field, excel = False):
        ''' This methods plots the MAX, MIN or MAXABS element forces for the
            component and value field passed in.
            Inputs:
                component: name of the component to plot acc. to mapping file
                env_type: MAX, MIN or MAXABS according to the desired envelope
                value_field: force compoment acc. to F06 order
                excel: if True, then results will be stored in the excel workbook
            Output:
                plot of the results (seaborn heatmap image)
                and results saved in the excel file if requested
        '''
        # Mapping extraction for given component
        elm_mapping_flt, x_labels, y_labels, n_dim, m_dim, = mapping_extraction(component, self.mapping_path)
        
        # Mask creation for filtering results before plotting: 0 for -666 elements and 1 for all the others
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
        
        # Envelope options
        if env_type == 'MAX':
            output_env = np.max(output, axis=0)               # Maximum values
            output_env_lc_idx = np.argmax(output, axis=0)     # Index in axis=0 of maximum values (=lc index)
        elif env_type == 'MIN':
            output_env = np.min(output, axis=0)               # Minimum values
            output_env_lc_idx = np.argmin(output, axis=0)     # Index in axis=0 of minimum values (=lc index)
        elif env_type == 'MAXABS':
            output_env = np.max(np.absolute(output), axis=0)  # Maximum absolute values
            output_env_lc_idx = np.argmax(np.absolute(output), axis=0)    # Index in axis=0 of max abs values (=lc index)
        else:
            print('Incorrect envelope type selection. Choose MAX, MIN or MAXABS')
            
        # From lc_index to lc id
        output_env_lc_flt = [self.load_cases[idx] for idx in output_env_lc_idx.reshape(-1).tolist()]
        output_env_lc = np.array(output_env_lc_flt).reshape(n_dim, m_dim)
        # Removing values from -666 elements
        output_env_lc = np.multiply(output_env_lc, elm_mapping_mask)

        # Plotting heatmap (fishtail shape)
        plt.figure(figsize=(40,20))
        plt.subplot(2, 1, 1)
        plot_1 = sns.heatmap(output_env, annot=True, fmt='.1f', annot_kws={"size": 20},
                             linewidths=2, cmap='coolwarm',
                             xticklabels=x_labels, yticklabels=y_labels);
        plt.subplot(2, 1, 2)
        plot_2 = sns.heatmap(output_env_lc, annot=True, fmt='.0f', annot_kws={"size": 20},
                             cmap=ListedColormap(['whitesmoke']),
                             linewidths=1, linecolor='White',
                             xticklabels=x_labels, yticklabels=y_labels);
        
        plot_img = fig2img(plot_1.get_figure())
        
        # Saving results in the excel workbook if required
        if excel:
            self.ws_counter += 1
            tag = 'Component: {}. {} element forces in dimension {} and corresponding load cases'.format(
                component, env_type, value_field)
            save_in_excel(self.workbook, self.ws_counter, tag, [output_env, output_env_lc], img = plot_img)
            
        return

    
    def plot_eforces(self, lc, component, value_field, excel = False):
        ''' This methods plots element forces for the load case,
            component and value field passed in.
            Inputs:
                lc: load case to plot
                component: name of the component to plot acc. to mapping file
                value_field: force compoment acc. to F06 order
                excel: if True, then results will be stored in the excel workbook
            Output:
                plot of the results (seaborn heatmap image)
                and results saved in the excel file if requested
        '''   
        # Mapping extraction for given component
        elm_mapping_flt, x_labels, y_labels, n_dim, m_dim, = mapping_extraction(component, self.mapping_path)
        
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
        
        # Saving results in the excel workbook if required
        if excel:
            self.ws_counter += 1
            tag = 'Component: {}. Element forces in dimension {} and for load case '.format(
                component, value_field) + str(lc)
            save_in_excel(self.workbook, self.ws_counter, tag, [output], img = plot_img)
            
        return

    
    def plot_component_mapping(self, component, excel = False):
        ''' Method for plotting the mapping of a given component
            Input:
                component
                excel: if True, then results will be stored in the excel workbook
            Output:
                plot of the results (seaborn heatmap image)
                and results saved in the excel file if requested
        '''
        # Mapping extraction for given component
        elm_mapping_flt, x_labels, y_labels, n_dim, m_dim, = mapping_extraction(component, self.mapping_path)

        # From element_mapping (flatten) to index_mapping
        # Note that element with id -666 is turn into index 0 (and later in nan when getting the results)
        elm_mapping_flt = [elm if elm != -666 else np.nan for elm in elm_mapping_flt]

        # Deflattening
        elm_mapping = np.array(elm_mapping_flt).reshape(n_dim, m_dim)

        # Plotting heatmap (fishtail shape)
        plt.figure(figsize=(40,10))
        plot = sns.heatmap(elm_mapping, annot=True, fmt='.0f', annot_kws={"size": 20},
                           cmap=ListedColormap(['whitesmoke']),
                           linewidths=1, linecolor='white',
                           xticklabels=x_labels, yticklabels=y_labels);
        
        plot.set_xticklabels(plot.get_xmajorticklabels(), fontsize = 30);
        plot.set_yticklabels(plot.get_ymajorticklabels(), fontsize = 30);
        ticklbls = plot.get_xticklabels(which='both')
        for x in ticklbls:
            x.set_ha('left')
        
        plt.yticks(rotation=0)
    
        plt.title('Element mapping of component: {}'.format(component), fontsize = 20)
        #plt.xlabel('X-label', fontsize = 15)
        #plt.ylabel('Y-label', fontsize = 15)
        
        plot_img = fig2img(plot.get_figure())
        
        # Saving results in the excel workbook if required
        if excel:
            self.ws_counter += 1
            tag = 'Component: {}. Mapping of elements'.format(component)
            save_in_excel(self.workbook, self.ws_counter, tag, [elm_mapping], img = plot_img)

        return
