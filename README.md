# OP-Map
A python library with functions and methods for reading .op2 files and element-mapping plotting. It is extensively based on the capabilities of the library **[pyNastran](https://pynastran-git.readthedocs.io/en/latest/#)**.

**Input files:**
+ .op2 file
+ mapping file (in .json format)
+ .bdf file associated to the .op2 file; only if 2D element results need to be transformed to material coordinates

**Capabilities of OP-Map:**
+ read .op2 files created with solver NASTRAN, Altair OPTISTRUCT
+ transform .op2 2D-element results from element coordiante (native format in .op2) to material coordinates for 2D elements
+ plot element mappings of specific structural components as defined in the mapping file.
+ plot mappings of element forces for a specific components and a single load case
+ plot mappings of maximum, minimum or maximum absolute elemment forces for a specific component for all load cases included in the .op2 file
+ change the mapping file initially loaded for a new one
+ list all load cases in the loaded .op2 file
+ save all results in a dedicated Excel workbook


FRLM, v0.6
