import sys
import time
import pythoncom
import win32com.client as win32
import os
import unicodedata
import re
import sys
import csv
print(sys.executable)
import datetime
import logging
from pathlib import Path
from files import *
from sheet import *
from sqlite import *
from mass import *

for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)
logging.basicConfig(
    filename="log.log",
    filemode="a",
    format='%(asctime)s,%(msecs)d %(levelname)-8s [%(pathname)s:%(lineno)d in ' \
           'function %(funcName)s] %(message)s',
    datefmt='%Y-%m-%d:%H:%M:%S',
    level=logging.DEBUG
)
a=1
logging.info('Starting up:'+str(a))

CSV_ROOT="C:\\Users\\30698\\Desktop\\ws-500\\csv\\"
STEP_ROOT="C:\\Users\\30698\\Desktop\\ws-500\\step\\"

SWV = 2022
# API Version
SWAV = SWV-1992

CSV_FOLDER="C://Users//30698//AppData//Roaming//CAD//csv//"
ARG_NULL = win32.VARIANT(pythoncom.VT_DISPATCH, None)

# Consts
sw = win32.Dispatch("SldWorks.Application.{}".format(SWAV))


class part_config:
    def __init__(self, c, f):
        if(c==None):
            self.config=[]
            self.config_names=[]
            self.dictionary=[]
        else:
            self.config=c
            self.config_names=c[-1][4:]
            self.dictionary=[e[0:4] for e in c]
        if(f==None):
            self.fname=""
        else:
            self.fname= f

def get_active_doc():
    return sw.ActiveDoc


def get_open_windows():
    swframe = sw.Frame
    return swframe.modelWindows

def get_open_windows():
    swframe = sw.Frame
    return swframe.modelWindows

def get_open_windows_fnames():
    return [i.ModelDoc.GetPathName for i in get_open_windows()]

def get_open_windows_pairs():
    return [[i.ModelDoc,i.Title,i.ModelDoc.GetPathName] for i in get_open_windows()]



def active_path():
    Model = sw.ActiveDoc
    a=Model.GetPathName
    a=a.replace("\\","\\\\")
    return a

def return_active_doc_path():
    """
    returns active doument's path
    """
    Model = sw.ActiveDoc
    a=Model.GetPathName
    a=a.replace("\\","\\\\")
    return a

def return_active_doc_path1():
    """
    returns active doument's path
    """
    Model = sw.ActiveDoc
    a=Model.GetPathName
    #a=a.replace("\\","\\\\")
    return a

def export_configuration_to_step_file(model,conf_name,filepath):
    model2.ShowConfiguration2(conf_name)
    model2.EditRebuild3
    save_model_as(model2,filepath)
         
def get_active_configuration_name(model):
    swconfig=model.GetActiveConfiguration
    return swconfig.Name

def get_model_configuration_names(model):
    config_names = model.GetConfigurationNames
    return [i for i in config_names]

            
def change_dimensions_at_part(model,part):
    return [change_dimensions_from_configuration_name(model,conf,part.dictionary) for conf in part.config_names]

def get_assemblies_at_folder(path):
    l = []
    dir_list = os.listdir(path)
    for x in dir_list:
        if x.endswith(".SLDASM"):
            if not(x.startswith("~$")):
                l.append(x)
    return l

def save_model_as(model,step_file):
    arg1 = win32.VARIANT(pythoncom.VT_DISPATCH, None)
    arg2 = win32.VARIANT(pythoncom.VT_BOOL, 0)
    arg3 = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    arg4 = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    ret = model.Extension.SaveAs2(step_file, 0, 1, arg1, "", arg2, arg3, arg4)
    return ret

def export_dwg_pdf(model):
    a= model.GetPathName
    #directory=get_directory_from_path(a)
    name=get_basename_from_path(a)+"-"+create_sheet_filename(model)
    directory="C:\\temp\\"
    pdf= directory + "\\" + name + "_" +now() + ".pdf"
    pdf.replace("\\","\\\\")
    save_model_as(model,pdf)
    
    dwg= directory + "\\" + name + "_" +now() + ".dwg"
    dwg.replace("\\","\\\\") 
    save_model_as(model,dwg)

    dxf= directory + "\\" + name + "_" +now() + ".dxf"
    dwg.replace("\\","\\\\") 
    save_model_as(model,dxf)

def export_file_to_step(file):
    model=open_activate_doc(file)
    return export_model_to_step(model)


def export_model_to_step(model):
    arg1 = win32.VARIANT(pythoncom.VT_DISPATCH, None)
    arg2 = win32.VARIANT(pythoncom.VT_BOOL, 0)
    arg3 = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    arg4 = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    name=model.GetPathName
    step_file=get_next_available_step_from_path(name)
    print(step_file)
    ret = model.Extension.SaveAs2(step_file, 0, 1, arg1, "", arg2, arg3, arg4)
    return ret


def export_active_model_to_csv(csv_name):
    model=get_active_doc()
    value = model.GetComponents(False)
    C1 = get_active_components(model,value)
    write_table_to_csv(C1,csv_name)


def export_model_to_csv(model1,csv_name):
    value = model1.GetComponents(False)
    C1 = get_active_components(model1,value)
    write_table_to_csv(C1,csv_name)


def get_components_from_doc_and_write_them_to_csv(doc,name):
    #doc = sw.ActiveDoc
    components_list = doc.GetComponents(False)
    C = get_active_components(doc,components_list)
    write_table_to_csv(C,name)
    return C

def get_active_components(Model,value):
    C = [[]]
    for f in value:
        try:
            fmodel = f.GetModelDoc2 # returns an object
        except:
            fmodel = "NaN"
            continue
        try:
            fpath = fmodel.GetPathName
        except:
            fpath = "NaN"
            continue
        try:
            fname = f.Name2
        except:
            fname = "NaN"
        try:
            fconfig = f.ReferencedConfiguration
        except:
            fconfig = "NaN"
        try:
            fmat = fmodel.MaterialIdName
        except:
            fmat = "NaN"
        try:
            ftitle = fmodel.GetTitle
        except:
            ftitle = "NaN"
        mp=get_mass_properties(fmodel)    
        C.append([fname,
                  fpath,
                  fconfig,
                  fmat,
                  ftitle,
                  mp["volume"],
                  mp["surface"],
                  mp["mass"]
                  ])
    return C
    #  C.append([f.Name2, ff.GetPathname, ff.GetTitle, ActiveConfig])

def SetPartCustomMaterial(Model,configuration,material):
    ARG_NULL = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    Model.SetMaterialPropertyName2(configuration, "C://ProgramData//SolidWorks//SOLIDWORKS 2022//Custom Materials//Custom Materials.sldmat", material)

    
    
def select_component(Model,component):
    Model.Extension.SelectByID2(component, "COMPONENT", 0, 0, 0, False, 0, ARG_NULL, 0)
    swmgr = Model.SelectionManager
    swcomp = swmgr.GetSelectedObject6(1,-1)
    return swcomp


    
def change_configuration_of_component(Model,component,configuration):
    Model.Extension.SelectByID2(component, "COMPONENT", 0, 0, 0, False, 0, ARG_NULL, 0)
    swmgr = Model.SelectionManager
    swcomp = swmgr.GetSelectedObject6(1,-1)
    swcomp.ReferencedConfiguration = configuration
    Model.EditRebuild3

def change_dimensions_from_configuration_name(Model,conf,base_descr):
    code = conf.split("-")
    add_configuration(Model,conf,conf)
    Model.ShowConfiguration2(conf)
    for i in base_descr:
        if i[2] != '':
            j = int(i[2])
            print(i)
            a1 = Model.Parameter(i[1])
            try:
                a1.SetSystemValue3(float(code[j])*i[3],1,ARG_NULL)
            except:
                print("No such dimension:"+i[1])
    Model.EditRebuild3


"""
| Configuration name: | var-1 |
| Description:        | var-5 |
| Comment:            | var-2 |
| Part number on BOM: | var-3 |
"""
def add_configuration2(Model,conf,partno,description):
    """ add configuration to Model """
    ConfigMgr = Model.ConfigurationManager
    ConfigMgr.AddConfiguration(conf , conf, partno, 1, "", description)
    Model.ShowConfiguration2(conf)



def add_configuration(Model,conf,description):
    """ add configuration to Model """
    ConfigMgr = Model.ConfigurationManager
    ConfigMgr.AddConfiguration(conf , conf, conf, 1, "", description)
    Model.ShowConfiguration2(conf)


def wall(conf):
    """generates base scrubber configuration from assembly configuration description"""
    internal_slot_diameter = int(conf[1]) #internal Diameter
    length = int(conf[2])
    thickness = int(conf[3])
    mylist = [internal_slot_diameter, length, thickness]
    strlist = [str(i) for i in mylist]
    return '-'.join(strlist)

def down_base(conf):
    """generates base scrubber configuration from assembly configuration code description"""
    internal_slot_diameter = int(conf[0]) #internal Diameter
    external_diameter = internal_slot_diameter + 100
    slot_width = int(conf[3])+3
    slot_cut_depth = 5
    thickness = 10
    mylist = [internal_slot_diameter, external_diameter, slot_width,slot_cut_depth,thickness]
    strlist = [str(i) for i in mylist]
    return '-'.join(strlist)




def suppress_feature(Model,featname):
    ModelExt = Model.Extension
    ModelExt.SelectByID2(featname,"BODYFEATURE", 0, 0, 0, False, 0, ARG_NULL, 0)
    Model.EditSuppress2
    Model.EditRebuild3

def supress_feature_on_all_configurations(model,feature):
    configurations=get_model_configuration_names(model)
    for conf in configurations:
        model.ShowConfiguration2(conf)
        suppress_feature(model,feature)


        
def unsuppress_feature(Model,featname):
    ModelExt = Model.Extension
    ModelExt.SelectByID2(featname,"BODYFEATURE", 0, 0, 0, False, 0, ARG_NULL, 0)
    Model.EditUnSuppress2
    Model.EditRebuild3

def write_table_to_csv(table,pathname):
    with open(pathname, 'w', newline='') as csvfile:
        for ci in table:
            spamwriter = csv.writer(csvfile) #, dialect='excel')
            spamwriter.writerow(ci)

def open_file1(pathname):
    path_model = Path(pathname)
    ext = path_model.suffix
    swOpenDocOptions_Silent = 0
    arg5 = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 128)
    arg6 = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 128)
    if ext == ".SLDPRT" or ext == ".sldprt":
        enum = 1
    if ext == ".SLDASM" or ext == ".sldasm":
        enum = 2
    if ext == ".SLDASM" or ext == ".sldasm":
        enum = 2
    if ext == ".SLDDRW" or ext == ".slddrw":
        enum = 3
    Model = sw.OpenDoc6(pathname, enum, swOpenDocOptions_Silent, "", arg5,arg6)
    return Model

def open_assembly(pathname):
    enum = 2 #swDocAssembly
    return open_file(pathname,enum)

def open_part(pathname):
    enum = 1  #swDocPart
    return open_file(pathname,enum)

def open_doc(pathname):
    enum = 3  #swDocDRAWING
    return open_file(pathname,enum)

def activate_doc(pathname):
    arg6 = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 128)
    model=sw.ActivateDoc3(pathname,False,2,arg6)
    return model

def open_activate_doc(pathname):
    model1=open_file1(pathname)
    model1=activate_doc(pathname)
    return model1


def GetPathname(doc):
    Model = doc.GetPathName
    return Model

def active_document_components():
    list = []
    doc = sw.ActiveDoc
    value = doc.GetComponents(False)
    C = []
    for f in value:
      # print(f.FirstFeature)
      #print(f.Name2)
      try:
        ff = f.GetModelDoc2
        sMatDB = ["h"]
        AC = ff.GetActiveConfiguration.Name
        mat =""
        #mat = ff.GetMaterialPropertyName2(AC, sMatDB)
        mat = ff.MaterialIdName
        C.append([f.Name2, ff.GetPathname, ff.GetTitle, AC,mat])
      except:
        1
    return C

def change_component_configuration2(model,component,configuration):
    """"
    changes components configuration at active doc

    Parameters
    ----------
    model: model to change configuration
    component: string name
    configuration: string name
    
    Returns
    -------
    Nothing

    """
    model.Extension.SelectByID2(component, "COMPONENT", 0, 0, 0, False, 0, ARG_NULL, 0)
    swmgr = model.SelectionManager
    swcomp = swmgr.GetSelectedObject6(1,-1)
    swcomp.ReferencedConfiguration = configuration
    model.EditRebuild3


    
def change_component_configuration(component,configuration):
    """"
    changes components configuration at active doc

    Parameters
    ----------
    component: string name
    configuration: string name
    
    Returns
    -------
    Nothing

    """
    Model = sw.ActiveDoc
    Model.Extension.SelectByID2(component, "COMPONENT", 0, 0, 0, False, 0, ARG_NULL, 0)
    swmgr = Model.SelectionManager
    swcomp = swmgr.GetSelectedObject6(1,-1)
    swcomp.ReferencedConfiguration = configuration
    Model.EditRebuild3

def change_value_of_dimension_at_configuration (value,parameter,configuration):
    Model = sw.ActiveDoc
    Model.ShowConfiguration2(configuration)
    D1 = Model.Parameter(parameter)
    D1.SystemValue = value * 0.001

def change_value_of_dimension_at_configuration_no_scale (value,parameter,configuration):
    Model = sw.ActiveDoc
    Model.ShowConfiguration2(configuration)
    D1 = Model.Parameter(parameter)
    D1.SystemValue = value
    
def add_configuration_to_active_doc(s):
    Model = sw.ActiveDoc
    ConfigMgr = Model.ConfigurationManager    
    ConfigMgr.AddConfiguration(s, s, s, 1, "", s)

def get_all_open_documents ():
    list = []
    doc = sw.GetFirstDocument
    
    while doc != None:
        #      list.append(doc.GetTitle)
        #      list.append(doc.GetType)
        #      list.append(doc.GetPathName)
        list_element = [doc.GetTitle, doc.GetType, doc.GetPathname]
        list.append(list_element)
        doc = doc.GetNext
        return list

def change_model_dimensions(model,array):
    for i in array:
        D1 = swModel.Parameter(i[0])
        D1.SetSystemValue3(i[1] * 0.001 , 1)
    

def list_of_model_configurations(Model):
    """
    Returns list of model configuaration names
    """
    swConfMgr = Model.ConfigurationManager
    swConf = swConfMgr.ActiveConfiguration
    vConfigNameArr = Model.GetConfigurationNames
    return vConfigNameArr


def name_of_active_configuration(Model):
    """
    Returns name of active configuration
    """
    swConfMgr = Model.ConfigurationManager
    swConf = swConfMgr.ActiveConfiguration
    vConfigNameArr = Model.GetConfigurationNames
    return swConf.Name

#doc = sw.GetFirstDocument
#print(name_of_active_configuration(doc))

#value = doc.GetComponents(False)

def is_property (val):
   regexp1 = ".@@+"
   if re.search(regexp1,val):
       ss = re.split("@@",val)
       return ss[0]
   else:
       return False

#assert(is_property("ass@@1212") == "ass")   
   
def create_active_property_name(Model,property_name):
   fname = r'{}'.format(Model.GetPathName)
   basename = os.path.basename(fname)
   ConfigurationName = Model.GetActiveConfiguration.Name
   mystring = property_name + "@@" + ConfigurationName + "@" + basename
   return mystring

#print(create_active_property_name(sw.ActiveDoc,"hello"))

def add_properties_from_org_table(Model,inp):
    """
    add properties from org table(list) to a model
    Model properties to add
    inp table list
    """
    swConfMgr = Model.ConfigurationManager
    swConf = swConfMgr.ActiveConfiguration
    ActiveConf = Model.GetActiveConfiguration
    CPM = ActiveConf.CustomPropertyManager
    vNameArr = CPM.GetNames

    print("Number of custom properties for configuration {" + ActiveConf.Name + "} are " + str(CPM.Count))

    for f in inp:
#     print(f[0],":",'\"'+str(f[1])+'\"')
      PropertyName = f[0]

      ss = is_property(str(f[1]))
      if ss:
        PropertyValue = "\"" + create_active_property_name(Model,f[1]) + "\""
      else:
        PropertyValue = str(f[1])
      PropertyType = 30 #30 for text
      OverWrite = 1
      CPM.Add3(PropertyName, PropertyType, PropertyValue, OverWrite)
    #if arg3=0 it does not overwrite, returns 2 if exists
    #if arg3=1 it overwrites value
    #30 type is string
    #CPM.Add3("CustomProperty",30,"MyCustomProperty2",1)

    
def Model_table_view (doc):
    """

    """
    value = doc.GetComponents(False)
    C = []
    for f in value:
        try:
            ff = f.GetModelDoc2
            sMatDB = ["h"]
            AC = ff.GetActiveConfiguration.Name
            mat =""
            mat = ff.MaterialIdName
            C.append([f.Name2, ff.GetPathname, ff.GetTitle, AC,mat])
        except:
            print("Error")
    return C
            
def print_property_list(Model):
    """
    
    """

    ActiveConf = Model.GetActiveConfiguration
    CPM = ActiveConf.CustomPropertyManager
    vNameArr = CPM.GetNames
    print("File = " + Model.GetPathName)
    print("Number of custom properties for configuration {" + ActiveConf.Name + "} are " + str(CPM.Count))
    for f in CPM.GetNames:
        print( str(f) + " : " + str(CPM.Get(f)) + " : " + str(CPM.GetType2(f)))

    
def get_property_list(Model):
    """
    
    """
    ActiveConf = Model.GetActiveConfiguration
    CPM = ActiveConf.CustomPropertyManager
    vNameArr = CPM.GetNames
    
    C = []
    for f in CPM.GetNames:
        try:
            C.append( [str(f) , str(CPM.Get(f)) , str(CPM.GetType2(f)) ] )
        except:
            print("Error")
    return C
