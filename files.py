class file:
    def __init__(self):
        pass


import os

CSV_ROOT="C:\\temp\\"
STEP_ROOT="C:\\temp\\"


def now_string():
    x=datetime.datetime.now()
    y=str(x.year)+str(x.month)+str(x.day)+str(x.minute)+str(x.second)
    return y

now=now_string

def get_files_at_folder_by_suffix(path,suffix,exclude,flag=0):
    l = []
    dir_list = os.listdir(path)
    for x in dir_list:
        if x.endswith(suffix) and not(x.startswith(exclude)):
            if flag == 0:
                l.append(path+x)
            else:
                l.append(x)
                
    return l

def get_python_files_at_folder(path,flag=0):
    return get_files_at_folder_by_suffix(path,".py",".",flag)

def get_org_files_at_folder(path,flag=0):
    return get_files_at_folder_by_suffix(path,".org",".",flag)

def get_assembly_files_at_folder(path,flag=0):
    return get_files_at_folder_by_suffix(path,".SLDASM","~$",flag)


def change_suffix(flist,pre,suf):
    s = []
    for i in flist:
      j = os.path.splitext(i)
      s.append(pre+j[0]+suf)
    return s
    
    def get_files_at_folder_by_suffix(self,path,suffix,exclude,flag=0):
        l = []
        dir_list = os.listdir(path)
        for x in dir_list:
            if x.endswith(suffix) and not(x.startswith(exclude)):
                if flag == 0:
                    l.append(path+x)
                else:
                    l.append(x)
        return l

    def get_python_files_at_folder(self,path,flag=0):
        return self.get_files_at_folder_by_suffix(path,".py",".",flag)

    def get_org_files_at_folder(self,path,flag=0):
        return self.get_files_at_folder_by_suffix(path,".org",".",flag)

    def get_assembly_files_at_folder(self,path,flag=0):
        return self,get_files_at_folder_by_suffix(path,".SLDASM","~$",flag)

    def change_suffix(self,flist,pre,suf):
        s = []
        for i in flist:
            j = os.path.splitext(i)
            s.append(pre+j[0]+suf)
        return s
   
    def concatenate_file_list(self,path,pre="",suf=".STEP"):
        i0 = self.get_org_files_at_folder(path)
        i1 = self.get_org_files_at_folder(path,1)
        i2 = change_suffix(i1,pre,".STEP")
        l=[]
        for j in range(len(i0)):
            l.append([i0[j],i1[j],i2[j]])
        return l

def get_basename_from_path (pn):
    pp=pn.split('\\')[-1]
    return pp.split('.')[-2]

def get_suffix_from_path (pn):
    return pn.split('.')[-1]

def get_directory_from_path (pn):
    return os.path.dirname(pn)

def create_step(pn,i):
    a=get_directory_from_path(pn)+"\\"+get_basename_from_path(pn)+str("_")+str(i)+".STEP"
    return a

def get_next_available_step_from_path(pn):
    a=get_directory_from_path(pn)+"\\"+get_basename_from_path(pn)+".STEP"
    i = 0
    while os.path.isfile(Path(a)):
        a=create_step(pn,i)
        i=i+1
    return a

def make_csv_path(model,root=CSV_ROOT):
    s0=model.GetPathname
    s1=get_basename_from_path(s0)
    csv_name=root+s1+"_"+now_string()+".csv"
    return csv_name


def make_step_path(model,root=STEP_ROOT):
    s0=model.GetPathname
    s1=get_basename_from_path(s0)
    csv_name=root+s1+"_"+now_string()+".step"
    return csv_name

