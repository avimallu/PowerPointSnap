
from oletools.olevba import VBA_Parser, TYPE_OLE, TYPE_OpenXML, TYPE_Word2003_XML, TYPE_MHTML
import os.path
import zipfile
import os
import shutil

source_file = 'Snap.pptm'
folder_name = source_file.split('.')[0]
if os.path.exists(os.getcwd() + '\\' + folder_name):
    shutil.rmtree(os.getcwd() + '\\' + folder_name)
os.mkdir(folder_name)

vbaparser = VBA_Parser(source_file)

for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_macros():
    
    name_string = folder_name + '/' + vba_filename
    file_ = open(name_string, 'w')
    file_.write(vba_code)
    file_.close()
    
vbaparser.close()
    
customUI = zipfile.ZipFile(source_file, 'r').extract('customUI/customUI14.xml', folder_name)
#file_ = open(folder_name + '\\CustomUI.xml', 'w')
#file_.write(customUI)
#file_.close()
