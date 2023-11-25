from quickestspects.buildDocx import createdocx

import glob

def loadxlsx():
    """Load the xlsx file"""
    folder_path = "./xlsx/" 
    imgs_path = "./imgs/"
    xlsx_files = glob.glob(folder_path + "*.xlsx")
    txt_file = 'quickestspecs.txt'
    
    for xlsx_file in xlsx_files:
        createdocx(xlsx_file, imgs_path, txt_file)
        
def main():
    loadxlsx()

if __name__ == "__main__":
        main()  