import glob
from quickestpects.format import createdocx

def loadxlsx():
    """Load the xlsx file"""
    folder_path = "./xlsx/"
    imgs_path = "./imgs/"
    xlsx_files = glob.glob(folder_path + "*.xlsx")
    
    for xlsx_file in xlsx_files: #loop through all the files
        createdocx(xlsx_file, imgs_path)

def main():
    loadxlsx()

if __name__ == "__main__":
        main()