from app.buildDocx import createdocx
import glob

def main():
    """Load the xlsx file and create DOCX"""
    
    folder_path = "./xlsx/" 
    imgs_path = "./imgs/"
    xlsx_files = glob.glob(folder_path + "*.xlsx")
    
    for xlsx_file in xlsx_files:
        createdocx(xlsx_file, imgs_path)

if __name__ == "__main__":
    main()