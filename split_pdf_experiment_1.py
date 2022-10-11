import os
import PyPDF2
from PyPDF2 import PdfFileWriter, PdfFileReader
import fire
import fitz
import collections
import zipfile
import pandas as pd
import shutil


class PDFSplitter(object):

    def check_duplicate(self,inputallplansdir):

        list_all_pdf = []
        a=0
        for dirpath, dirnames, filenames in os.walk(inputallplansdir):
            for file in filenames:
                if file.endswith(".pdf"):
                    a+=1
                    print(a,file)
                    list_all_pdf.append(file)
        
        check_list_all_pdf = ([item for item, count in collections.Counter(list_all_pdf).items() if count > 1])
        print(check_list_all_pdf)

    def unzip(self,inputsrcfolder,outputrawfolder):

        for dirpath, dirnames, filenames in os.walk(inputsrcfolder):
            for f in filenames:
                if f.endswith(".zip"):
                    with zipfile.ZipFile(os.path.join(dirpath,f),'r') as zip_ref:
                        zip_ref.extractall(outputrawfolder)


    def split_pdf(self,inputdir,outputdir):

        outputdir = outputdir+"/"+"Processed_splitpdf"
        nuos = ["Chorus","City Link","LINZ","Spark","Transpower","Vector","Vocus","Vodafone","Watercare"]

        a=0
        multiple_nuo_list = []
        for dirpath, dirnames, filenames in os.walk(inputdir):
            for file in filenames:
                folder_origin = ((dirpath.split("/"))[-1]).replace(" ","")
                if file.endswith((".pdf",".PDF")):
                    match = [ele for ele in nuos if ((ele.lower()).replace(" ","") in (file.lower()).replace(" ",""))]
                    a+=1
                    if len(match) == 0:
                        filename0 = os.path.join(dirpath,file)
                        inputpdf0 = PdfFileReader(open(filename0, "rb"),strict=False)
                        print(a,file,"Unclassified")

                        for i0 in range(inputpdf0.numPages):
                            output0 = PdfFileWriter()
                            output0.addPage(inputpdf0.getPage(i0))
                            split_filename0 = ((file.split("."))[0]).replace(" ","")
                            concat_folder_filename0 = folder_origin+"_Unclassified_"+split_filename0
                            os.makedirs(outputdir+"/"+folder_origin+"/01_unclassified-pdf",exist_ok=True)
                            with open(outputdir+"/"+folder_origin+"/01_unclassified-pdf"+ "/" +f"{concat_folder_filename0}_%02d.pdf" % (i0+1), "wb") as outputStream0:
                                output.write(outputStream0)
                    
                    elif len(match) == 1:
                        filename = os.path.join(dirpath,file)
                        inputpdf = PdfFileReader(open(filename, "rb"),strict=False)
                        print(a,file,match[0])

                        for i in range(inputpdf.numPages):
                            output = PdfFileWriter()
                            output.addPage(inputpdf.getPage(i))
                            split_filename = ((file.split("."))[0]).replace(" ","")
                            concat_folder_filename = folder_origin+"_"+match[0]+"_"+split_filename
                            os.makedirs(outputdir+"/"+folder_origin+"/"+match[0],exist_ok=True)
                            with open(outputdir+"/"+folder_origin+"/"+match[0]+ "/" +f"{concat_folder_filename}_%02d.pdf" % (i+1), "wb") as outputStream:
                                output.write(outputStream)
                    else:
                        multiple_nuo_list.append(file)
                    
                else:
                    os.makedirs(outputdir+"/"+folder_origin+"/"+"00_non-pdf",exist_ok=True)
                    shutil.copyfile(os.path.join(dirpath,file),os.path.join(outputdir+"/"+folder_origin+"/"+"00_non-pdf",file))
        print(multiple_nuo_list)
    
    
    def split_pdf_to_image(self,inputprocessedsplitpdfdir,outputxlsx,mode="default",dpi_val="Null",format="tiff"):
        
        byd_request_number_list = []
        nuos_list = []
        plan_name_list = []
        
        a=0
        for dirpath, dirnames, filenames in os.walk(inputprocessedsplitpdfdir):
            for file in filenames:
                if file.endswith((".pdf",".PDF")):
                    if mode == "default" and dpi_val == "Null":
                        file_path = os.path.join(dirpath,file)
                        doc = fitz.open(file_path)  # open document
                        for page in doc:
                            a+=1
                            pix = page.get_pixmap()  # render page to an image
                            name, extension = os.path.splitext(file)
                            normpath = os.path.normpath(dirpath)
                            split_path = normpath.split(os.sep)
                            byd_request_number_list.append(split_path[-2])
                            if split_path[-1] == "01_unclassified-pdf":
                                nuos_list.append("Unclassified")
                            else:
                                nuos_list.append(split_path[-1])
                            plan_name_list.append(name+"."+format)
                            outputconvimagedir = dirpath.replace("Processed_splitpdf","Processed_image")
                            os.makedirs(outputconvimagedir,exist_ok=True)
                            pix.save(os.path.join(outputconvimagedir,f"{name}.{format}"))
                            print(a,f"Successfully generated {format} of {file}")
                    elif mode == "high-res":
                        file_path = os.path.join(dirpath,file)
                        dpi = dpi_val
                        zoom = dpi/72
                        magnify = fitz.Matrix(zoom,zoom)
                        doc = fitz.open(file_path)  # open document
                        for page in doc:
                            a+=1
                            pix = page.get_pixmap(matrix=magnify)  # render page to an image
                            name, extension = os.path.splitext(file)
                            normpath = os.path.normpath(dirpath)
                            split_path = normpath.split(os.sep)
                            byd_request_number_list.append(split_path[-2])
                            if split_path[-1] == "01_unclassified-pdf":
                                nuos_list.append("Unclassified")
                            else:
                                nuos_list.append(split_path[-1])
                            plan_name_list.append(name+"."+format)
                            outputconvimagedir = dirpath.replace("Processed_splitpdf","Processed_image")
                            os.makedirs(outputconvimagedir,exist_ok=True)
                            pix.save(os.path.join(outputconvimagedir,f"{name}.{format}"))
                            print(a,f"Successfully generated {format} of {file}")
                else:
                    outputconvimagedir = dirpath.replace("Processed_splitpdf","Processed_image")
                    os.makedirs(outputconvimagedir,exist_ok=True)
                    shutil.copyfile(os.path.join(dirpath,file),os.path.join(outputconvimagedir,file))
        
        df = pd.DataFrame()
        df['BYD Request Number'] = byd_request_number_list
        df['Plan Name'] = plan_name_list
        df['NUO'] = nuos_list
        df['Assigned to'] = ""
        df['Georeference'] = ""
        df['Upload'] = ""
        df['QA by'] = ""
        df['QA'] = ""
        df['Phase 2 Time Started'] = ""
        df['Phase 3 Time Ended'] = ""
        df.to_excel(outputxlsx)
    
    # def split_pdf(self,inputdir,outputdir):

    #     nuos = ["Chorus","City Link","LINZ","Spark","Transpower","Vector","Vocus","Vodafone","Watercare"]

    #     folder_origin_list = []
    #     a=0
    #     success_file = []
    #     for dirpath, dirnames, filenames in os.walk(inputdir):
    #         for file in filenames:
    #             folder_origin = ((dirpath.split("/"))[-1]).replace(" ","")
    #             if file.endswith((".pdf",".PDF")):
    #                 for h in nuos:
    #                     if (h.lower()).replace(" ","") in (file.lower()).replace(" ",""):
    #                         a+=1
    #                         print(a,file,h)
    #                         filename = os.path.join(dirpath,file)
    #                         inputpdf = PdfFileReader(open(filename, "rb"),strict=False)
    #                         success_file.append(file)

    #                         for i in range(inputpdf.numPages):
    #                             output = PdfFileWriter()
    #                             output.addPage(inputpdf.getPage(i))
    #                             split_filename = ((file.split("."))[0]).replace(" ","")
    #                             concat_folder_filename = folder_origin+"_"+h+"_"+split_filename
    #                             folder_origin_list.append(folder_origin)
    #                             os.makedirs(outputdir+"/"+folder_origin+"/"+h,exist_ok=True)
    #                             with open(outputdir+"/"+folder_origin+"/"+h+ "/" +f"{concat_folder_filename}_%02d.pdf" % (i+1), "wb") as outputStream:
    #                                 output.write(outputStream)
    #                     else:
    #                         os.makedirs(outputdir+"/"+folder_origin+"/"+"01_unclassified-pdf",exist_ok=True)
    #                         shutil.copyfile(os.path.join(dirpath,file),os.path.join(outputdir+"/"+folder_origin+"/"+"01_unclassified-pdf",file))
    #             else:
    #                 os.makedirs(outputdir+"/"+folder_origin+"/"+"00_non-pdf",exist_ok=True)
    #                 shutil.copyfile(os.path.join(dirpath,file),os.path.join(outputdir+"/"+folder_origin+"/"+"00_non-pdf",file))
        
    #     unique_folder_origin_list = list(set(folder_origin_list))
    #     for g in unique_folder_origin_list:
    #         for dirpath_1, dirname_1, filenames_1 in os.walk(outputdir+"/"+g+"/"+"01_unclassified-pdf"):
    #             for f_unc in filenames_1:
    #                 if f_unc in success_file:
    #                     os.remove(os.path.join(outputdir+"/"+g+"/"+"01_unclassified-pdf",f_unc))
    #                 else:
    #                     filename_unc = os.path.join(dirpath_1,f_unc)
    #                     inputpdf_unc = PdfFileReader(open(filename_unc, "rb"),strict=False)

    #                     for i_unc in range(inputpdf_unc.numPages):
    #                         output_unc = PdfFileWriter()
    #                         output_unc.addPage(inputpdf_unc.getPage(i_unc))
    #                         split_filename_unc = ((f_unc.split("."))[0]).replace(" ","")
    #                         concat_folder_filename_unc = g+"_Unclassified_"+split_filename_unc
    #                         with open(outputdir+"/"+g+"/"+"01_unclassified-pdf"+ "/" +f"{concat_folder_filename_unc}_%02d.pdf" % (i_unc+1), "wb") as outputStream_unc:
    #                             output_unc.write(outputStream_unc)
    #                     os.remove(os.path.join(dirpath_1,f_unc))

if __name__ == "__main__":
    fire.Fire(PDFSplitter)