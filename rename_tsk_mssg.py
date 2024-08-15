import os, re, sys
import win32com.client

"""
    nomaina mapē esošajām vēstulēm nosaukumus BEZ rekursijas, lai vieglāk ir saprast, ko ir Namejā jādara steidzami un ko nav tik steidzami
// https://stackoverflow.com/questions/26322255/parsing-outlook-msg-files-with-python 06.2023
"""

CONST_DIR_SLASH = "\\"
CONST_FILE_TYPE_MSG = ".msg"
# šis nav 255, jo izskatās, ka rename funkcija visus 255 char skaita kopā arī ar faila atrašanās vietas direktorijas celiņu
CONST_FILE_NAME_SIZE = 120

regEx_msg = r"\.msg$"
regEx_msgTask = re.compile(r"Uzdevuma teksts: \w+")
regEx_msgExpiration = re.compile(r"Izpildes termiņš: \d{2}.\d{2}.\d{4} \d{2}:\d{2}")
regEx_msgDoc = re.compile(r"Dokuments: .*\(")
directory = os.getcwd()
working_dir_list = [directory + CONST_DIR_SLASH]

msg_task = []
msg_task_str = ""

msg_expiration = []
msg_expiration_str = ""

msg_doc = []
msg_doc_str = ""

msg_extension = 0

def ren_msgName(msg_file, file_names):
    try:   
        # os.system("echo %time%")
        ## mp4_file = msg_file.replace(".mkv", ".mp4")
        msg = outlook.OpenSharedItem(msg_file)
        
        # ar string replace tiek noņemti visi liekie simboli, kurus neļauj likt jaunajiem failu nosaukumiem
        # print(msg.Body) # print tikai priekš testa
        msg_task = regEx_msgTask.findall(msg.Body)
        msg_task_str = "".join(msg_task)
        msg_task_str = msg_task_str.replace("Uzdevuma teksts: ", "")
        
        msg_expiration = regEx_msgExpiration.findall(msg.Body)
        msg_expiration_str = "".join(msg_expiration)
        msg_expiration_str = msg_expiration_str.replace("Izpildes termiņš: ", "")
        
        msg_doc = regEx_msgDoc.findall(msg.Body)
        msg_doc_str = "".join(msg_doc)
        msg_doc_str = msg_doc_str.replace("Dokuments: ", "")
        
        # print(msg_task_str)
        # print(msg_expiration_str)
        # print(msg_doc_str)
        
        rename_name = msg_task_str + "-" + msg_expiration_str + "-" + msg_doc_str
        rename_name = rename_name.replace(" ", "_")
        rename_name = rename_name.replace(":", "")
        rename_name = rename_name.replace(")", "")
        rename_name = rename_name.replace("(", "")
        rename_name = rename_name.replace("/", "")
        rename_name = rename_name.replace("“", "") # quotes in latvian language
        rename_name = rename_name.replace("”", "") # quotes in latvian language
        
        ren_msg_file = directory + CONST_DIR_SLASH + rename_name + CONST_FILE_TYPE_MSG
        
        # for TESTING purpouses to see what kind data resides in walues and next code logic in 'if else' staement works as intended
        # print(fr"file new name: {rename_name}")
        # print(fr"file new name size: {len(rename_name)}")
        # if (len(rename_name) > CONST_FILE_NAME_SIZE):
            # print(fr"file new name shorter wersion: {rename_name[0:CONST_FILE_NAME_SIZE]}")
        # print(fr"file new path name: {ren_msg_file}")
        # print(fr"files in current directory: {file_names}")
        
        del msg # atbrīvo vietu atmiņā ?
        
        # there are two cases when should avoid renaming:
        #   one, when in directory already resides file in renamed form
        #   second, when there is file which WILL be renamed in renamed form, which already resides in current directory as separate file as known as file dublicate, but with different name
        if (rename_name not in msg_file and rename_name + CONST_FILE_TYPE_MSG not in file_names):
            if (len(rename_name) > CONST_FILE_NAME_SIZE): # windows and linux cant have file names which are longer than 255 char, but paths can be longer than 255 chars
                ren_msg_file = directory + CONST_DIR_SLASH + rename_name[0:CONST_FILE_NAME_SIZE] + CONST_FILE_TYPE_MSG # rename_name[0:254] can print first n chars as from char array, in this case, prints from 0 to 254 chars
            os.rename(msg_file, ren_msg_file)
            
            print('File was renamed to: ' + rename_name + '\n')
        else:
            print('File was not renamed, because it already exists and was already been parsed before by this script\n')

        # r == raw string works well with whitespaces in folder paths
        # f == function string for easier variable and function assignment to string
        # f and r can be combined

    # os.system("echo %time%")
    except Exception as e:
        pass
        # print("Unknown error")
        print(e)
    
    # šis skripts uz doto brīdi darbojas tikai esošā direktorijā
    # ja atrod mkv failus, atrod, ja ne, tad nē, neko vairāk nedara
for item in working_dir_list:
    print(item)
    # directory mainīsies atkarībā no direktoriju saraksta, ka ir pirms tam iegūts
    file_names = os.listdir(item)
    # print(file_names)
    msg_file = ""

    # print("working dir: \n \t" + item)
    # izšķiro, ka tikai msg failus apstrādā un ignorē visu pārējo
    for d in file_names:
        x = re.search(regEx_msg, d)
        
        if x is not None:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            # print("file name: \n \t" + d)
            msg_file = d
            
            print(item + CONST_DIR_SLASH + msg_file)
            ren_msgName(item + CONST_DIR_SLASH + msg_file, file_names)
            
            # sys.exit('Testing mssg file content output')
           
            del outlook
            msg_extension = msg_extension + 1
            # break
        else:
            pass
            
print('\nProcessed doc file count: ' + str(msg_extension))