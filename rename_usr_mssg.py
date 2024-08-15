import os, re, sys
import win32com.client

"""
    nomaina mapē esošajām vēstulēm nosaukumus BEZ rekursijas, lai vieglāk ir saprast, ko ir Namejā jādara steidzami un ko nav tik steidzami
// https://stackoverflow.com/questions/26322255/parsing-outlook-msg-files-with-python 06.2023
"""

CONST_DIR_SLASH = "\\"
CONST_FILE_TYPE_MSG = ".msg"
CONST_FILE_FIRST_NAME = "Pieejas dati Ozolam "
# šis nav 255, jo izskatās, ka rename funkcija visus 255 char skaita kopā arī ar faila atrašanās vietas direktorijas celiņu
CONST_FILE_NAME_SIZE = 120

# \ norāda, ka "." ir punkts, nevis jebkura simbola attēlojums regexī
regEx_msg = r"\.msg$"
regEx_msg_f_name = r"^Pieejas dati Ozolam.*"
regEx_msg_replace_name = r"Pieejas dati Ozolam.*msg"
regEx_msgUsrName = re.compile(r"Lietotājvārds: \w+")
# regEx_msgExpiration = re.compile(r"Izpildes termiņš: \d{2}.\d{2}.\d{4} \d{2}:\d{2}")
# regEx_msgDoc = re.compile(r"Dokuments: .*\(")

directory = os.getcwd()
# if directory dont have any subdir, then by logic of code we still can work in working main dir
working_dir_list = [directory + CONST_DIR_SLASH]

msg_usr_name = []
msg_usr_name_str = ""

msg_expiration = []
msg_expiration_str = ""

msg_doc = []
msg_doc_str = ""

msg_extension = 0

# will get array with listdir and does it recursively
def read_folder_list(dir, dir_slash):

    dir_list = os.listdir(dir)
    for dir_incr in dir_list:
        next_dir = dir + dir_slash + dir_incr
        
        if os.path.isdir(next_dir):
        
            working_dir_list.append(next_dir)
            read_folder_list(next_dir, dir_slash)
        
    # print("script end")
    return

# testa izdruka ar tabbiem, ja tādi ir
# palīdz pārskatīt dirketoriju saturu, kad ir sagatavots to saraksts apstrādei
def test_print(list, tab):
    for i in list:
        print(tab + str(i))

# funkcijai "msg_file" tiek pārsauksts uz "dir_file", jo "msg_file" ir ar visu path, nevis tikai faila nosaukums
def ren_msgName(dir_file, file_names):
    try:   
        # os.system("echo %time%")
        ## mp4_file = dir_file.replace(".mkv", ".mp4")
        msg = outlook.OpenSharedItem(dir_file)
        
        # ar string replace tiek noņemti visi liekie simboli, kurus neļauj likt jaunajiem failu nosaukumiem
        # print(msg.Body) # print tikai priekš testa
        msg_usr_name = regEx_msgUsrName.findall(msg.Body)
        msg_usr_name_str = "".join(msg_usr_name)
        msg_usr_name_str = msg_usr_name_str.replace("Lietotājvārds: ", "")
        
        # print(f"dir_file: {dir_file}")
        # print(f"msg_usr_name_str: {msg_usr_name_str}")
        
        rename_name = CONST_FILE_FIRST_NAME + msg_usr_name_str + CONST_FILE_TYPE_MSG
        ren_dir_file = ""
        ren_dir_file = re.sub(regEx_msg_replace_name, rename_name, dir_file)
        
        # print(f"\trename_name: {rename_name}")
        # print(f"\tdir_file: {dir_file}")
        # print(f"\tren_dir_file: {ren_dir_file}")
        
        # for TESTING purpouses to see what kind data resides in walues and next code logic in 'if else' staement works as intended
        # print(fr"file new name: {rename_name}")
        # print(fr"file new name size: {len(rename_name)}")
        # if (len(rename_name) > CONST_FILE_NAME_SIZE):
            # print(fr"file new name shorter wersion: {rename_name[0:CONST_FILE_NAME_SIZE]}")
        # print(fr"file new path name: {ren_dir_file}")
        # print(fr"files in current directory: {file_names}")
        
        del msg # atbrīvo vietu atmiņā ?
        
        # there are two cases when should avoid renaming:
        #   one, when in directory already resides file in renamed form
        #   second, when there is file which WILL be renamed in renamed form, which already resides in current directory as separate file as known as file dublicate, but with different name
        if (rename_name not in dir_file and rename_name not in file_names):
            if (len(rename_name) > CONST_FILE_NAME_SIZE): # windows and linux cant have file names which are longer than 255 char, but paths can be longer than 255 chars
                ren_dir_file = ren_dir_file.replace(rename_name, rename_name[0:CONST_FILE_NAME_SIZE] + CONST_FILE_TYPE_MSG)  # rename_name[0:254] can print first n chars as from char array, in this case, prints from 0 to 254 chars
            os.rename(dir_file, ren_dir_file)
            
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
    
# šis skripts apskata visas direktorijas, kas atrodas run faila mapē
# directory ecomes 'item' from workig_dir_list
read_folder_list(directory, CONST_DIR_SLASH)
# test_print(working_dir_list, "")
for item in working_dir_list:
    # print(item)
    # directory mainīsies atkarībā no direktoriju saraksta, ka ir pirms tam iegūts
    file_names = os.listdir(item)
    # test_print(file_names, "\t")
    msg_file = ""
    
    # file_names ir atseviķi katrai direktrijai savs!!!!!!!!!!!!!!

    # print("working dir: \n \t" + item)
    # izšķiro, ka tikai tos msg failus ar nosaukumu, kas satur "Pieejas dati Ozolam", apstrādā un ignorē visu pārējo
    for d in file_names:
        x = re.search(regEx_msg, d)
        y = re.search(regEx_msg_f_name, d)
        
        if x is not None and y is not None:
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