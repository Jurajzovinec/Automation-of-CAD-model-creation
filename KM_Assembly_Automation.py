import csv
import os
import tkinter as tk
from difflib import SequenceMatcher
import creopyson
from PIL import ImageTk, Image
import sys
from tkinter import filedialog
from tkinter import messagebox
import shutil
import time
import xlrd
import logging


#--L------OO----GG----
#--L-----O--O--G------
#--L-----O--O--G--GG--
#--LLLL---OO----GGG---

#---Logging---
try:
    os.remove("assembly_automation.log")
except:
    pass
logging.basicConfig(filename= "assembly_automation.log", level=logging.INFO)
logger = logging.getLogger()


# ------GG-----U----U----I--------
# -----G--G----U----U----I--------
# ----G--------U----U----I--------
# ----G---GGG--U----U----I--------
# -----G---G----U--U-----I--------
# ------GGG------UU------I--------


# -------MAIN_FUNCTION-------
def build_graphical_user_interface():
    creoson_setup()

    global master
    master = tk.Tk()
    master.geometry("500x600")
    master.resizable(0, 0)
    master.title('Kraussmaffei Assembly Automation')

    rows = 0
    while rows < 13:
        master.rowconfigure(rows, weight=1)
        master.columnconfigure(rows, weight=1)
        rows += 1

    background = ImageTk.PhotoImage(Image.open("IconPictures/Graphical_User_Int_Theme.png"))
    tk.Label(master, image=background).grid(row=0, column=0, rowspan=100, columnspan=100)

    global order_number_entry
    order_number_entry = tk.Entry(master)
    order_number_entry.config(width=20, font=('Helvetica', 14), borderwidth=4)
    order_number_entry.insert(0, '--New_Order_Number--')
    order_number_entry.grid(row=11, column=1)

    global check_mark_ico
    check_mark_ico = ImageTk.PhotoImage(Image.open("IconPictures/check_mark.png"))

    # Confirmation button for picking gm
    global picked_type_confirmation_button
    picked_type_confirmation_button = tk.Button(image=check_mark_ico, command=list_CAD_mastermodels)
    picked_type_confirmation_button.grid(row=2, column=2, rowspan=2)

    #10-07-2020 - Rename entries have been disabled. To enable this function uncomment following blocks. (km-macro with condition)

    #global rename_from_entry
    #rename_from_entry = tk.Entry(master)
    #rename_from_entry.config(width=20, font=('Helvetica', 13), borderwidth=2)
    #rename_from_entry.insert(0, '--Rename-From--')
    #rename_from_entry.grid(row=12, column=0)
    #
    #global rename_to_entry
    #rename_to_entry = tk.Entry(master)
    #rename_to_entry.config(width=20, font=('Helvetica', 13), borderwidth=2)
    #rename_to_entry.insert(0, '--Rename-To--')
    #rename_to_entry.grid(row=12, column=1)
    #
    #global rename_must_contain_entry
    #rename_must_contain_entry = tk.Entry(master)
    #rename_must_contain_entry.config(width=20, font=('Helvetica', 13), borderwidth=2)
    #rename_must_contain_entry.insert(0, '--Must*Contain--')
    #rename_must_contain_entry.grid(row=12, column=2)

    rename_ico = ImageTk.PhotoImage(Image.open("IconPictures/compare_zs_63.png"))
    button_rename = tk.Button(image=rename_ico, command=compare_master_model)
    button_rename.grid(row=13, column=1)

    open_source_ico = ImageTk.PhotoImage(Image.open("IconPictures/source_folder.png"))
    open_source_button = tk.Button(image=open_source_ico, command=open_database_folder)
    open_source_button.grid(row=14, column=1)

    run_ico = ImageTk.PhotoImage(Image.open("IconPictures/robot_assemble.png"))
    global run_button
    run_button = tk.Button(image=run_ico, command=assembling_procces)
    run_button.config(state='disabled')
    run_button.grid(row=13, column=0)

    quit_ico = ImageTk.PhotoImage(Image.open("IconPictures/quit_icon.png"))
    quit_button = tk.Button(image=quit_ico, command=close_graphical_user_interface)
    quit_button.grid(row=14, column=2)

    reset_ico = ImageTk.PhotoImage(Image.open("IconPictures/reset.png"))
    reset_button = tk.Button(image=reset_ico, command=reset_graphical_user_interface)
    reset_button.grid(row=13, column=2)

    global feedback_button
    feedback_ico = ImageTk.PhotoImage(Image.open("IconPictures/feedback_folder.png"))
    feedback_button = tk.Button(image=feedback_ico)
    feedback_button.grid(row=14, column=0)

    list_gm_folders_function()
    master.mainloop()


def list_gm_folders_function():

    """This function looks into database excel file and determines Machine types of Injection unit machines according to names of sheets."""

    global list_gm_types;
    list_gm_types = []
    global database_path;
    database_path = 'DatabaseFolder\\mastermodels_database.xlsx'
    global picked_type
    global types_option

    input_workbook = xlrd.open_workbook(database_path)
    for sheet in input_workbook.sheets():
        list_gm_types.append(sheet.name)

    picked_type = tk.StringVar(master)
    picked_type.set(list_gm_types[0])  # initial value

    types_option = tk.OptionMenu(master, picked_type, *list_gm_types)
    types_option.config(height=1, width=35, font=('Helvetica 9 bold'))
    types_option.grid(row=2, column=0, rowspan=2, columnspan=2)


def list_CAD_mastermodels():

    """This function lists all mastermodels (CAD names), according to machine type picked by user."""

    global list_mastermodels;
    list_mastermodels = []
    global picked_gm
    global list_mastermodels_option
    global positions_mastermodels;
    positions_mastermodels = []
    global input_worksheet

    types_option.config(state='disabled')
    picked_type_confirmation_button.config(state='disabled')

    input_workbook = xlrd.open_workbook(database_path)
    input_worksheet = input_workbook.sheet_by_name(picked_type.get())

    for row_value in range(input_worksheet.nrows):
        if input_worksheet.cell_value(row_value, 0) != '' and input_worksheet.cell_value(row_value, 0) != 'CAD mastermodel name':
            list_mastermodels.append(input_worksheet.cell_value(row_value, 0))
            position_mastermodel = {'CAD_name': input_worksheet.cell_value(row_value, 0), 'rows_start': row_value}
            positions_mastermodels.append(position_mastermodel.copy())

    picked_gm = tk.StringVar(master)
    picked_gm.set(list_mastermodels[0])  # initial value

    list_mastermodels_option = tk.OptionMenu(master, picked_gm, *list_mastermodels)
    list_mastermodels_option.config(height=1, width=35, font=('Helvetica 9 bold'))
    list_mastermodels_option.grid(row=3, column=0, rowspan=2, columnspan=2)

    global confirm_picked_machine_size
    confirm_picked_machine_size = tk.Button(image=check_mark_ico, command=list_properties_function)
    confirm_picked_machine_size.grid(row=3, column=2, rowspan=2)


def list_properties_function():

    """Following function determines properties of CAD model according to sheet."""

    global source_assembly_name;
    source_assembly_name = picked_gm.get()
    global list_clamp_sizes;
    list_clamp_sizes = []
    global list_powerpacks;
    list_powerpacks = []
    global list_primary_plast;
    list_primary_plast = []
    global list_secondary_plast;
    list_secondary_plast = []
    global special_sign;
    special_sign = None
    global list_second_powerpacks
    list_second_powerpacks = []
    properties = [list_clamp_sizes, list_powerpacks, list_primary_plast, list_secondary_plast, special_sign, list_second_powerpacks]

    list_mastermodels_option.config(state='disabled')
    confirm_picked_machine_size.config(state='disabled')
    run_button.config(state='normal')

    # extract picked master_model range
    for position_mastermodel in positions_mastermodels:
        current_index = positions_mastermodels.index(position_mastermodel)
        try:
            upcoming_dict = positions_mastermodels[current_index + 1]
        except IndexError:
            rows_finish = position_mastermodel['rows_start'] + 5
        else:
            rows_finish = upcoming_dict['rows_start'] - 1
        finally:
            position_mastermodel["rows_finish"] = rows_finish

    for each_dictionary in positions_mastermodels:
        if each_dictionary['CAD_name'] == picked_gm.get():
            start_range = each_dictionary['rows_start']
            end_range = each_dictionary['rows_finish'] + 1
            working_range_CAD_master = range(start_range, end_range)
            break

    # Now we have picked range ! so we can create lists of all CAD_master_model_parameters
    for row_value in working_range_CAD_master:
        # Creating clamp_units:
        if input_worksheet.cell_value(row_value, 1):
            list_clamp_sizes.append(input_worksheet.cell_value(row_value, 1))
        # Creating powerpacks:
        if input_worksheet.cell_value(row_value, 2):
            list_powerpacks.append(input_worksheet.cell_value(row_value, 2))
        # Creating Primary_plast_options:
        if input_worksheet.cell_value(row_value, 3):
            list_primary_plast.append(input_worksheet.cell_value(row_value, 3))
        # Creating Secondary_plast options:
        if input_worksheet.cell_value(row_value, 4):
            list_secondary_plast.append(input_worksheet.cell_value(row_value, 4))
        # Setting special sign
        if input_worksheet.cell_value(row_value, 5):
            special_sign = input_worksheet.cell_value(row_value, 5)
        # Creating list of second powerpacks - suited for GXL machines
        if input_worksheet.cell_value(row_value, 6):
            list_second_powerpacks.append(input_worksheet.cell_value(row_value, 6))


        for mastermodel_property in properties:
            if mastermodel_property == [] or mastermodel_property == None:
                properties.remove(mastermodel_property)
    # Very well, here we have determined which properties our master model has ! - probably not needed
    print(properties)
    # Here we go lists all properties we need if more than two options exist - list box window is relevant
    # First is clamp size
    if len(list_clamp_sizes) > 1:
        global picked_clamp_size
        global clamp_sizes_option
        row_grid = 4
        picked_clamp_size = tk.StringVar(master)
        picked_clamp_size.set(list_clamp_sizes[0])  # initial value
        clamp_sizes_option = tk.OptionMenu(master, picked_clamp_size, *list_clamp_sizes)
        clamp_sizes_option.config(height=1, width=35, font=('Helvetica 10 bold'))
        clamp_sizes_option.grid(row=row_grid, column=0, rowspan=2, columnspan=2)
        clamp_sizes_label = tk.Label(master, text="Clamp sizes", font=('Helvetica 10 bold'))
        clamp_sizes_label.grid(row=row_grid, column=1, rowspan=2, columnspan=2)

    if len(list_powerpacks) > 1:
        global picked_power_pack
        global picked_power_pack_option
        row_grid = 5
        picked_power_pack = tk.StringVar(master)
        picked_power_pack.set(list_powerpacks[0])  # initial value
        picked_power_pack_option = tk.OptionMenu(master, picked_power_pack, *list_powerpacks)
        picked_power_pack_option.config(height=1, width=35, font=('Helvetica 10 bold'))
        picked_power_pack_option.grid(row=row_grid, column=0, rowspan=2, columnspan=2)
        pp_sizes_label = tk.Label(master, text="PowerPack", font=('Helvetica 10 bold'))
        pp_sizes_label.grid(row=row_grid, column=2, rowspan=2, columnspan=2)

    if len(list_primary_plast) > 1:
        global picked_primary_plastification
        global main_plastification_option
        row_grid = 6
        picked_primary_plastification = tk.StringVar(master)
        picked_primary_plastification.set(list_primary_plast[0])  # initial value
        main_plastification_option = tk.OptionMenu(master, picked_primary_plastification, *list_primary_plast)
        main_plastification_option.config(height=1, width=35, font=('Helvetica 10 bold'))
        main_plastification_option.grid(row=row_grid, column=0, rowspan=2, columnspan=2)
        pl1_sizes_label = tk.Label(master, text="Plastification 1", font=('Helvetica 10 bold'))
        pl1_sizes_label.grid(row=row_grid, column=2, rowspan=2, columnspan=2)

    if len(list_secondary_plast) > 1:
        global picked_secondary_plastification
        global secondary_plast_option
        row_grid = 7
        picked_secondary_plastification = tk.StringVar(master)
        picked_secondary_plastification.set(list_secondary_plast[0])  # initial value
        secondary_plast_option = tk.OptionMenu(master, picked_secondary_plastification, *list_secondary_plast)
        secondary_plast_option.config(height=1, width=35, font=('Helvetica 10 bold'))
        secondary_plast_option.grid(row=row_grid, column=0, rowspan=2, columnspan=2)
        pl2_sizes_label = tk.Label(master, text="Plastification 2", font=('Helvetica 10 bold'))
        pl2_sizes_label.grid(row=row_grid, column=2, rowspan=2, columnspan=2)

    if len(list_second_powerpacks) > 1:
        global picked_second_powerpack
        global picked_second_powerpack_option
        row_grid = 8
        picked_second_powerpack = tk.StringVar(master)
        picked_second_powerpack.set(list_second_powerpacks[0])  # initial value
        picked_second_powerpack_option = tk.OptionMenu(master, picked_second_powerpack, *list_second_powerpacks)
        picked_second_powerpack_option.config(height=1, width=35, font=('Helvetica 9 bold'))
        picked_second_powerpack_option.grid(row=row_grid, column=0, rowspan=2, columnspan=2)
        pp2_sizes_label = tk.Label(master, text="PowerPack 2", font=('Helvetica 10 bold'))
        pp2_sizes_label.grid(row=row_grid, column=2, rowspan=2, columnspan=2)


def reset_graphical_user_interface():
    """Button reaction function"""
    close_graphical_user_interface()
    build_graphical_user_interface()


def close_graphical_user_interface():
    """Button reaction function"""
    master.destroy()


def open_database_folder():
    """Button reaction function"""
    database_folder_path = '.\DatabaseFolder'
    os.startfile(database_folder_path)


def open_feedback_folder():
    """Button reaction function"""
    feedback_folder_path = '.\FeedbackFolder'
    os.startfile(feedback_folder_path)


def open_log_file():
    """Button reaction function"""
    log_file = 'assembly_automation.log'
    os.startfile(log_file)


# -----A------PPPP---I---
# ----A-A-----P---P--I---
# ---A---A----P--P---I---
# --AAAAAAA---PPP----I---
# -A-------A--P------I---
# A---------A-P------I---

# -------MAIN_FUNCTION-------
def assembling_procces():

    """ This function covers whole process of assembling. All the steps of this essential function are described below in comments."""

    try:
        in_progress_icon = ImageTk.PhotoImage(Image.open("IconPictures/in_progress.png"))
        error_icon = ImageTk.PhotoImage(Image.open("IconPictures/error_icon.png"))
        successful_finish_icon = ImageTk.PhotoImage(Image.open("IconPictures/green_check_mark_icon.png"))

        app_status_label = tk.Label(image=in_progress_icon)
        app_status_label.grid(row=11, column=2)
        skip_preparation = False

        if len(order_number_entry.get()) != 6:
            yes_no_preparation = tk.messagebox.askquestion('Invalid Order Number', 'Order number is not valid. Do you want to continue without preparation of mastermodel?'
                                                                                   ' Currently opened mastermodel will be reference for procces of automation.'
                                                                                    ' This procces will delete assemblies and assemble components from loaded ZS63 text file.', icon='warning')
            if yes_no_preparation:
                skip_preparation = True
            else:
                exit()
        # 1 function loads ZS_63 - this function is must in all cases
        load_zs_63()
        # 2 preparation function - depends on order number validity
        if skip_preparation == False:
            preparation_master_model()
        # 3 Now we procces and pair CAD with ERP system
        read_ZS63_pair_with_CAD(pair_also=True)
        # 4 Now we want to make our Master model lighter and faster. Therefore we delete all necessary ERP numbers out of models (from source model
        # Advantage of this function is this saves model with lower quality name for example 2000575_SP1400.prt will be not deleted if such ERP number
        # is found in ZS63
        deleting_models()
        # 5 Because special assemblies are not usually part of Master model we create them according to result of previous function.
        create_sa_groups()
        # 6 This may look like like duplicity, however since we have special assemblies in our master model we can also pair them.
        read_ZS63_pair_with_CAD(pair_also=True)
        # 7 Assembling is based on our global dictionary named all_lists. This dictionary stores three key values ERP number - CAD group name - SAP group name
        # Of course we will assemble models that we can assemble !
        for every_erp_sapname_cadgroup in all_lists:
            ERP_material_number = every_erp_sapname_cadgroup['ERP_number']
            cad_parent_model = every_erp_sapname_cadgroup['CAD_group_name']
            if cad_parent_model != 'Not defined':
                assemble_model(ERP_material_number, cad_parent_model)
        # 8 Here is our feedback function - provides feedback about non - paired models which exist ! Therefore they deserve to be placed in to master model !
        existing_nonassembled_models_feedback()
        creopyson.file_open(creo_client, file_=current_master_model)
        creopyson.file_save(creo_client, file_=current_master_model)
        set_default_view()
        tk.messagebox.showinfo('Automation status', 'Automation completed !')
        app_status_label.photo = successful_finish_icon
    except:

        app_status_label.destroy()
        error_status_button = tk.Button(master, image=error_icon, command = open_log_file)
        error_status_button.photo = error_icon
        error_status_button.grid(column=2, row=11)
        print("there should be icon")
        sys.exc_info()
        logger.exception("message")

        try:
            rename_config_control("no")
        except:
            pass


def get_session_information():

    global gmXXXX
    global current_master_model
    get_session_info = creopyson.file_get_fileinfo(creo_client)
    current_master_model = (get_session_info['file'])
    gmXXXX = current_master_model[(len(current_master_model)) - 10:(len(current_master_model)) - 4]


def open_material_number(ERP_material_number):
    """This function tests whether ERP material number exists in Windchill and if exists it will assign its modelname to model_name variable"""

    global model_name
    model_name = ''
    try:
        creopyson.file.open_(creo_client, file_=(ERP_material_number + '.PRT'), display=False)
    except RuntimeError as error:
        pass
        print('Part material number ' + ERP_material_number + ' does not exist !')
    else:
        creopyson.file.close_window(creo_client, file_=(ERP_material_number + '.PRT'))
        model_name = ERP_material_number + '.PRT'
        print('Yes material number exists ! Model name is ' + model_name)
        logger.info('Yes material number exists ! Model name is ' + model_name)
    try:
        creopyson.file.open_(creo_client, file_=(ERP_material_number + '.ASM'), display=False)
    except RuntimeError as error:
        pass
        print('Assembly material number ' + ERP_material_number + ' does not exist !')
    else:
        creopyson.file.close_window(creo_client, file_=(ERP_material_number + '.ASM'))
        model_name = ERP_material_number + '.ASM'
        print('Yes material number exists ! Model name is ' + model_name)
        logger.info('Yes material number exists ! Model name is ' + model_name)


def creoson_setup():

    """This function connects this application to PTC Creo session via Creoson"""

    global creo_client
    creo_client = creopyson.Client()
    try:
        creo_client.connect()
        print('Creoson is running')
        logger.info('Creoson is running')
    except ConnectionError:
        creoson_folder = os.path.dirname(sys.argv[0]) + '\\creoson'
        os.startfile(creoson_folder)
        tk.messagebox.showinfo("Kraussmaffei Assembly Automation", "Creoson is not running. Start Creoson before starting Automation app.")
        logger.critical("Kraussmaffei Assembly Automation", "Creoson is not running. Start Creoson before starting Automation app.")
        exit()


def bom_recursion(nest_dict, list_of_recursed_bom=[]):

    for key, value in nest_dict.items():
        if isinstance(value, dict):
            bom_recursion(value)
        elif isinstance(value, list):
            for each in value:
                if isinstance(each, dict):
                    bom_recursion(each)
        else:
            list_of_recursed_bom = list_of_recursed_bom
            if key != 'generic':
                list_of_recursed_bom.append(("{0} : {1}".format(key, value)))

    return list_of_recursed_bom


def list_gm_groups(exclude_PRT_files=False, first_level_only=False):

    """This function lists all asm and prt files which shares gmXXXX or any other six numbers in the session, while of model tree are 3 levels are looped."""

    global list_components
    global current_master_model
    global order_number

    list_components = []
    get_session_info = creopyson.file_get_fileinfo(creo_client)
    current_master_model = (get_session_info['file'])
    gmXXXX = current_master_model[(len(current_master_model)) - 10:(len(current_master_model)) - 4]
    order_number = gmXXXX
    'In case we do not want to have .prt files (like skeletons) in our list - its for assembly purpose since we can not assembly into .prt models.'
    if exclude_PRT_files == True:
        gmXXXX = gmXXXX + '.asm'
    list_master_comp = creopyson.feature_list(creo_client, name='*' + gmXXXX + '*', file_=current_master_model, no_datum=True, type_='COMPONENT')
    list_components.append(current_master_model)
    'Author note: this could be recursed.'
    for unit in list_master_comp:
        list_components.append(unit['name'].lower())
        sub_comp = unit['name']
        list_sub_components = creopyson.feature_list(creo_client, name='*' + gmXXXX + '*', file_=sub_comp, no_datum=True, type_='COMPONENT')
        if first_level_only == False:
            for each_sub in list_sub_components:
                list_components.append(each_sub['name'].lower())
                sub_sub_component = each_sub['name']
                list_sub_sub_components = creopyson.feature_list(creo_client, name='*' + gmXXXX + '*', file_=sub_sub_component, no_datum=True, type_='COMPONENT')
                for each_sub_sub in list_sub_sub_components:
                    list_components.append(each_sub_sub['name'].lower())


def assemble_model(ERP_material_number, cad_parent_model):

    """This function assembles material number into injection_machine group"""

    # At first we determine whether ERP_model_exists by open material number function.
    open_material_number(ERP_material_number)
    # Then we can proceed with assemble model function
    if model_name != '':
        # Now we test whether model already exists in cad_parent_model. If yes we skip this step to avoid cad model duplicity:
        skip = False
        list_components_in_assy_raw = creopyson.feature_list(creo_client, file_=cad_parent_model, no_datum=True, type_='COMPONENT')
        list_components_in_assy = []
        for names in list_components_in_assy_raw:
            list_components_in_assy.append(names['name'])
        for component_name in list_components_in_assy:
            if model_name.upper() == component_name.upper():
                skip = True
                logger.info("Model "+component_name+" was not assembled in "+cad_parent_model+" due to its existence in origin model.")
                open_model_and_resume_all_groups(component_name=cad_parent_model, gmXXXX_only=False)
        if skip == False:
            pick_csy(cad_parent_model, model_name)
            child_csy = picked_csy
            pick_csy(cad_parent_model, cad_parent_model)
            parent_csy = picked_csy
            if parent_csy != 'CSY does not exist' and child_csy != 'CSY does not exist':
                creopyson.file_assemble(creo_client, into_asm=cad_parent_model, file_=model_name, constraints=[{"asmref": parent_csy, "compref": child_csy, "type": "csys"}])
            if parent_csy == 'CSY does not exist' and child_csy != 'CSY does not exist':
                #Check whether skeleton exists function might be enhanced.
                check_whether_skeleton_exists(cad_parent_model)
                if skeleton_csy != 'CSY does not exist':
                    creopyson.file_assemble(creo_client, file_=model_name, into_asm=cad_parent_model, ref_model=skeleton_name, constraints=[{"asmref": skeleton_csy, "compref": child_csy, "type": "csys"}])
                else:
                    creopyson.file_assemble(creo_client, into_asm=cad_parent_model, file_=model_name, constraints=[{"type": "fix"}], package_assembly=True)
            else:
                creopyson.file_assemble(creo_client, into_asm=cad_parent_model, file_=model_name, package_assembly=True)


def pick_csy(cad_parent_model, checked_model):

    """This function determines coordinate systems for assembling purposes."""

    global picked_csy
    picked_csy = ""
    max_ratio = 0
    list_csys_model_raw = creopyson.feature_list(creo_client, file_=checked_model, type_='COORDINATE SYSTEM', no_comp=False)
    list_csys = []

    for every_dict in list_csys_model_raw:
        list_csys.append(every_dict['name'])
    # Test amount of csys in model
    # At first we test whether there is more than one csys. If so Sequence matcher function is used to determine correct CSY.
    if len(list_csys) > 1:
        for every_csy in list_csys:
            measure_similarity = SequenceMatcher(None, cad_parent_model, every_csy)
            every_csy_ratio = measure_similarity.ratio()
            if every_csy_ratio > max_ratio:
                max_ratio = every_csy_ratio
                best_match_csy = every_csy
        # There also has to be added option when ratio equals zero for all compared csys
        if max_ratio != 0:
            picked_csy = best_match_csy
        else:
            picked_csy = list_csys[0]
    # Then we test if number of csys is only one. If so this csy is assigned to picked_csy
    if len(list_csys) == 1:
        picked_csy = list_csys[0]
    # At last there is possibility of when there is none csys in model. In this case following string is assigned to picked_Csy
    if len(list_csys) == 0:
        picked_csy = 'CSY does not exist'

    try:
        # 23-06-2020 this section will add possible coordinate system for MX machine
        # Requires more development
        if len(list_csys) > 1 and zs_63_injection_unit != "":
            range_csys = []
            # These coordinates system may have their own range of numbers. The first condition of these ,,range csys,, is that they contain ,,SP,,
            for each_csy in list_csys:
                if "SP" in each_csy.upper():
                    range_csys.append(each_csy)
            # At this part we can determine whether range_csys has multiple elements. If yes, we create objects of csys where we determine several properties
            mx_csys_objects = []
            if len(range_csys) > 1:
                for each_csy in range_csys:
                    sp_pos = each_csy.index("SP")
                    range_in_csy_string = each_csy[sp_pos+2:]
                    range_in_csy_list = range_in_csy_string.split("-")
                    boundering_values = []
                    for each_range_value in range_in_csy_list:
                        boundering_values.append(only_numerics(each_range_value))
                current_mx_csy_object = {
                    "name": each_csy,
                    "floor": min(boundering_values),
                    "ceil": max(boundering_values)
                }
                mx_csys_objects.append(current_mx_csy_object)
                print(current_mx_csy_object)
                print("Above is csy object")
            if len(mx_csys_objects)>1:
                for each_csy_object in mx_csys_objects:
                    if zs_63_injection_unit >= each_csy_object["floor"] and zs_63_injection_unit <= each_csy_object["ceil"]:
                        picked_csy = each_csy_object["name"]
                        logger.info("MX variation of picked CSY is "+picked_csy)
    except:
        logger.exception("message")


def check_whether_skeleton_exists(cad_parent_model):

    """This function determines whether skeleton model exists. If this model exists it looks for its csys and picks the best matching"""

    global skeleton_name
    global skeleton_csy

    skeleton_comps = creopyson.feature_list(creo_client, name='*' + 'SKEL' + '*', file_=cad_parent_model, no_datum=True, type_='COMPONENT')
    list_skel_comp = []
    for every_SKEL in skeleton_comps:
        list_skel_comp.append(every_SKEL['name'])
    if len(list_skel_comp) == 1:
        skeleton_name = list_skel_comp[0]
    pick_csy(cad_parent_model, skeleton_name)
    skeleton_csy = picked_csy


def read_ZS63_pair_with_CAD(pair_also=False, test=False):

    """This function loads ZS_63 Bom and converts it to list of dictionaries. Dictionary contains information of CAD name group, SAP group and Material Number"""

    global all_lists
    all_lists = []
    m_groups_list = []
    ze_groups_list = []
    global sa_groups_list
    sa_groups_list = []
    zs_63_raw = []
    zs_63 = []
    global special_sign


    if test == True:
        pass
        #special_sign = None
        #special_sign=False

    with open('ZS_63_source_folder\\zs_63.txt') as zs_data:

        zs_63_raw = zs_data.readlines()

        # This is encoding part - neccesary to implement because of various operating systems
        for each_line in zs_63_raw:
            each_line.encode("utf-8", "ignore")
            line_text = str(each_line.encode("utf-8", "ignore"))
            line_text = line_text.replace("b'", "")
            zs_63.append(str(line_text))

        # Newly create section where some risky signs will be removed from zs_63.txt
        if type(zs_63) == list:
            for each_element in zs_63:
                try:
                    if "°C" in each_element:
                        zs_63.remove(each_element)
                        print(each_element+" is removed as banned symbol.")
                except:
                        logger.info("banned symbol was not possible to remove ")

            for each_element in zs_63:
                try:
                    if "?" in each_element:
                        zs_63[zs_63.index(each_element)] = each_element.replace("?", " ")
                        print(each_element+" is removed as banned symbol ?")
                except:
                    pass
            for each_element in zs_63:
                try:
                    if "#" in each_element and each_element[2]!="#":
                        zs_63[zs_63.index(each_element)] = each_element.replace("#", " ")
                        print(each_element+" is removed as banned symbol #")
                except:
                    pass

        # Here we try to set injection unit size for MX
        try:
            global zs_63_injection_unit; zs_63_injection_unit = ""
            if "MX" in current_master_model.upper():
                following_line = False
                for line in zs_63:
                    if following_line == True:
                        list_of_splited_line = line.split("/")
                        zs_63_injection_unit_raw = list_of_splited_line[1]
                        zs_63_injection_unit = zs_63_injection_unit_raw
                        for each_letter in zs_63_injection_unit_raw:
                            if each_letter == "0":
                                zs_63_injection_unit = zs_63_injection_unit[1:]
                            else:
                                break
                        print(zs_63_injection_unit)
                        break
                    if "--------------------------------------------------------" in line:
                        following_line = True
                logger.info('Machine type is MX therefore  program tried to set injection unit according to ZS_63, the value is ' + zs_63_injection_unit)
        except:
            logger.exception("message")

        for line in zs_63:
            first_symbols = line[0:3]
            line = line.replace(' 000  ', '      ')
            if 'GM' or 'PL' in first_symbols:
                m_groups_list.append(line)
            if 'ZE#' in first_symbols:
                ze_groups_list.append(line)
            if 'SA' in first_symbols:
                sa_groups_list.append(line)

        for every_element in m_groups_list:
            split_element = every_element.split()
            pair_group_and_number = []
            for each_split_element in split_element:
                each_split_element = each_split_element.strip()
                if len(each_split_element) > 6 and len(each_split_element) < 9 and '7' not in each_split_element[0:1] and each_split_element.isnumeric():
                    pair_group_and_number.append(each_split_element)
                if 'M' in each_split_element and '.' in each_split_element or 'C' in each_split_element and '.' in each_split_element:
                    pair_group_and_number.append(each_split_element)
            if len(pair_group_and_number) == 2:
                mat_nr = pair_group_and_number[0]
                sap_group = pair_group_and_number[1]
                group_mat_nr_dict = {'SAP_group_name': sap_group, 'ERP_number': mat_nr, 'CAD_group_name': 'Not defined'}
                all_lists.append(group_mat_nr_dict.copy())

        for every_element in ze_groups_list:
            split_element = every_element.split()
            pair_group_and_number = []
            for each_split_element in split_element:
                each_split_element = each_split_element.strip()
                if len(each_split_element) > 6 and len(each_split_element) < 9 and '7' not in each_split_element[0:1] and each_split_element.isnumeric():
                    pair_group_and_number.append(each_split_element)
                if '.' in each_split_element:
                    try:
                        float(each_split_element)
                    except ValueError:
                        pass
                    else:
                        pair_group_and_number.append('ZE' + each_split_element)
            if len(pair_group_and_number) == 2:
                mat_nr = pair_group_and_number[0]
                sap_group = pair_group_and_number[1]
                group_mat_nr_dict = {'SAP_group_name': sap_group, 'ERP_number': mat_nr, 'CAD_group_name': 'Not defined'}
                all_lists.append(group_mat_nr_dict.copy())

        for every_element in sa_groups_list:
            split_element = every_element.split()
            pair_group_and_number = []
            for each_split_element in split_element:
                each_split_element = each_split_element.strip()
                if len(each_split_element) < 4 and len(each_split_element) > 1 and each_split_element.isnumeric():
                    SAValue = int(each_split_element)
                    if SAValue > 29 and SAValue < 300:
                        if len(each_split_element) == 2:
                            each_split_element = 'SA0' + each_split_element
                        if len(each_split_element) == 3:
                            each_split_element = 'SA' + each_split_element
                        pair_group_and_number.append(each_split_element)
                if len(each_split_element) > 6 and len(each_split_element) < 9 and '7' not in each_split_element[0:1] and each_split_element.isnumeric():
                    pair_group_and_number.append(each_split_element)
            if len(pair_group_and_number) == 2:
                mat_nr = pair_group_and_number[1]
                sap_group = pair_group_and_number[0]
                group_mat_nr_dict = {'SAP_group_name': sap_group, 'ERP_number': mat_nr, 'CAD_group_name': 'Not defined'}
                all_lists.append(group_mat_nr_dict.copy())

        if pair_also == True:

            set_model_convention_on_the_fly()

            list_gm_groups()

            if special_sign != None:
                m6_list = []
                previous_dict = {}
                # count how many times M6_M01 occures in all lists
                for every_dict in all_lists:
                    if 'M6' in every_dict['SAP_group_name']:
                        if not any(x == every_dict['SAP_group_name'] for x in m6_list):
                            m6_list.append(every_dict['SAP_group_name'])
                        else:
                            if every_dict['SAP_group_name'] != previous_dict:
                                every_dict['SAP_group_name'] = special_sign + every_dict['SAP_group_name']
                                previous_dict = every_dict['SAP_group_name']

            for every_dict in all_lists:

                # Here we are going to generate possible names

                # 1st quality is the name which equals to the name of ['SAP_group_name'] for example M6.E30, ZE07.12 etc.
                name_quality_level_1 = every_dict['SAP_group_name']

                # 2nd quality replaces dod with underscore M6_E30, M22_M20_2 etc.
                name_quality_level_2 = name_quality_level_1.replace('.', '_')

                # 3rd quality takes level 2 and second underscore replaces with dash M22_M20-2
                name_quality_level_2_split = name_quality_level_2.split('_')
                name_quality_level_3 = ""
                if len(name_quality_level_2_split) == 3:
                    name_quality_level_3 = name_quality_level_2_split[0] + '_' + name_quality_level_2_split[1] + '-' + name_quality_level_2_split[2]

                # 4th quality is specialized for ZE machines - especially for the naming of pads. In CAD model is no Zero on third position ZE02_50
                name_quality_level_4 = ""
                if name_quality_level_2[2] == "0":
                    name_quality_level_4 = name_quality_level_2[0 : 2 : ] + name_quality_level_2[2 + 1 : :]

                # 5th quality adds X to ending of group name
                name_quality_level_5 = ""
                try:
                    if name_quality_level_2[len(name_quality_level_2)-1] != ('_'):
                        name_quality_level_5 = name_quality_level_2[0:len(name_quality_level_2) - 1] + 'X'
                except IndexError:
                    pass

                # 6th quality removes last and adds X to ending of group name
                name_quality_level_6 = ""
                try:
                    if name_quality_level_5[len(name_quality_level_5)-1] != ("_"):
                        name_quality_level_6 = name_quality_level_5[0:len(name_quality_level_5) - 2] + 'X'
                except IndexError:
                    pass

                # 7th quality takes string changes dot to dash ZE25.50 -> ZE25-50
                name_quality_level_7  = name_quality_level_1.replace('.', '-')

                # 8th quality uses level 7 and removes 0 from third index
                name_quality_level_8 = ""
                if name_quality_level_7[2] == "0":
                    name_quality_level_8 = name_quality_level_7[0 : 2 : ] + name_quality_level_7[2 + 1::]

                # 9th level takes another number from string (if it is not bordering)
                name_quality_level_9 = name_quality_level_5.replace('.', '-')

                # 10th level takes another number from string (if it is not bordering)
                name_quality_level_10 = name_quality_level_6.replace('.', '-')

                # 11th quality takes level 2 and first underscore removes E_M5_M50 -> EM5_M50
                name_quality_level_11 = ""
                if len(name_quality_level_2_split) == 3:
                    name_quality_level_11 = name_quality_level_2_split[0] + '' + name_quality_level_2_split[1] + '_' + name_quality_level_2_split[2]

                # At this point we can build list
                all_tested_names = [name_quality_level_1, name_quality_level_2, name_quality_level_3,
                                    name_quality_level_4, name_quality_level_5, name_quality_level_6,
                                    name_quality_level_7, name_quality_level_8, name_quality_level_9,
                                    name_quality_level_10, name_quality_level_11]

                # Deletion of empty elements of the list
                for every_tested_name in all_tested_names:
                    if every_tested_name == "":
                        all_tested_names.remove(every_tested_name)

                # Finally the pairing it self. Looped throughout every possible name (for specific SAP name) and every CAD group.
                for every_name in all_tested_names:
                    for every_cad_group in list_components:
                        if '_' + every_name.lower() + '_' in every_cad_group:
                            print(name_quality_level_1 + ' paired with ' + every_cad_group)
                            logger.info(name_quality_level_1 + ' paired with ' + every_cad_group)
                            every_dict['CAD_group_name'] = every_cad_group
                            break


            # Short computation to determine quality of pairing process. Output is logged.
            succesfull_pairing = 0
            for each_dict in all_lists:
                print(each_dict)
                if each_dict["CAD_group_name"] != "Not defined":
                    succesfull_pairing=succesfull_pairing+1

            percentage = succesfull_pairing / len(all_lists) * 100
            logger.info("Percentage of defined pairs is "+str(percentage) + "%.")
            print("Percentage of defined pairs is "+str(percentage) + "%.")
            logger.info('end of pairing')


def create_new_copy():

    """This function creates new copy of master model. Old script allow (with mapkey) is stored in this function if the update of this lines will get wrong."""

    list_gm_groups()
    new_number = order_number_entry.get()

    script_allow_rename = "Select `main_dlg_cur` `appl_casc`;~ Close `main_dlg_cur` `appl_casc`;\~ Command `ProCmdRibbonOptionsDlg` ;\~ Select `ribbon_options_dialog` `PageSwitcherPageList` 1 `ConfigLayout`;\~ Trail `UI Desktop` `UI Desktop` `PREVIEW_POPUP_TIMER` `main_dlg_w1:PHTLeft.AssyTree:<NULL>`;\~ Activate `main_dlg_cur` `main_dlg_cur`;\~ Activate `main_dlg_cur` `main_dlg_cur`;\~ Activate `ribbon_options_dialog` `ConfigLayout.AddOpt`;\~ Input `add_opt` `InputOpt` `let_proe_rename_pdm_objects`;\~ Update `add_opt` `InputOpt` `let_proe_rename_pdm_objects`;\~ FocusOut `add_opt` `InputOpt`;~ Open `add_opt` `EditPanel`;\~ Close `add_opt` `EditPanel`;~ Select `add_opt` `EditPanel` 1 `yes`;\~ FocusOut `add_opt` `EditPanel`;~ Activate `add_opt` `AddOpt`;\~ Activate `ribbon_options_dialog` `OkPshBtn`;\~ FocusIn `UITools Msg Dialog Future` `no`;\~ Activate `UITools Msg Dialog Future` `no`;"
    #creopyson.interface_mapkey(creo_client, script_allow_rename)
    rename_config_control("yes")
    for every_component in list_components:
        new_name = every_component.replace(order_number, new_number)
        try:
            creopyson.file_rename(creo_client, file_=every_component, new_name=new_name, onlysession=True)
            print('this is new model = ' + new_name)
            logger.info('this is new model = ' + new_name)
            #creopyson.file_save(creo_client, file_=new_name)
        except:
            logger.exception("message")
    script_ban_rename = "Select `main_dlg_cur` `appl_casc`;~ Close `main_dlg_cur` `appl_casc`;\~ Command `ProCmdRibbonOptionsDlg` ;\~ Select `ribbon_options_dialog` `PageSwitcherPageList` 1 `ConfigLayout`;\~ Activate `ribbon_options_dialog` `ConfigLayout.AddOpt`;\~ Input `add_opt` `InputOpt` `let_proe_rename_pdm_objects`;\~ Update `add_opt` `InputOpt` `let_proe_rename_pdm_objects`;\~ FocusOut `add_opt` `InputOpt`;~ Activate `add_opt` `AddOpt`;\~ Activate `ribbon_options_dialog` `OkPshBtn`;\~ FocusIn `UITools Msg Dialog Future` `no`;\~ Activate `UITools Msg Dialog Future` `no`;"
    #creopyson.interface_mapkey(creo_client, script_ban_rename)
    #creopyson.file_save(creo_client)
    rename_config_control("no")


def deleting_models():

    folder_name = picked_type.get()
    Delete_exclude_file = 'DeleteExclude\\' + folder_name + '\\DeleteExclude.csv'
    with open(Delete_exclude_file, 'r') as csv_file:
        csv_reader = csv.reader(csv_file)
        delete_exclude_list = []
        for line in csv_reader:
            if line[0] != "":
                delete_exclude_list.append(line[0])
    print('These assemblies wont be deleted')
    logger.info('These assemblies wont be deleted')
    print(delete_exclude_list)
    logger.info(delete_exclude_list)

    list_gm_groups(exclude_PRT_files=False)

    read_ZS63_pair_with_CAD(pair_also=True)

    erp_numbers = []

    for every_dict in all_lists:
        erp_numbers.append(every_dict['ERP_number'])

    print(erp_numbers)

    print(list_components)

    for every_assembly in list_components:
        list_components_in_assy = []
        if any(ele.upper() in every_assembly for ele in delete_exclude_list) or any(ele.lower() in every_assembly for ele in delete_exclude_list):
            print('this group is avoided ' + every_assembly)
            logger.info('this group is avoided ' + every_assembly)
        else:
            try:
                creopyson.file_open(creo_client, file_=every_assembly)
                list_components_in_assy_raw = creopyson.feature_list(creo_client, file_=every_assembly, no_datum=True, type_='COMPONENT')
                for names in list_components_in_assy_raw:
                    list_components_in_assy.append(names['name'])
                for every_listed_comp in list_components_in_assy:
                    in_list_components = False
                    for every_erp_number in erp_numbers:
                        # print(every_erp_number +'  '+ every_listed_comp)
                        if every_erp_number in every_listed_comp:
                            print('this number is going to be avoided ' + every_erp_number)
                            logger.info(('this number is going to be avoided ' + every_erp_number))
                            in_list_components = True
                    if order_number.lower() in every_listed_comp or order_number.upper() in every_listed_comp or 'SKEL' in every_listed_comp or in_list_components == True:
                        pass
                    else:
                        creopyson.feature_delete(creo_client, file_=every_assembly, name=every_listed_comp)
                        print('model is deleted ' + every_listed_comp + ' in ' + every_assembly)
                        logger.info('model is deleted ' + every_listed_comp + ' in ' + every_assembly)
            except RuntimeError:
                pass
            else:
                open_model_and_resume_all_groups(every_assembly)
                creopyson.file_close_window(creo_client, file_=every_assembly)
    logger.info(current_master_model)
    creopyson.file_open(creo_client, file_=current_master_model)


def resume_all_groups():

    """This function resumes all master models which contain GMXXXX. Later on this function started to use Mapkey instead of API command, because of resolve mode."""

    get_session_information()
    allow_conflicts()
    main_groups = []

    open_model_and_resume_all_groups(current_master_model)

    list_master_comp = creopyson.feature_list(creo_client, name='*' + gmXXXX + '*', file_=current_master_model, no_datum=True, type_='COMPONENT')
    for main_group in list_master_comp:
        main_groups.append(main_group["name"])

    for every_main_group_name in main_groups:
        open_model_and_resume_all_groups(every_main_group_name)

    creopyson.file_open(creo_client, file_=current_master_model)


def remove_master_groups_accoding_to_GUI():

    """ This function removes top master assemblies groups. Here we will remove level 2 assemblies
        by simple comparing their strings (or names) to picked dimensions (by user). Every top assembly
        goes through test, which determines whether dimension is useful - if assembly is useless is joined to list
        dimension_to_remove. Every property is specially tested. """

    dimensions_to_remove = []
    top_assemblies = []

    allow_conflicts()
    # Here we start by creating of list of the top assemblies
    top_assemblies_raw = creopyson.feature_list(creo_client, file_=source_assembly_name, no_datum=True, type_='COMPONENT')
    for every_top_assembly in top_assemblies_raw:
        top_assemblies.append(every_top_assembly['name'])

    # Here we are going to delete unnecessary clamp sizes
    if len(list_clamp_sizes) > 1:
        for every_elemenet in list_clamp_sizes:
            if every_elemenet != picked_clamp_size.get():
                dimensions_to_remove.append(every_elemenet)
        for every_top_assembly in top_assemblies:
            if any(x in every_top_assembly for x in dimensions_to_remove):
                try_delete_model(every_top_assembly)
        dimensions_to_remove = []

    # Special approach is given to powerpacks
    if len(list_powerpacks) > 1:
        # test whether is more than one powerpack in mastermodel
        for every_elemenet in list_powerpacks:
            # for each element in powerpacks (10 or 15) we are going to test whether string of powerpack *-10_, _10_ is in group
            # also here we are going to determine correct naming
            correct_powerpacks_testing_shape_1 = "-" + picked_power_pack.get() + "_"
            correct_powerpacks_testing_shape_2 = "_" + picked_power_pack.get() + "_"
            if every_elemenet != picked_power_pack.get():
                # if every element is not picked (for example "10")
                powerpacks_testing_shape_1 = "-" + every_elemenet + "_"
                powerpacks_testing_shape_2 = "_" + every_elemenet + "_"
                dimensions_to_remove.append(powerpacks_testing_shape_1)
                dimensions_to_remove.append(powerpacks_testing_shape_2)
        for every_top_assembly in top_assemblies:
            name_of_assembly = every_top_assembly.replace("ZE", "")
            name_of_assembly.replace("ze", "")
            global correct_group_exists;
            correct_group_exists = False
            if any(x in every_top_assembly for x in dimensions_to_remove):
                # before deletion of tested model we have to make sure that model of picked.type PP exists in mastermodel. If not, we will check whether
                # such a model exists in CS. If yes user will be informed about model and its up to him to add this model to mastermodel and also source model
                list_of_correct_powerpack_groups = []
                if powerpacks_testing_shape_1 in name_of_assembly and special_sign not in name_of_assembly:
                    correct_powerpack_master_model_11 = every_top_assembly.replace(powerpacks_testing_shape_1, correct_powerpacks_testing_shape_1)
                    correct_powerpack_master_model_12 = every_top_assembly.replace(powerpacks_testing_shape_1, correct_powerpacks_testing_shape_2)
                    list_of_correct_powerpack_groups = [correct_powerpack_master_model_11, correct_powerpack_master_model_12]
                if powerpacks_testing_shape_2 in name_of_assembly and special_sign not in name_of_assembly:
                    correct_powerpack_master_model_21 = every_top_assembly.replace(powerpacks_testing_shape_2, correct_powerpacks_testing_shape_1)
                    correct_powerpack_master_model_22 = every_top_assembly.replace(powerpacks_testing_shape_2, correct_powerpacks_testing_shape_2)
                    list_of_correct_powerpack_groups = [correct_powerpack_master_model_21, correct_powerpack_master_model_22]
                for every_correct_pp in list_of_correct_powerpack_groups:
                    if any(x == every_correct_pp for x in top_assemblies):
                        correct_group_exists = True
                    else:
                        if creopyson.file_exists(creo_client, file_=every_correct_pp):
                            assemble_model(every_correct_pp, current_master_model)
                            print("Powerpack model has been assembled to mastermodel. Name of group -" + every_correct_pp)
                            logger.info("Powerpack model has been assembled to mastermodel. Name of group -" + every_correct_pp)
                            correct_group_exists = True
                if correct_group_exists == False:
                    # Here should come the part where we rename model and we do not delete it!
                    open_model_and_resume_all_groups(every_top_assembly)
                    open_model_and_rename_groups(every_top_assembly, replace_from=powerpacks_testing_shape_1, replace_to=correct_powerpacks_testing_shape_1)
                    open_model_and_rename_groups(every_top_assembly, replace_from=powerpacks_testing_shape_2, replace_to=correct_powerpacks_testing_shape_1)
                else:
                    try_delete_model(every_top_assembly)
        dimensions_to_remove = []

        # We do this top assembly selection again, because some assemblies can be changed (renaming)
        top_assemblies_raw = creopyson.feature_list(creo_client, file_=source_assembly_name, no_datum=True, type_='COMPONENT')
        for every_top_assembly in top_assemblies_raw:
            top_assemblies.append(every_top_assembly['name'])

        # This special approach is given to also to second powerpacks
        if len(list_second_powerpacks) > 1:
            # test whether is more than one powerpack in mastermodel
            for every_elemenet in list_second_powerpacks:
                # for each element in powerpacks (01 or 02) we are going to test whether string of powerpack *-01_, _01_ is in group
                # also here we are going to determine correct naming
                correct_powerpacks_testing_shape_1 = "-" + picked_second_powerpack.get() + "_"
                correct_powerpacks_testing_shape_2 = "_" + picked_second_powerpack.get() + "_"
                if every_elemenet != picked_second_powerpack.get():
                    # if every element is not picked (for example "10")
                    powerpacks_testing_shape_1 = "-" + every_elemenet + "_"
                    powerpacks_testing_shape_2 = "_" + every_elemenet + "_"
                    dimensions_to_remove.append(powerpacks_testing_shape_1)
                    dimensions_to_remove.append(powerpacks_testing_shape_2)
            for every_top_assembly in top_assemblies:
                name_of_assembly = every_top_assembly.replace("ZE", "")
                name_of_assembly.replace("ze", "")
                correct_group_exists = False
                if any(x in name_of_assembly for x in dimensions_to_remove) and special_sign in name_of_assembly:
                    # before deletion of tested model we have to make sure that model of picked.type PP exists in mastermodel. If not, we will check whether
                    # such a model exists in CS. If yes user will be informed about model and its up to him to add this model to mastermodel and also source model
                    list_of_correct_powerpack_groups = []
                    if powerpacks_testing_shape_1 in name_of_assembly and special_sign in name_of_assembly:
                        correct_powerpack_master_model_11 = every_top_assembly.replace(powerpacks_testing_shape_1, correct_powerpacks_testing_shape_1)
                        correct_powerpack_master_model_12 = every_top_assembly.replace(powerpacks_testing_shape_1, correct_powerpacks_testing_shape_2)
                        list_of_correct_powerpack_groups = [correct_powerpack_master_model_11, correct_powerpack_master_model_12]
                    if powerpacks_testing_shape_2 in name_of_assembly and special_sign in name_of_assembly:
                        correct_powerpack_master_model_21 = every_top_assembly.replace(powerpacks_testing_shape_2, correct_powerpacks_testing_shape_1)
                        correct_powerpack_master_model_22 = every_top_assembly.replace(powerpacks_testing_shape_2, correct_powerpacks_testing_shape_2)
                        list_of_correct_powerpack_groups = [correct_powerpack_master_model_21, correct_powerpack_master_model_22]
                    for every_correct_pp in list_of_correct_powerpack_groups:
                        if any(x == every_correct_pp for x in top_assemblies):
                            correct_group_exists = True
                        else:
                            if creopyson.file_exists(creo_client, file_=every_correct_pp):
                                #creopyson.file_open(creo_client, file_=every_correct_pp)
                                #pdf_name = "_" + every_correct_pp.replace('.', '_')
                                #pdf_name = pdf_name.replace("-", "_")
                                assemble_model(every_correct_pp, current_master_model)
                                print("Powerpack model has been assembled to mastermodel. Name of group -"+ every_correct_pp)
                                logger.info("Powerpack model has been assembled to mastermodel. Name of group -" + every_correct_pp)
                                #creopyson.file_open(creo_client, file_=every_correct_pp)
                                #creopyson.interface_export_pdf(creo_client, filename=pdf_name, dirname=os.path.dirname(sys.argv[0]) + '\FeedbackFolder')
                                #print("Screenshot of missing powerpack group " + every_correct_pp + " has been created and stored to feedback folder.")
                                #creopyson.file_close_window(creo_client, file_=every_correct_pp)
                                correct_group_exists = True
                                #open_source_master_model()
                    if correct_group_exists == False:
                        # Here should come the part where we rename model and we do not delete it!
                        open_model_and_resume_all_groups(every_top_assembly)
                        open_model_and_rename_groups(every_top_assembly, replace_from=powerpacks_testing_shape_1, replace_to=correct_powerpacks_testing_shape_1)
                        open_model_and_rename_groups(every_top_assembly, replace_from=powerpacks_testing_shape_2, replace_to=correct_powerpacks_testing_shape_1)
                    else:
                        try_delete_model(every_top_assembly)
            dimensions_to_remove = []

    # We do this top assembly selection again, because some assemblies can be changed (renaming)
    top_assemblies_raw = creopyson.feature_list(creo_client, file_=source_assembly_name, no_datum=True, type_='COMPONENT')
    for every_top_assembly in top_assemblies_raw:
        top_assemblies.append(every_top_assembly['name'])

    # Next we are going to look on primary plastification. Special approach is also important because we would like to avoid possible
    # deleting of secondary plastification. Therefore we add next (special sign condition to this test)
    if len(list_primary_plast) > 1:
        for every_elemenet in list_primary_plast:
            if every_elemenet != picked_primary_plastification.get():
                dimensions_to_remove.append(every_elemenet)
        for every_top_assembly in top_assemblies:
            if any(x in every_top_assembly for x in dimensions_to_remove):
                if special_sign != None:
                    if special_sign not in every_top_assembly:
                        try_delete_model(every_top_assembly)
                else:
                    try_delete_model(every_top_assembly)
        dimensions_to_remove = []

    # Last models to remove are secondary plastification unnecessary models
    if len(list_secondary_plast) > 1:
        for every_elemenet in list_secondary_plast:
            if every_elemenet != picked_secondary_plastification.get():
                dimensions_to_remove.append(every_elemenet)
        for every_top_assembly in top_assemblies:
            if any(x in every_top_assembly for x in dimensions_to_remove):
                if special_sign != None:
                    if special_sign in every_top_assembly:
                        try_delete_model(every_top_assembly)


def load_zs_63():

    """This function loads zs63 SAP transaction - should be converted to SAP scripting technique"""

    global zs_63_folder
    zs_63_folder = os.path.dirname(sys.argv[0]) + '/ZS_63_source_folder'
    global feedback_folder
    feedback_folder = os.path.dirname(sys.argv[0]) + '/FeedbackFolder'

    remove_files_from_folder(zs_63_folder)
    remove_files_from_folder(feedback_folder)

    sap_source = tk.Tk()
    sap_source.filename = filedialog.askopenfilename(initialdir='\\', title='Choose ZS63 file', filetypes=(('text files', '*.txt'), ('all files', '*.*')))
    sap_source.destroy()

    current_folder_path = sap_source.filename
    final_folder_path = os.path.dirname(sys.argv[0]) + '\\ZS_63_source_folder\\ZS_63.txt'

    try:
        shutil.copyfile(current_folder_path, final_folder_path)
    except FileExistsError:
        os.replace(current_folder_path, final_folder_path)
    except FileNotFoundError:
        close_graphical_user_interface()

    # os.replace(current_folder_path, final_folder_path)


def remove_files_from_folder(folder_name):
    for f in os.walk(folder_name):
        for fileX in f:
            print(fileX)
        for each in fileX:
            os.remove(folder_name + '\\' + each)
            logger.info("removing file " + each)


def open_source_master_model():

    """This function opens source master model."""

    creopyson.file_open(creo_client, file_=source_assembly_name)


def preparation_master_model():

    """This function stands for preparation of the master model."""

    # 1 Yes purpose of this function is very obvious.
    open_source_master_model()
    # 2 This function removes all unnecessary assemmblies on the first level - maybe pushed to # 2
    remove_master_groups_accoding_to_GUI()
    # 3 This function resumes all supressed groups in mastermodel on two levels - maybe pushed to # 3
    resume_all_groups()
    # 4 This function created new copy of mastermodel. Function is based on breaking Kraussmaffei windchill using rules and its use is very questionable.
    create_new_copy()


def create_sa_groups():

    """ This function creates SA coordinates systems and empty SA assemblies.
        At first function gathers information about session according to which is then able to use
        master model or skeleton. Skeleton name is determine with proximity function therefore this
        has to work even with worse data quality of master models (especially MX machines). In the
        process function test whether assembly exists, in order to avoid possible renaming conflicts."""

    # Get session information - we need current order number
    creopyson.file_open(creo_client, file_=current_master_model)
    get_session_information()
    current_order_number = current_master_model[(len(current_master_model)) - 10:(len(current_master_model)) - 4]
    machine_type = current_master_model[0:3]
    machine_type = machine_type.replace('_', '')
    # This is expected name of skeleton model. We are going to look on all of the models and find the most simillar one.
    skeleton_name = current_master_model.replace('.asm', '_skel.prt')
    list_with_skeletons = []
    list_without_skeletons = []
    x = creopyson.bom.get_paths(creo_client, skeletons=True, top_level=True)
    x = x["children"]
    x = x["children"]
    for every_dict in x:
        list_with_skeletons.append(every_dict["file"])
    y = creopyson.bom.get_paths(creo_client, skeletons=False, top_level=True)
    y = y["children"]
    y = y["children"]
    for every_dict in y:
        list_without_skeletons.append(every_dict["file"])
    for every_file in list_without_skeletons:
        list_with_skeletons.remove(every_file)
    if len(list_with_skeletons) == 0:
        tk.messagebox.showerror('Automation status', 'Automation aborted! There is not skeleton model in mastermodel.')
        exit()
    elif len(list_with_skeletons) == 1:
        CAD_skel_name = list_with_skeletons[0]
    elif len(list_with_skeletons) > 1:
        skel_similarity = 0
        for each_possible_skel in list_with_skeletons:
            tested_similarity = SequenceMatcher.ratio(None, each_possible_skel, skeleton_name)
            if tested_similarity > skel_similarity:
                skel_similarity = tested_similarity
                CAD_skel_name = each_possible_skel

    print(CAD_skel_name)
    logger.info("CAD skeleton name is "+CAD_skel_name)

    # When we have correct skeleton we are going to look at global dictionary all_lists
    # Not defined CAD names are goind to be created. This module also covers possibilities: SA_group exists, CSY_exists

    for every_dict in all_lists:
        sap_group_name = every_dict['SAP_group_name']
        ERP_material_number = every_dict['ERP_number']
        cad_parent_model = every_dict['CAD_group_name']
        if cad_parent_model == 'Not defined' and 'SA' in sap_group_name:
            new_cad_parent = machine_type + '_' + sap_group_name + '_' + current_order_number + '.asm'
            new_csy_name = 'K_' + sap_group_name
            every_dict['CAD_group_name'] = new_cad_parent
            list_gm_groups(first_level_only=True)
            creopyson.file_close_window(creo_client, file_=current_master_model)
            if not any(x == new_cad_parent for x in list_components):
                creopyson.file_open(creo_client, CAD_skel_name)
                list_csys_model_raw = creopyson.feature_list(creo_client, file_=CAD_skel_name, type_='COORDINATE SYSTEM', no_comp=False)
                list_csys = []
                for every_dict in list_csys_model_raw:
                    list_csys.append(every_dict['name'])
                max_ratio = 0
                for every_csy in list_csys:
                    measure_similarity = SequenceMatcher(None, 'K_M2', every_csy)
                    every_csy_ratio = measure_similarity.ratio()
                    if every_csy_ratio > max_ratio:
                        max_ratio = every_csy_ratio
                        k_m2_csy = every_csy
                if not any(x == new_csy_name for x in list_csys):
                    create_csy_script = "~ Command `ProCmdDatumCsys` ;\
                                         ~ Open `storage_conflicts` `OptMenu1`;\
                                         ~ Close `storage_conflicts` `OptMenu1`;\
                                         ~ Select `storage_conflicts` `OptMenu1` 1 `resolution1`;\
                                         ~ Activate `storage_conflicts` `OK_PushButton`;\
                                         ~ Trigger `Odui_Dlg_00` `t1.OriginPlacement` 2 `0` `constr`;\
                                         ~ Focus `Odui_Dlg_00` `t1.OriginPlacement`;\
                                         ~ FocusIn `Odui_Dlg_00` `t1.OriginPlacement`;\
                                         ~ Select `Odui_Dlg_00` `t1.OriginPlacement` 2 `0` `constr`;\
                                         ~ Focus `Odui_Dlg_00` `t1.OriginPlacement`;\
                                         ~ RButtonArm `Odui_Dlg_00` `t1.OriginPlacement` 2 `0` `constr`;\
                                         ~ PopupOver `Odui_Dlg_00` `t1.OriginCollector_Wmo01` 1 `t1.OriginPlacement`;\
                                         ~ Open `Odui_Dlg_00` `t1.OriginCollector_Wmo01`;\
                                         ~ Trigger `Odui_Dlg_00` `t1.OriginPlacement` 2 `` ``;\
                                         ~ Timer `UI Desktop` `UI Desktop` `CollectorWdg_FocusTimer`;\
                                         ~ Close `Odui_Dlg_00` `t1.OriginCollector_Wmo01`;\
                                         ~ Activate `Odui_Dlg_00` `t1.DelOne_Wmo01`;\
                                         ~ Trigger `Odui_Dlg_00` `t1.OriginPlacement` 2 `` ``;\
                                         ~ FocusOut `Odui_Dlg_00` `t1.OriginPlacement`;\
                                         ~ Command `ProCmdMdlTreeSearch` ;\
                                         ~ Select `selspecdlg0` `RuleTab` 1 `Attributes`;\
                                         ~ Open `selspecdlg0` `SelOptionRadio`;\
                                         ~ Close `selspecdlg0` `SelOptionRadio`;\
                                         ~ Select `selspecdlg0` `SelOptionRadio` 1 `Coord Sys`;\
                                         ~ Update `selspecdlg0` `ExtRulesLayout.ExtBasicNameLayout.BasicNameList` \
                                         `K_M2`;~ Activate `selspecdlg0` `EvaluateBtn`;\
                                         ~ Activate `selspecdlg0` `ApplyBtn`;~ Activate `selspecdlg0` `CancelButton`;\
                                         ~ FocusIn `Odui_Dlg_00` `t1.OriginPlacement`;\
                                         ~ Trigger `Odui_Dlg_00` `t1.OriginPlacement` 2 `0` `constr`;\
                                         ~ Trigger `Odui_Dlg_00` `t1.OriginPlacement` 2 `` ``;\
                                         ~ FocusOut `Odui_Dlg_00` `t1.OriginPlacement`;\
                                         ~ Select `Odui_Dlg_00` `pg_vis_tab` 1 `tab_3`;\
                                         ~ FocusOut `Odui_Dlg_00` `t3.datum_csys_name`;\
                                         ~ Activate `Odui_Dlg_00` `Odui_Dlg_00`;\
                                         ~ Input `Odui_Dlg_00` `t3.datum_csys_name` ``;\
                                         ~ Update `Odui_Dlg_00` `t3.datum_csys_name` ``;\
                                         ~ FocusOut `Odui_Dlg_00` `t3.datum_csys_name`;\
                                         ~ Activate `Odui_Dlg_00` `Odui_Dlg_00`;\
                                         ~ Input `Odui_Dlg_00` `t3.datum_csys_name` `K_SAGROUP`;\
                                         ~ Update `Odui_Dlg_00` `t3.datum_csys_name` `K_SAGROUP`;\
                                         ~ Activate `Odui_Dlg_00` `t3.datum_csys_name`;\
                                         ~ FocusOut `Odui_Dlg_00` `t3.datum_csys_name`;\
                                         ~ Activate `Odui_Dlg_00` `stdbtn_1`;"
                    create_csy_script = create_csy_script.replace('K_M2', k_m2_csy)
                    create_csy_script = create_csy_script.replace('K_SAGROUP', new_csy_name)
                    creopyson.interface_mapkey(creo_client, create_csy_script)

                    continue_after_mapkey = False
                    list_csys_mapkey_test = []
                    while not continue_after_mapkey:
                        list_csys_model_raw = creopyson.feature_list(creo_client, file_=CAD_skel_name, type_='COORDINATE SYSTEM', no_comp=False)
                        for every_dict in list_csys_model_raw:
                            list_csys_mapkey_test.append(every_dict['name'])
                        time.sleep(0.2)
                        if any(x == new_csy_name for x in list_csys_mapkey_test):
                            continue_after_mapkey = True
                        list_csys_mapkey_test = []

                    print("Coordinate system for " + sap_group_name + " has been added to main skeleton.")
                    creopyson.file_close_window(creo_client, CAD_skel_name)
                    creopyson.file_open(creo_client, file_=current_master_model)
                if not creopyson.file_exists(creo_client, file_=new_cad_parent):
                    creopyson.file_open(creo_client, file_='MACHINETYPE_SAGROUP_ORDERNUMBER.ASM')
                    rename_config_control("yes")
                    creopyson.file_rename(creo_client, file_="MACHINETYPE_SAGROUP_ORDERNUMBER.ASM", new_name=new_cad_parent, onlysession=True)
                    print("Special assembly " + sap_group_name + " has been created.")
                    creopyson.feature_rename(creo_client, new_name=new_csy_name, name="K_SAGROUP", file_=new_cad_parent)
                    rename_config_control("no")
                    picked_csy = new_csy_name
                else:
                    pick_csy(current_order_number, new_cad_parent)
                    list_csys_model_raw = creopyson.feature_list(creo_client, file_=new_cad_parent, type_='COORDINATE SYSTEM', no_comp=False)
                    list_csys = []
                    for every_dict in list_csys_model_raw:
                        list_csys.append(every_dict['name'])
                    if len(list_csys) > 1:
                        picked_csy = list_csys[0]
                    else:
                        picked_csy = list_csys
                creopyson.file_open(creo_client, file_=current_master_model)
                list_gm_groups(first_level_only=True)
                if not any(x == new_cad_parent.upper() for x in list_components) and not any(x == new_cad_parent.lower() for x in list_components):
                    try:
                        creopyson.file_assemble(creo_client, file_=new_cad_parent, into_asm=current_master_model, ref_model=CAD_skel_name, constraints=[{"asmref": new_csy_name, "compref": new_csy_name, "type": "csys"}])
                    except RuntimeError:
                        pass
                    else:
                        print("Special assembly " + sap_group_name + " has been assembled to mastermodel.")
    creopyson.file_open(creo_client, file_=current_master_model)


def allow_conflicts():
    """This script is handy when it comes to giving program permissions to solve conflict."""

    old_script = "~ Command `ProCmdDatumCsys` ;~ Activate `storage_conflicts` `OK_PushButton`;~ Close `Odui_Dlg_00` `Odui_Dlg_00`;"
    script = "~ Activate `storage_conflicts` `OK_PushButton`;~ Close `Odui_Dlg_00` `Odui_Dlg_00`;"
    creopyson.interface_mapkey(creo_client, script)


def existing_nonassembled_models_feedback():

    """ This function is purposed to provide feedback to user about assembling process.
        It collects all non - assembled models and check their existence.
        If model exists, function creates pdf and saves this pdf to PartToConsider folder.
        This method of controlling non-assembled parts eases work of creators."""

    for every_dict in all_lists:
        if every_dict['CAD_group_name'] == 'Not defined':
            open_material_number(every_dict['ERP_number'])
            if model_name != '':
                jpeg_name_raw = model_name.replace('.', '_') + "_" + every_dict['SAP_group_name'].replace('.', '_')
                jpeg_name = ""
                for each_char in jpeg_name_raw:
                    if each_char.isalnum() or each_char == "_":
                        jpeg_name += each_char
                creopyson.file_open(creo_client, file_=model_name)
                image_location_dict = creopyson.interface_export_image(creo_client, file_type="JPEG", filename=jpeg_name)
                image_location = image_location_dict["dirname"] + image_location_dict["filename"]
                final_image_path = os.path.dirname(sys.argv[0]) + '/FeedbackFolder/' + image_location_dict["filename"]
                print(image_location)

                try:
                    shutil.move(image_location, final_image_path)
                #except FileExistsError:
                    #os.replace(image_location, final_image_path)
                except:
                    logger.warning("There was some problem to store screenshot of model "+ final_image_path)
                    logger.exception("message")

    feedback_button.config(command=open_feedback_folder)


def try_delete_model(model_to_delete):
    try:
        creopyson.feature_delete(creo_client, clip=True, name=model_to_delete)
    except RuntimeError:
        pass
    else:
        print('This group has been removed ' + model_to_delete)
        logger.info('This group has been removed ' + model_to_delete)


def open_model_and_rename_groups(component_name, replace_from, replace_to):
    """Name of this function is very clear. Model is opened and closed by itself. Creo session returns to previous model."""

    get_session_info = creopyson.file_get_fileinfo(creo_client)
    current_model = (get_session_info['file'])

    creopyson.file_open(creo_client, file_=component_name)
    allow_conflicts()
    list_function_components_raw = creopyson.feature_list(creo_client, name='*' + gmXXXX + '*', file_=component_name, no_datum=True, type_='COMPONENT')
    list_function_components = []
    list_function_components.append(component_name)

    for every_component_raw in list_function_components_raw:
        list_function_components.append(every_component_raw['name'])

    rename_config_control("yes")

    for every_component in list_function_components:
        if replace_from in every_component:
            new_name = every_component.replace(replace_from, replace_to)
            try:
                creopyson.file_rename(creo_client, file_=every_component, new_name=new_name, onlysession=True)
                print('this is new model = ' + new_name)
                logger.info('this is new model = ' + new_name)
            except RuntimeError:
                print('RuntimeError while renaming model' + new_name)
                logger.warning('RuntimeError while renaming model' + new_name)
            except NameError:
                print('NameError while renaming model' + new_name)
                logger.warning('NameError while renaming model' + new_name)

    script_ban_rename = "Select `main_dlg_cur` `appl_casc`;~ Close `main_dlg_cur` `appl_casc`;\~ Command `ProCmdRibbonOptionsDlg` ;\~ Select `ribbon_options_dialog` `PageSwitcherPageList` 1 `ConfigLayout`;\~ Activate `ribbon_options_dialog` `ConfigLayout.AddOpt`;\~ Input `add_opt` `InputOpt` `let_proe_rename_pdm_objects`;\~ Update `add_opt` `InputOpt` `let_proe_rename_pdm_objects`;\~ FocusOut `add_opt` `InputOpt`;~ Activate `add_opt` `AddOpt`;\~ Activate `ribbon_options_dialog` `OkPshBtn`;\~ FocusIn `UITools Msg Dialog Future` `no`;\~ Activate `UITools Msg Dialog Future` `no`;"
    creopyson.interface_mapkey(creo_client, script_ban_rename)

    creopyson.file_close_window(creo_client)
    creopyson.file_open(creo_client, file_=current_model)


def open_model_and_resume_all_groups(component_name, gmXXXX_only=True):

    get_session_info = creopyson.file_get_fileinfo(creo_client)
    current_model = (get_session_info['file'])

    if gmXXXX_only:
        key_gmXXXX=gmXXXX
    else:
        key_gmXXXX=""

    id_list = []
    resume_all_mapkey = "~ Command `ProCmdMdlTreeSearch`;\
                             ~ Activate `selspecdlg0` `SelScopeCheck` 0;\
                             ~ Input `selspecdlg0` `SelOptionRadio` `Component`;\
                             ~ Update `selspecdlg0` `SelOptionRadio` `Component`;\
                             ~ Select `selspecdlg0` `CascadeButton1`;\
                             ~ Close `selspecdlg0` `CascadeButton1`;\
                             ~ Activate `selspecdlg0` `Suppressed` 1;\
                             ~ Select `selspecdlg0` `RuleTab` 1 `Misc`;\
                             ~ Update `selspecdlg0` `ExtRulesLayout.ExtBasicIDLayout.InputIDPanel` `ID_COMPONENT`;\
                             ~ Activate `selspecdlg0` `EvaluateBtn`;\
                             ~ Activate `selspecdlg0` `EvaluateBtn`;\
                             ~ Select `selspecdlg0` `ResultList` -1;~ Activate `selspecdlg0` `ApplyBtn`;\
                             ~ Activate `selspecdlg0` `CancelButton`;\
                             ~ Close `selspecdlg0` `selspecdlg0`;\
                             ~ Command `ProCmdResume@PopupMenuTree`;\
                            ~ Activate `storage_conflicts` `OK_PushButton`;"

    creopyson.file_open(creo_client, file_=component_name)
    allow_conflicts()

    top_assemblies_raw = creopyson.feature_list(creo_client, file_=component_name, name='*' + key_gmXXXX + '*', no_datum=True, type_='COMPONENT')
    for every_main_group_name in top_assemblies_raw:
        if every_main_group_name["status"] == "SUPPRESSED":
            id_list.append(every_main_group_name["feat_id"])

    all_id_resumed = False

    while all_id_resumed == False:
        for each_id in id_list:
            resume_id = resume_all_mapkey.replace("ID_COMPONENT", str(each_id))
            creopyson.interface_mapkey(creo_client, resume_id)
        components_list = creopyson.feature_list(creo_client, file_=component_name, name='*' + key_gmXXXX + '*', no_datum=True, type_='COMPONENT')
        testing_suppressed_comp = []
        for each_comp in components_list:
            if each_comp["status"] == "SUPPRESSED":
                testing_suppressed_comp.append(every_main_group_name["feat_id"])

        if testing_suppressed_comp == []:
            print("all assies are resumed in this group name " + component_name + ".")
            logger.info("all assies are resumed in this group name " + component_name + ".")
            all_id_resumed = True
        else:
            time.sleep(0.1)

    creopyson.file_open(creo_client, file_=current_model)


def set_model_convention_on_the_fly():

    """This function aligns model naming, due to what users can expect higher percentage of succesfully added models."""
    list_gm_groups()
    get_session_information()
    list_of_numbers = [1, 2, 3, 4, 5, 6, 7, 8, 9, 0]
    for every_component in list_components:
        for every_number in list_of_numbers:
            if "_"+str(every_number)+"_"+gmXXXX in every_component:
                rename_config_control("yes")
                new_name = every_component.replace("_"+str(every_number)+"_"+gmXXXX, "-"+str(every_number)+"_"+gmXXXX)
                try:
                    creopyson.file_rename(creo_client, file_=every_component, new_name=new_name, onlysession=True)
                    try_remove_from_ws(every_component)
                    #creopyson.file_save(creo_client, file_=new_name)
                    logger.info('Repaired convention in model = ' + new_name)
                except:
                    pass
                rename_config_control("no")


def set_spritze_unit_for_mx():
    pass


def try_remove_from_ws(filename_to_remove):
    list_of_file_names = [filename_to_remove]
    try:
        creopyson.windchill_clear_workspace(creo_client, filenames=list_of_file_names)
    except:
        logger.warning("Program was unable to remove "+filename_to_remove + " from workspace.")


def set_default_view():
    try:
        list_views = creopyson.view_list(creo_client)
        if any(view_name_in_list == "DEFAULT" for view_name_in_list in list_views):
            creopyson.view_activate(creo_client, name="DEFAULT")
    except:
        logger.warning("Program failed to set Default view")


def rename_config_control(boolean):
    if boolean == "yes":
        creopyson.creo_set_config(client=creo_client, name="let_proe_rename_pdm_objects", value="yes")
    elif boolean == "no":
        creopyson.creo_set_config(client=creo_client, name="let_proe_rename_pdm_objects", value="no")


def compare_master_model():

    creoson_setup()
    get_session_information()
    read_ZS63_pair_with_CAD()
    bom_raw = transform_bom()
    bom_order_numbers = []
    # We will filter all the parts in opened assembly. Only parts (children) containing Order number substring will be saved to bom_order_number list
    order_number = get_session_order_number()
    for every_information in bom_raw:
        if order_number in every_information["parent"]:
            #bom_order_numbers.remove(every_information)
            bom_order_numbers.append(every_information)
    print(bom_order_numbers)
    #Comparation of CAD with SAP
    for every_dict in all_lists:
        if (every_dict['ERP_number']==x["child"][0:-4] for x in bom_order_numbers):
            print("Model is found in SAP and mastermodel "+every_dict['ERP_number'])
        else:
            jpeg_name_raw = model_name.replace('.', '_') + "_" + every_dict['SAP_group_name'].replace('.', '_')
            jpeg_name = ""
            for each_char in jpeg_name_raw:
                if each_char.isalnum() or each_char == "_":
                    jpeg_name += each_char
            creopyson.file_open(creo_client, file_=model_name)
            if model_name!="":
                image_location_dict = creopyson.interface_export_image(creo_client, file_type="JPEG", filename=jpeg_name)
                image_location = image_location_dict["dirname"] + image_location_dict["filename"]
                final_image_path = os.path.dirname(sys.argv[0]) + '/FeedbackFolder/' + image_location_dict["filename"]
                print(image_location)
                try:
                    shutil.move(image_location, final_image_path)
                # except FileExistsError:
                # os.replace(image_location, final_image_path)
                except:
                    logger.warning("There was some problem to store screenshot of model " + final_image_path)
                    logger.exception("message")
    creopyson.file_open(creo_client, file_=current_master_model)
    feedback_button.config(command=open_feedback_folder)



# Other useful functions which are not used in application


def bom_recursion(nest_dict, list_of_recursed_bom=[]):

    for key, value in nest_dict.items():
        if isinstance(value, dict):
            bom_recursion(value)
        elif isinstance(value, list):
            for each in value:
                if isinstance(each, dict):
                    bom_recursion(each)
        else:
            list_of_recursed_bom = list_of_recursed_bom
            if key != 'generic':
                list_of_recursed_bom.append(("{0} : {1}".format(key, value)))

    return list_of_recursed_bom


def only_numerics(seq):
    seq_type = type(seq)
    return seq_type().join(filter(seq_type.isdigit, seq))


def transform_bom():

    bom = []
    bom_cleansed = []
    bom_raw = bom_recursion(creopyson.bom_get_paths(creo_client))
    bom_raw.pop()
    bom_raw.reverse()

    for x in range(len(bom_raw)):
        if x % 2 == 0:
            pair_root_file_dict = {'root': bom_raw[x].replace('seq_path : ', ''), 'model': bom_raw[x+1].replace('file : ', '')}
            bom_cleansed.append(pair_root_file_dict)

    for each_pair in bom_cleansed:
        list_split_root = each_pair['root'].split('.')
        list_split_root.pop()
        joined_root = '.'.join(list_split_root)
        for each in bom_cleansed:
            if each['root'] == joined_root:
                pair_parent_child_dict = {'child': each_pair['model'], 'parent': each['model']}
                bom.append(pair_parent_child_dict)
                break
    return bom


def get_session_order_number():

    """This function slices order number from CAD master model for example "210952" is sliced from "cx_0250_210952.asm"""

    current_model = creopyson.file_get_fileinfo(creo_client)["file"]
    return current_model[-10:-4]


if __name__=="__main__":
    build_graphical_user_interface()
    #creoson_setup()
    #read_ZS63_pair_with_CAD(pair_also=True, test=True)



