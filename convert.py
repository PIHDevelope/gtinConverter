import os
import pathlib
import re
import xml.etree.ElementTree as ET
from ntpath import join
from typing import List

import colorama
import xlsxwriter
from colorama import Back, Fore, Style
from genericpath import isfile
from pynput import keyboard

INPUT_DIRECTORY: str = "Input"
INPUT_FILE_EXT: str = "xml"

OUTPUT_DIRECTORY: str = "Output"
OUTPUT_FILE_EXT: str = "xlsx"

DICTIONARY_DIRECTORY: str = "Dictionary"
DICTIONARYFILE_NAME: str = "dict"
DICTIONARY_FILE_EXT: str = "txt"

SHOW_INSTRUCTION = False

DOT = "."
SLASH = "/"


def create_directory(directory_name: str) -> None:
    try:
        os.mkdir(directory_name)
    except OSError as error:
        pass


def check_for_directory_is_exists_and_create_if_not(directory_name: str) -> bool:
    directory_is_exists = os.path.exists(directory_name)
    if not directory_is_exists:
        create_directory(directory_name)
    return directory_is_exists


def filter_file_list_by_file_extension(file_list: List[str], extension: str) -> List[str]:
    extension = DOT + extension
    return list(filter(lambda file_name: pathlib.Path(
        file_name).suffix == extension,  file_list))


def get_file_list(directory_name: str) -> List[str]:
    return [f for f in os.listdir(
        directory_name) if isfile(join(directory_name, f))]


def get_file_name(file: str) -> str:
    return pathlib.Path(file).stem


def on_press(key):
    pass


def on_release(key):
    return False


if __name__ == "__main__":

    colorama.init()

    print(f"{Back.BLUE}█▀█ ▄▀█ █▀▀ █ █▀▀ █ █▀▀   █ █▄░█ ▀█▀ █▀▀ █▀█ █▄░█ ▄▀█ ▀█▀ █ █▀█ █▄░█ ▄▀█ █░░   █░█ █▀█ █▀ █▀█ █ ▀█▀ ▄▀█ █░░{Style.RESET_ALL}")
    print(f"{Back.BLUE}█▀▀ █▀█ █▄▄ █ █▀░ █ █▄▄   █ █░▀█ ░█░ ██▄ █▀▄ █░▀█ █▀█ ░█░ █ █▄█ █░▀█ █▀█ █▄▄   █▀█ █▄█ ▄█ █▀▀ █ ░█░ █▀█ █▄▄{Style.RESET_ALL}\n")

    #
    input_exits = check_for_directory_is_exists_and_create_if_not(
        INPUT_DIRECTORY)
    output_exists = check_for_directory_is_exists_and_create_if_not(
        OUTPUT_DIRECTORY)
    dictinary_exists = check_for_directory_is_exists_and_create_if_not(
        DICTIONARY_DIRECTORY)
    #
    input_files_list: List[str] = get_file_list(INPUT_DIRECTORY)
    #
    instruction_list: List[str] = []
    if not input_exits or SHOW_INSTRUCTION:
        instruction_list.append(
            f"Положите файлы  для преобразования в папку {INPUT_DIRECTORY}.")
    if not output_exists or SHOW_INSTRUCTION:
        instruction_list.append(
            f"Заберите преобразованные файлы в папке {INPUT_DIRECTORY}.")
    if not dictinary_exists or SHOW_INSTRUCTION:
        instruction_list.append(
            f"Положите файл в папку {DICTIONARY_DIRECTORY}.")
    if len(instruction_list) > 0:
        print("\n".join(instruction_list))
    #
    input_file_list = filter_file_list_by_file_extension(get_file_list(
        INPUT_DIRECTORY), INPUT_FILE_EXT)
    dictionary_file_list = filter_file_list_by_file_extension(get_file_list(
        DICTIONARY_DIRECTORY), DICTIONARY_FILE_EXT)
    #
    gtin_dictionary: dict[str, str] = {}
    #
    if len(dictionary_file_list) == 0:
        print(f"{Back.RED}Словарь пуст! Убедитесь, что в папке {DICTIONARY_DIRECTORY} распологается словарь. Словарь - это файл с расширением {DICTIONARY_FILE_EXT}.{Style.RESET_ALL}")
    else:
        #
        search_pattern = r"(.*)(\s+)(\d{13}$)"
        for dictionary_file in dictionary_file_list:
            with open(SLASH.join([DICTIONARY_DIRECTORY, dictionary_file])) as gtin_dictionary_file:
                for line in gtin_dictionary_file:
                    search_result = re.findall(search_pattern, line)
                    if len(search_result) > 0:
                        gtin_dictionary[search_result[0]
                                        [2]] = search_result[0][0]
                    else:
                        pass
        #
        search_pattern = r"(\))(\d+)(\()"
        if len(input_file_list) == 0:
            print(f"{Back.RED}Отсутствуют входные файлы для преобразования! Убедитесь, что в папке {INPUT_DIRECTORY} распологаются файлы. Это файлы с расширением {INPUT_FILE_EXT}.{Style.RESET_ALL}.{Style.RESET_ALL}")
        else:
            print(
                f"{Back.BLUE}Все преобразованные файлы находятся в папке {OUTPUT_DIRECTORY}. После выполнения преобразования входные файлы удаляются!{Style.RESET_ALL}")
            for input_file in input_file_list:
                output_file_name = get_file_name(input_file)
                output_file = DOT.join([output_file_name, OUTPUT_FILE_EXT])
                output_file_path = SLASH.join([OUTPUT_DIRECTORY, output_file])
                input_file_path = SLASH.join(
                    [INPUT_DIRECTORY, input_file])
                input_file_root_node = ET.parse(input_file_path).getroot()
                row_number = 0
                with xlsxwriter.Workbook(output_file_path) as workbook:
                    worksheet = workbook.add_worksheet()
                    for child in input_file_root_node:
                        attribute = child.attrib["BarCode"]
                        search_result = re.findall(search_pattern, attribute)
                        gtin_value = None
                        if len(search_result) > 0:
                            gtin_value = search_result[0][1]
                            # 04xxx -> 4xxx
                            if gtin_value[0] == "0":
                                gtin_value = gtin_value[1:]
                        if gtin_value is not None:
                            if gtin_value in gtin_dictionary:
                                worksheet.write(
                                    row_number, 0, gtin_dictionary[gtin_value])
                                worksheet.write(row_number, 1, gtin_value)
                                worksheet.write(row_number, 2, 1)
                                row_number += 1
                        else:
                            print(
                                f"{Back.RED}Значение {gtin_value} отсутствует в словаре!.{Style.RESET_ALL}")
                print(
                    f"{Back.GREEN}Входной файл \"{input_file}\" преобразован в \"{output_file}\".{Style.RESET_ALL}")
                os.remove(input_file_path)
print("Нажмите любую клавишу для закрытия...")
with keyboard.Listener(
        on_press=on_press,
        on_release=on_release) as listener:
    listener.join()
#
