import logging
import os
import time
from logging import DEBUG

import pandas as pd
from re import search
import argparse
from pathlib import Path
from alive_progress import alive_bar, config_handler
from colorama import init, Fore
from colorama import Style
from pandas import DataFrame
from pandas.core.interchange.dataframe_protocol import DataFrame

_VERSION = 0.1
_PRG_DIR = Path("./").absolute()
_LOG_FILE = _PRG_DIR / "ecs2wincc.log"
_MAGADAN_UTC = 11  # Магаданское время +11 часов к UTC
DEBUG = True
QTY_TAGS = 0  # количество тегов
QTY_VALS = 0  # количество значений
XLS_FILE_ECS = _PRG_DIR / "Points.xlsx"
XLS_FILE_WINCC = _PRG_DIR / "tags_wincc_test.xlsx"
XLS_FILE_WINCC_TEMPLATE = _PRG_DIR / "wincc_template.xlsx"
POINT_TYPE = r"unimotor"
PLC = r"994"
FILTER = r"pu044"
WINCC_TEMPLATE_DF = pd.DataFrame()
ECS_TAGS_DF = pd.DataFrame()

init(autoreset=True)
log_format = f"%(asctime)s - %(levelname)s -(%(funcName)s(%(lineno)d) - %(message)s"
logging.basicConfig(
    format=log_format,
    level=logging.INFO,
    filename=_LOG_FILE,
    encoding="UTF-8"
)
logger = logging.getLogger()


def get_parent_info(parent_tag_name: str) -> str:
    """Получение  информации о родителе"""
    global ECS_TAGS_DF
    parent_info = f"{parent_tag_name} {ECS_TAGS_DF.loc[ECS_TAGS_DF['Designation'].str.contains(parent_tag_name, case=False), 'DefaultText'].values[0]}"

    return parent_info


def get_childs(parent_tag_name: str) -> DataFrame:
    """Получение информации о дочерних тегах"""
    mask = ECS_TAGS_DF['FunctionalHierarchy'].str.contains(parent_tag_name, case=False) & (
        ECS_TAGS_DF['Designation'].str.contains("int", case=False) |
        ECS_TAGS_DF['Designation'].str.contains("stp", case=False) |
        ECS_TAGS_DF['Designation'].str.contains("str", case=False)
    )
    return ECS_TAGS_DF[mask]


def make_interloc_df(interloc) -> DataFrame:
    """Создание DataFrame для тега интерлока"""
    head = [["Name", "Path", "Connection", "PLC tag", "DataType", "HMI DataType", "Length", "Coding", "Access Method", "Address", "Start value", "Quality Code", "Persistency", "Substitute value", "Tag value [en-US]", "Update Mode", "Comment [en-US]", "Limit Upper 2 Type", "Limit Upper 2", "Limit Lower 2 Type", "Limit Lower 2", "Linear scaling", "End value PLC", "Start value PLC", "End value HMI", "Start value HMI", "Synchronization"]]

    tag = interloc["Designation"]  #"020PU044U01"
    plc = interloc["IOType_0"][:3]  #"994"
    db_num: str = str(int(motor["IOType_6"]))  #"4314"
    parent_info = get_parent_info(motor["FunctionalHierarchy"])  # "994CD100G21 SPARE MOTOR"


    return

def make_motor_df(motor) -> DataFrame:
    """Создание DataFrame для тега мотора"""
    tag = motor["Designation"]  #"020PU044U01"
    plc = motor["IOType_0"][:3]  #"994"
    db_num: str = str(int(motor["IOType_6"]))  #"4314"
    parent_info = get_parent_info(motor["FunctionalHierarchy"])  # "994CD100G21 SPARE MOTOR"

    motor_df = WINCC_TEMPLATE_DF.copy()
    # replace part string value in column "Name" to variable value "tag"
    motor_df["Name"] = motor_df["Name"].str.replace("$tag_name$", tag)
    motor_df["Path"] = motor_df["Path"].str.replace("$tag_name$", tag)
    motor_df["Connection"] = motor_df["Connection"].str.replace("$plc_num$", plc)
    motor_df["Address"] = motor_df["Address"].str.replace("$dbnum$", db_num)
    # find row with "ParentInfo" in column "Name" and insert value in "Tag value [en-US]"
    motor_df.loc[motor_df["Name"].str.contains("ParentInfo", case=False), "Tag value [en-US]"] = parent_info
    # print(motor_df)
    # get childs tags
    childs_df = get_childs(motor["Designation"])

    return motor_df


def ecs2wincc(wincc_file: str, point_type: str, plc: str, filter: str) -> None:
    """Конвертация экспорта тегов  из XLSX ECS в формат импорта XLSX WinCC"""
    global ECS_TAGS_DF
    logger.info(f"Конвертация ECS -> WinCC:  -> {wincc_file}, Тип оборудования: {point_type}, PLC: {plc}, "
                f"Фильтр: {filter}")
    wincc_df = pd.DataFrame()

    # Фильтр по тегам
    ecs_tags_flt_df = ECS_TAGS_DF[ECS_TAGS_DF['PointType'].str.contains(point_type, case=False) &
                                  ECS_TAGS_DF['IOType_0'].str.contains(plc, case=False, na=False) &
                                  ECS_TAGS_DF['Designation'].str.contains(filter, case=False, na=False)]

    # iterating through the ecs rows in a loop
    for index, row in ecs_tags_flt_df.iterrows():
        # Process each row
        wincc_df = pd.concat([wincc_df, make_motor_df(row)], axis=0)

    print(wincc_df)
    # wincc_df["Quality Code"] = wincc_df["Quality Code"].astype(str)
    # wincc_df["Persistency"] = wincc_df["Persistency"].astype(str)
    # wincc_df["Linear scaling"] = wincc_df["Linear scaling"].astype(str)
    # wincc_df["Synchronization"] = wincc_df["Synchronization"].astype(str)
    wincc_df = wincc_df.astype(str)
    wincc_df["Limit Upper 2 Type"] = "None"
    wincc_df["Limit Lower 2 Type"] = "None"
    # write wincc_df to xlsx
    wincc_df.to_excel(wincc_file, index=False, header=True, sheet_name="Hmi Tags")


def main():
    global WINCC_TEMPLATE_DF, ECS_TAGS_DF
    parser = argparse.ArgumentParser(
        prog="ecs2wincc",
        description="Конвертация экспорта тегов  из XLSX ECS в формат импорта XLSX WinCC",
        epilog=f'2024 7Art v{_VERSION}'
    )
    parser.add_argument(
        "-v", "--version", action="version", version=f"Version {_VERSION}"
    )
    parser.add_argument(
        "-f", "--file", type=str, default=XLS_FILE_ECS, help=f"Имя файла с тегами ECS (По умолчанию {XLS_FILE_ECS})"
    )
    parser.add_argument(
        "-o", "--output", type=str, default=XLS_FILE_WINCC,
        help=f"Имя файла для сохранения данных WinCC (По умолчанию: {XLS_FILE_WINCC})"
    )
    parser.add_argument(
        "-t", "--point_type", type=str, default=POINT_TYPE, help="Тип оборудования (по умолчанию: unimotor)"

    )
    parser.add_argument(
        "-p", "--plc", type=str, default=PLC, help="PLC (по умолчанию: 994)"

    )
    parser.add_argument(
        "-ftr", "--filter", type=str, default=FILTER, help="Фильтр для тегов (по умолчанию: .*)"

    )

    args = parser.parse_args()
    ecs_file = args.file
    wincc_file = args.output
    point_type = args.point_type
    plc = args.plc
    tag_filter = args.filter

    try:
        WINCC_TEMPLATE_DF = pd.read_excel(XLS_FILE_WINCC_TEMPLATE)
    except Exception as e:
        logger.error(f"Ошибка чтения шаблона wincc: {e}")
        return

    try:
        ECS_TAGS_DF = pd.read_excel(ecs_file)
    except Exception as e:
        logger.error(f"Ошибка чтения файла ECS: {e}")
        return

    if not wincc_file:
        wincc_file = ecs_file.replace(".xlsx", "_wincc.xlsx")
    logger.info(f"Файл ECS: {ecs_file}, файл WinCC: {wincc_file}")

    ecs2wincc(wincc_file, point_type, plc, tag_filter)


if __name__ == "__main__":
    main()
