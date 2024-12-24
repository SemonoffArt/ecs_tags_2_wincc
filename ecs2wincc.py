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
_RESOURCES_DIR = _PRG_DIR / "resources"
_TEMPLATE_DIR = _RESOURCES_DIR / "templates"
_LOG_FILE = _PRG_DIR / "ecs2wincc.log"
_MAGADAN_UTC = 11  # Магаданское время +11 часов к UTC
DEBUG = True
QTY_TAGS = 0  # количество тегов
QTY_VALS = 0  # количество значений
XLS_ECS = _PRG_DIR / "Points.xlsx"
XLS_WINCC = _PRG_DIR / "tags_wincc_test.xlsx"
XLS_MOTOR_TEMPLATE = _TEMPLATE_DIR / "wincc_motor_template.xlsx"
XLS_INTERLOCK_TEMPLATE = _TEMPLATE_DIR / "wincc_interlock_template.xlsx"
POINT_TYPE = r"unimotor"
PLC = r"994"
FILTER = r"sj"
WINCC_MOTOR_TEMPLATE_DF = pd.DataFrame()
WINCC_INTERLOCK_TEMPLATE_DF = pd.DataFrame()
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


def get_children_interlock(parent_tag_name: str) -> DataFrame:
    """Получение информации о дочерних тегах, только интерлоки"""
    mask = ECS_TAGS_DF['FunctionalHierarchy'].str.contains(parent_tag_name, case=False) & (
        ECS_TAGS_DF['Designation'].str.contains(r"int\d+|stp\d+|str\d+", case=False, regex=True)
    )
    return ECS_TAGS_DF[mask]


def make_interlock_df(interlocks:DataFrame) -> DataFrame:
    """Создание DataFrame для тега интерлока"""

    tag = interlocks["Designation"]  #"020PU044U01"
    plc = interlocks["IOType_0"][:3]  #"994"
    db_num: str = str(int(motor["IOType_6"]))  #"4314"
    parent_info = get_parent_info(motor["FunctionalHierarchy"])  # "994CD100G21 SPARE MOTOR"


    return

def make_motor_df(motor) -> DataFrame:
    """Создание DataFrame для тега мотора"""
    tag = motor["Designation"]  #"020PU044U01"
    plc = motor["IOType_0"][:3]  #"994"
    # db_num: str = str(int(motor["IOType_6"]))  #"4314"
    db_num = str(int(motor["IOType_6"]))  #"4314"
    parent_info = get_parent_info(motor["FunctionalHierarchy"])  # "994CD100G21 SPARE MOTOR"
    description = motor["DefaultText"]  # "SPARE MOTOR"

    motor_df = WINCC_MOTOR_TEMPLATE_DF.copy()
    # replace part string value in column "Name" to variable value "tag"
    # motor_df["Name"] = motor_df["Name"].str.replace("$tag_name$", tag)
    # motor_df["Path"] = motor_df["Path"].str.replace("$tag_name$", tag)
    # motor_df["Connection"] = motor_df["Connection"].str.replace("$plc_num$", plc)
    # motor_df["Address"] = motor_df["Address"].str.replace("$dbnum$", db_num)
    # # find row with "ParentInfo" in column "Name" and insert value in "Tag value [en-US]"
    # motor_df.loc[motor_df["Name"].str.contains("ParentInfo", case=False), "Tag value [en-US]"] = parent_info
    # motor_df.loc[motor_df["Name"].str.contains("DefaultText", case=False), "Tag value [en-US]"] = description
    motor_df = motor_df.replace([r'\$tag_name\$', r'\$plc_num\$', r'\$dbnum\$', r'\$parent_info\$', r'\$description\$'],
                                [tag, plc, db_num, parent_info, description], regex=True)

    # print(motor_df)
    # get childs tags
    children_df = get_children_interlock(motor["Designation"])

    for _, row in children_df.iterrows():
        interlock_df = WINCC_INTERLOCK_TEMPLATE_DF.copy()
        # extract from string "Path" part after last colon
        interlock_name = row["Path"].split(":")[-1]
        # interlock_df = interlock_df.replace([r'\$tag_name\$', r'\$interlock\$', r'\$plc_num\$', r'\$addr\$', r'\$description\$'],
        #                                     [tag, interlock_name, plc, f"{int(row["IOType_3"])}", row["DefaultText"]], regex=True)
        interlock_df = interlock_df.replace(
            [r'\$tag_name\$', r'\$interlock\$', r'\$plc_num\$', r'\$addr\$', r'\$description\$'],
            [tag, interlock_name, plc, f"{str(row["IOType_3"])[:-2]}", row["DefaultText"]], regex=True)

        motor_df = pd.concat([motor_df, interlock_df], axis=0)


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
    for _, row in ecs_tags_flt_df.iterrows():
        # Process each row
        wincc_df = pd.concat([wincc_df, make_motor_df(row)], axis=0)

    # print(wincc_df)
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
    global WINCC_MOTOR_TEMPLATE_DF, WINCC_INTERLOCK_TEMPLATE_DF,  ECS_TAGS_DF
    parser = argparse.ArgumentParser(
        prog="ecs2wincc",
        description="Конвертация экспорта тегов  из XLSX ECS в формат импорта XLSX WinCC",
        epilog=f'2024 7Art v{_VERSION}'
    )
    parser.add_argument(
        "-v", "--version", action="version", version=f"Version {_VERSION}"
    )
    parser.add_argument(
        "-f", "--file", type=str, default=XLS_ECS, help=f"Имя файла с тегами ECS (По умолчанию {XLS_ECS})"
    )
    parser.add_argument(
        "-o", "--output", type=str, default=XLS_WINCC,
        help=f"Имя файла для сохранения данных WinCC (По умолчанию: {XLS_WINCC})"
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
        WINCC_MOTOR_TEMPLATE_DF = pd.read_excel(XLS_MOTOR_TEMPLATE)
    except Exception as e:
        logger.error(f"Ошибка чтения шаблона wincc motor: {e}")
        return

    try:
        WINCC_INTERLOCK_TEMPLATE_DF = pd.read_excel(XLS_INTERLOCK_TEMPLATE)
        WINCC_INTERLOCK_TEMPLATE_DF = WINCC_INTERLOCK_TEMPLATE_DF.astype(str)

    except Exception as e:
        logger.error(f"Ошибка чтения шаблона wincc interlock: {e}")
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
