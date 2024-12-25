import logging
import pandas as pd
import argparse
from pathlib import Path
from alive_progress import alive_bar, config_handler
from colorama import init, Fore
from colorama import Style
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
XLS_WINCC = _PRG_DIR / "tags_wincc.xlsx"
XLS_MOTOR_TEMPLATE = _TEMPLATE_DIR / "wincc_motor_template.xlsx"
XLS_INTERLOCK_TEMPLATE = _TEMPLATE_DIR / "wincc_interlock_template.xlsx"
POINT_TYPE = r"unimotor"
PLC = r"997"
FILTER = r".*"
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
    logger.info(f"Get children for:  {parent_tag_name}")
    mask = ECS_TAGS_DF['FunctionalHierarchy'].str.contains(parent_tag_name, case=False) & (
        ECS_TAGS_DF['Designation'].str.contains(r"int\d+|stp\d+|str\d+", case=False, regex=True)
    )
    logger.info(f"Found children :  {len(ECS_TAGS_DF[mask])}")
    return ECS_TAGS_DF[mask]


def make_motor_df(motor) -> DataFrame:
    """Создание DataFrame для тега мотора"""
    logger.info(f"Make data frame for motor:  {motor["Designation"]}")
    tag = motor["Designation"]  #"020PU044U01"
    plc = motor["IOType_0"][:3]  #"994"
    db_num = str(int(motor["IOType_6"]))  #"4314"
    parent_info = get_parent_info(motor["FunctionalHierarchy"])  # "994CD100G21 SPARE MOTOR"
    description = motor["DefaultText"]  # "SPARE MOTOR"

    motor_df = WINCC_MOTOR_TEMPLATE_DF.copy()
    motor_df = motor_df.replace([r'\$tag_name\$', r'\$plc_num\$', r'\$dbnum\$', r'\$parent_info\$', r'\$description\$'],
                                [tag, plc, db_num, parent_info, description], regex=True)

    children_df = get_children_interlock(motor["Designation"])

    for _, row in children_df.iterrows():
        interlock_df = WINCC_INTERLOCK_TEMPLATE_DF.copy()
        # extract from string "Path" part after last colon
        interlock_name = row["Path"].split(":")[-1]

        interlock_df = interlock_df.replace(
            [r'\$tag_name\$', r'\$interlock\$', r'\$plc_num\$', r'\$addr\$', r'\$description\$'],
            [tag, interlock_name, plc, f"{str(row["IOType_3"])[:-2]}", row["DefaultText"]], regex=True)

        motor_df = pd.concat([motor_df, interlock_df], axis=0)

    return motor_df


def ecs2wincc(wincc_file: str, point_type: str, plc: str, filter: str) -> DataFrame:
    """Конвертация экспорта тегов  из XLSX ECS в формат импорта XLSX WinCC"""
    global ECS_TAGS_DF
    logger.info(f"Конвертация ECS -> WinCC:  -> {wincc_file}, Тип оборудования: {point_type}, PLC: {plc}, "
                f"Фильтр: {filter}")
    print(f"{Fore.YELLOW}Converting ECS tags to WINCC: {Fore.MAGENTA}{Style.RESET_ALL}", end="\t")
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
    print(f"{Fore.GREEN}OK {Style.RESET_ALL}")
    print(f"{Fore.YELLOW}Qty tags: {Fore.MAGENTA}{len(ecs_tags_flt_df)}{Style.RESET_ALL}", )
    return wincc_df


def open_templates() -> []:
    """Open xlsx templates"""
    print(f"{Fore.YELLOW}Opening templates: {Fore.MAGENTA}{_TEMPLATE_DIR}{Style.RESET_ALL}", end="\t")
    logger.info(f"Opening templates: {_TEMPLATE_DIR}")
    try:
        wincc_motor_template_df = pd.read_excel(XLS_MOTOR_TEMPLATE)
        wincc_interlock_template_df = pd.read_excel(XLS_INTERLOCK_TEMPLATE).astype(str)
    except Exception as e:
        logger.error(f"Error reading templates: {e}")
        raise
    print(f"{Fore.GREEN}OK {Style.RESET_ALL}")
    return [wincc_motor_template_df, wincc_interlock_template_df]



def open_ecs_tags_xlsx(ecs_file: str) -> DataFrame | None:
    """Open xlsx file with ECS tags"""

    print(f"{Fore.YELLOW}Opening ECS tags: {Fore.MAGENTA}{ecs_file}{Style.RESET_ALL}", end="\t")
    logger.info(f"Opening ECS tags: {ecs_file}")
    try:
        ecs_tags_df = pd.read_excel(ecs_file)
        logger.info(f"Read ECS tags: {ecs_file}")
    except Exception as e:
        logger.error(f"Error read ECS tags: {e}")
        raise Exception(f"Error read ECS tags: {e}")

    print(f"{Fore.GREEN}OK {Style.RESET_ALL}")
    return ecs_tags_df

def write_wincc_xlsx(wincc_df: DataFrame, wincc_file: str):
    """Write wincc_df to xlsx"""
    print(f"{Fore.YELLOW}Saving to: {Fore.MAGENTA}{str(wincc_file)} {Style.RESET_ALL}", end="\t")
    try:
        with pd.ExcelWriter(wincc_file, engine="xlsxwriter") as writer:
            wincc_df.to_excel(writer, index=False, header=True, sheet_name="Hmi Tags")

        print(f"{Fore.GREEN}OK {Style.RESET_ALL}")

    except Exception as e:
        print(f"{Fore.RED}FAULT{Style.RESET_ALL}")
        err_msg = f"{Fore.RED}Error: Can't write tags to: {str(wincc_file)}"
        print(err_msg)
        raise Exception(err_msg)


def main():
    global WINCC_MOTOR_TEMPLATE_DF, WINCC_INTERLOCK_TEMPLATE_DF, ECS_TAGS_DF
    print(
        f"\n{Fore.LIGHTWHITE_EX}ECS2WINCC v{_VERSION} - utility converts data from the xlsx tag export MOUSTACHE to xlsx for import into TIA WINCC Prof. {Style.RESET_ALL}")
    parser = argparse.ArgumentParser(
        prog="ecs2wincc",
        description="Конвертация экспорта тегов  из XLSX ECS в формат импорта XLSX WinCC Prof",
        epilog=f'2024 7Art v{_VERSION}'
    )
    parser.add_argument(
        "-v", "--version", action="version", version=f"Version {_VERSION}"
    )
    parser.add_argument(
         "file", type=str, default=XLS_ECS, help=f"Имя файла с тегами ECS (По умолчанию {XLS_ECS})"
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
    logger.info(f"Run program with arg:  {parser.parse_args()}")

    print(f"{Fore.YELLOW}Run program with arg: {Fore.MAGENTA}{parser.parse_args()}{Style.RESET_ALL}")
    with alive_bar(4, force_tty=True, length=12, spinner='classic') as bar:

        WINCC_MOTOR_TEMPLATE_DF, WINCC_INTERLOCK_TEMPLATE_DF = open_templates()
        bar()
        ECS_TAGS_DF = open_ecs_tags_xlsx(ecs_file)
        bar()
        wincc_df = ecs2wincc(wincc_file, point_type, plc, tag_filter)
        bar()
        write_wincc_xlsx(wincc_df, wincc_file)
        bar()

if __name__ == "__main__":
    main()
