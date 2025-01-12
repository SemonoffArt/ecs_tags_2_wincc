import logging
import pandas as pd
import argparse
from pathlib import Path
from alive_progress import alive_bar
from colorama import init, Fore, Style

class ECS2WinCCConverter:
    def __init__(self, version=0.1):
        self._version = version
        self._prg_dir = Path("./").absolute()
        self._resources_dir = self._prg_dir / "resources"
        self._template_dir = self._resources_dir / "templates"
        self._log_file = self._prg_dir / "ecs2wincc.log"
        self.templates_df = {}
        self.ecs_tags_df = pd.DataFrame()
        self._init_logging()
        self._init_colorama()

    def _init_logging(self):
        logging.basicConfig(
            format="%(asctime)s - %(levelname)s -(%(funcName)s(%(lineno)d) - %(message)s",
            level=logging.INFO,
            filename=self._log_file,
            encoding="UTF-8"
        )
        self.logger = logging.getLogger()

    def _init_colorama(self):
        init(autoreset=True)

    def get_parent_info(self, parent_tag_name: str) -> str:
        parent_info = f"{parent_tag_name} {self.ecs_tags_df.loc[self.ecs_tags_df['Designation'].str.contains(parent_tag_name, case=False), 'DefaultText'].values[0]}"
        return parent_info

    def get_children_interlock(self, parent_tag_name: str) -> pd.DataFrame:
        self.logger.info(f"Get children for: {parent_tag_name}")
        mask = self.ecs_tags_df['FunctionalHierarchy'].str.contains(parent_tag_name, case=False) & (
            self.ecs_tags_df['Designation'].str.contains(r"int\d+|stp\d+|str\d+", case=False, regex=True)
        )
        self.logger.info(f"Found children: {len(self.ecs_tags_df[mask])}")
        return self.ecs_tags_df[mask]

    def extract_point_type(self, point_type: str) -> str:
        if "motor" in point_type.lower():
            return "motor"
        elif "valve" in point_type.lower():
            return "valve"
        elif "analog" in point_type.lower():
            return "analog"
        else:
            return "unknown"

    def get_decimal_format(self, decimals: str) -> str:
        try:
            decimals = int(float(decimals))
            if decimals == 0:
                return "s999999"
            elif decimals == 1:
                return "s999999.9"
            elif decimals == 2:
                return "s999999.99"
            elif decimals == 3:
                return "s999999.999"
            else:
                return "s9999999"
        except Exception as e:
            self.logger.error(f"Error get decimal format: {e}")
            return ""

    def make_unit_df(self, unit) -> pd.DataFrame:
        self.logger.info(f"Make data frame for unit: {unit['Designation']}")
        tag = unit["Designation"]
        point_type = self.extract_point_type(unit["PointType"])
        plc = unit["IOType_0"][:3]
        db_num = unit["IOType_6"].replace(".0", "")
        parent_info = self.get_parent_info(unit["FunctionalHierarchy"])
        description = unit["DefaultText"]
        eu = f" {unit['Unit'].replace('Acesys.Unit.', '').replace('PointType.Unit.', '')}" if len(unit["Unit"]) > 0 else ""
        decimals = self.get_decimal_format(unit["Decimals"])
        trend_tag_name = f"{tag}.MSW.VALUE"

        unit_df = self.templates_df[point_type].copy()
        unit_df = unit_df.replace(
            [r'\$tag_name\$', r'\$plc_num\$', r'\$dbnum\$', r'\$parent_info\$', r'\$description\$', r'\$eu\$', r'\$decimals\$', r'\$trend_tag_name\$'],
            [tag, plc, db_num, parent_info, description, eu, decimals, trend_tag_name], regex=True
        )

        children_df = self.get_children_interlock(unit["Designation"])
        for _, row in children_df.iterrows():
            interlock_df = self.templates_df["interlock"].copy()
            interlock_name = row["Path"].split(":")[-1]
            interlock_df = interlock_df.replace(
                [r'\$tag_name\$', r'\$interlock\$', r'\$plc_num\$', r'\$addr\$', r'\$description\$'],
                [tag, interlock_name, plc, f"{str(row['IOType_3'])[:-2]}", row["DefaultText"]], regex=True
            )
            unit_df = pd.concat([unit_df, interlock_df], axis=0)

        return unit_df

    def ecs2wincc(self, wincc_file: str, point_type: str, plc: str, filter: str) -> pd.DataFrame:
        self.logger.info(f"Conversion ECS -> WinCC: -> {wincc_file}, Equipment type: {point_type}, PLC: {plc}, Filter: {filter}")
        print(f"{Fore.YELLOW}Converting ECS tags to WINCC: {Fore.MAGENTA}{Style.RESET_ALL}", end="\t")
        wincc_df = pd.DataFrame()

        ecs_tags_flt_df = self.ecs_tags_df[
            self.ecs_tags_df['PointType'].str.contains(point_type, case=False) &
            self.ecs_tags_df['IOType_0'].str.contains(plc, case=False, na=False) &
            self.ecs_tags_df['Designation'].str.contains(filter, case=False, na=False)
        ]
        qty_rows = 0
        for _, row in ecs_tags_flt_df.iterrows():
            if "SIMNONE" not in row["IOType_5"]:
                wincc_df = pd.concat([wincc_df, self.make_unit_df(row)], axis=0)
                qty_rows += 1

        wincc_df = wincc_df.astype(str)
        wincc_df["Limit Upper 2 Type"] = "None"
        wincc_df["Limit Lower 2 Type"] = "None"
        print(f"{Fore.GREEN}OK {Style.RESET_ALL}")
        print(f"{Fore.YELLOW}Qty tags: {Fore.MAGENTA}{qty_rows}{Style.RESET_ALL}")
        return wincc_df

    def open_templates(self):
        print(f"{Fore.YELLOW}Opening templates: {Fore.MAGENTA}{self._template_dir}{Style.RESET_ALL}", end="\t")
        self.logger.info(f"Opening templates: {self._template_dir}")
        try:
            self.templates_df["motor"] = pd.read_excel(self._template_dir / "wincc_motor_template.xlsx").astype(str)
            self.templates_df["valve"] = pd.read_excel(self._template_dir / "wincc_valve_template.xlsx").astype(str)
            self.templates_df["analog"] = pd.read_excel(self._template_dir / "wincc_analog_template.xlsx").astype(str)
            self.templates_df["interlock"] = pd.read_excel(self._template_dir / "wincc_interlock_template.xlsx").astype(str)
        except Exception as e:
            self.logger.error(f"Error reading templates: {e}")
            raise
        print(f"{Fore.GREEN}OK {Style.RESET_ALL}")

    def open_ecs_tags_xlsx(self, ecs_file: str):
        print(f"{Fore.YELLOW}Opening ECS tags: {Fore.MAGENTA}{ecs_file}{Style.RESET_ALL}", end="\t")
        self.logger.info(f"Opening ECS tags: {ecs_file}")
        try:
            self.ecs_tags_df = pd.read_excel(ecs_file).astype(str)
            self.logger.info(f"Read ECS tags: {ecs_file}")
        except Exception as e:
            self.logger.error(f"Error read ECS tags: {e}")
            raise Exception(f"Error read ECS tags: {e}")
        print(f"{Fore.GREEN}OK {Style.RESET_ALL}")

    def write_wincc_xlsx(self, wincc_df: pd.DataFrame, wincc_file: str):
        print(f"{Fore.YELLOW}Saving to: {Fore.MAGENTA}{str(wincc_file)} {Style.RESET_ALL}", end="\t")
        try:
            with pd.ExcelWriter(wincc_file, engine="xlsxwriter") as writer:
                wincc_df.to_excel(writer, index=False, header=True, sheet_name="Hmi Tags")
            print(f"{Fore.GREEN}OK {Style.RESET_ALL}")
        except Exception as e:
            print(f"{Fore.RED}FAULT{Style.RESET_ALL}")
            err_msg = f"{Fore.RED}Error: Can't write tags to: {str(wincc_file)} \n\r {e}"
            print(err_msg)
            raise Exception(err_msg)

    def run(self, ecs_file: str, wincc_file: str, point_type: str, plc: str, tag_filter: str):
        print(f"{Fore.GREEN}Converse file: {Fore.MAGENTA}{ecs_file} "
              f"{Fore.GREEN} PLC: {Fore.MAGENTA}{plc} "
              f"{Fore.GREEN}Point_type: {Fore.MAGENTA}{point_type} "
              f"{Fore.GREEN}Filter: {Fore.MAGENTA}{tag_filter}{Style.RESET_ALL}")
        with alive_bar(4, force_tty=True, length=24, bar='classic', spinner='classic') as bar:
            self.open_templates()
            bar()
            self.open_ecs_tags_xlsx(ecs_file)
            bar()
            wincc_df = self.ecs2wincc(wincc_file, point_type, plc, tag_filter)
            bar()
            self.write_wincc_xlsx(wincc_df, wincc_file)
            bar()
        print(f"{Fore.GREEN}Converse complete.{Style.RESET_ALL}", end="\t")

def main():
    parser = argparse.ArgumentParser(
        prog="ecs2wincc",
        description="Conversion of tag export from XLSX ECS to import format XLSX WinCC Prof",
        epilog=f'2024 7Art v0.1'
    )
    parser.add_argument("-v", "--version", action="version", version="Version 0.1")
    parser.add_argument("file", type=str, default="Points.xlsx", help="Name of the file with ECS tags (Default Points.xlsx)")
    parser.add_argument("-o", "--output", type=str, default="tags_wincc.xlsx", help="Name of the file to save WinCC data (Default: tags_wincc.xlsx)")
    parser.add_argument("-t", "--point_type", type=str, default="unimotor",
                        choices=["unimotor","valve","analog"],
                        help="Equipment type (unimotor|valve|analog) (Default: unimotor)")
    parser.add_argument("-p", "--plc", type=str, default="994", help="PLC (Default: 994)")
    parser.add_argument("-ftr", "--filter", type=str, default=".*", help="Tag filter (Default: .*)")

    args = parser.parse_args()
    converter = ECS2WinCCConverter()
    converter.run(args.file, args.output, args.point_type, args.plc, args.filter)

if __name__ == "__main__":
    main()