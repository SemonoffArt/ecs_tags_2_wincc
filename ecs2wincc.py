# Import required libraries
import logging  # For logging functionality
import pandas as pd  # For data manipulation and analysis
import argparse  # For command-line argument parsing
from pathlib import Path  # For path manipulation
from alive_progress import alive_bar  # For progress bar visualization
from colorama import init, Fore, Style  # For colored console output
from typing import Dict, Optional, Union
from dataclasses import dataclass
from enum import Enum, auto
import sys

class PointType(Enum):
    MOTOR = auto()
    VALVE = auto()
    ANALOG = auto()
    UNKNOWN = auto()

@dataclass
class ConversionConfig:
    """Configuration for ECS to WinCC conversion"""
    input_file: str
    output_file: str
    point_type: str
    plc: str
    tag_filter: str

class ECS2WinCCConverter:
    """
    Main class for converting ECS tags to WinCC format.
    Handles the conversion process from ECS Excel files to WinCC-compatible Excel files.
    """
    def __init__(self, version: float = 0.1) -> None:
        """
        Initialize the converter with default settings and paths.
        
        Args:
            version (float): Version number of the converter
        """
        self._version = version
        self._prg_dir = Path("./").absolute()  # Get absolute path of current directory
        self._resources_dir = self._prg_dir / "resources"  # Path to resources directory
        self._template_dir = self._resources_dir / "templates"  # Path to template files
        self._log_file = self._prg_dir / "ecs2wincc.log"  # Path to log file
        self.templates_df: Dict[str, pd.DataFrame] = {}  # Dictionary to store template dataframes
        self.ecs_tags_df = pd.DataFrame()  # DataFrame to store ECS tags
        self._init_logging()  # Initialize logging
        self._init_colorama()  # Initialize colorama for colored output

    def _init_logging(self) -> None:
        """
        Configure logging settings with timestamp, level, and message format.
        """
        logging.basicConfig(
            format="%(asctime)s - %(levelname)s -(%(funcName)s(%(lineno)d) - %(message)s",
            level=logging.INFO,
            filename=self._log_file,
            encoding="UTF-8"
        )
        self.logger = logging.getLogger()

    def _init_colorama(self) -> None:
        """
        Initialize colorama for colored console output.
        """
        init(autoreset=True)

    def get_parent_info(self, parent_tag_name: str) -> str:
        """
        Get information about a parent tag from the ECS tags DataFrame.
        
        Args:
            parent_tag_name (str): Name of the parent tag
            
        Returns:
            str: Combined parent tag information
            
        Raises:
            ValueError: If parent tag is not found
        """
        try:
            parent_info = f"{parent_tag_name} {self.ecs_tags_df.loc[self.ecs_tags_df['Designation'].str.contains(parent_tag_name, case=False), 'DefaultText'].values[0]}"
            return parent_info
        except IndexError:
            self.logger.error(f"Parent tag not found: {parent_tag_name}")
            raise ValueError(f"Parent tag not found: {parent_tag_name}")

    def get_children_interlock(self, parent_tag_name: str) -> pd.DataFrame:
        """
        Get all interlock children for a given parent tag.
        
        Args:
            parent_tag_name (str): Name of the parent tag
            
        Returns:
            pd.DataFrame: DataFrame containing interlock children
        """
        self.logger.info(f"Get children for: {parent_tag_name}")
        try:
            mask = self.ecs_tags_df['FunctionalHierarchy'].str.contains(parent_tag_name, case=False) & (
                self.ecs_tags_df['Designation'].str.contains(r"int\d+|stp\d+|str\d+", case=False, regex=True)
            )
            children_df = self.ecs_tags_df[mask]
            self.logger.info(f"Found children: {len(children_df)}")
            return children_df
        except Exception as e:
            self.logger.error(f"Error getting children for {parent_tag_name}: {e}")
            raise

    def extract_point_type(self, point_type: str) -> PointType:
        """
        Extract and standardize point type from the input string.
        
        Args:
            point_type (str): Raw point type string
            
        Returns:
            PointType: Standardized point type enum
        """
        point_type = point_type.lower()
        if "motor" in point_type:
            return PointType.MOTOR
        elif "valve" in point_type:
            return PointType.VALVE
        elif "analog" in point_type:
            return PointType.ANALOG
        else:
            return PointType.UNKNOWN

    def get_decimal_format(self, decimals: str) -> str:
        """
        Convert decimal number to WinCC format string.
        
        Args:
            decimals (str): Number of decimal places
            
        Returns:
            str: Formatted decimal string for WinCC
            
        Raises:
            ValueError: If decimals cannot be converted to integer
        """
        try:
            decimals = int(float(decimals))
            formats = {
                0: "s999999",
                1: "s999999.9",
                2: "s999999.99",
                3: "s999999.999"
            }
            return formats.get(decimals, "s9999999")
        except (ValueError, TypeError) as e:
            self.logger.error(f"Error converting decimals: {e}")
            raise ValueError(f"Invalid decimal format: {decimals}")

    def make_unit_df(self, unit: pd.Series) -> pd.DataFrame:
        """
        Create a DataFrame for a single unit with all its interlocks.
        
        Args:
            unit (pd.Series): Unit data from ECS tags
            
        Returns:
            pd.DataFrame: DataFrame containing unit and interlock information
            
        Raises:
            ValueError: If required unit data is missing
        """
        try:
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

            # Create base unit DataFrame from template
            if point_type not in self.templates_df:
                raise ValueError(f"Template not found for point type: {point_type}")
                
            unit_df = self.templates_df[point_type.name.lower()].copy()
            replacements = {
                r'\$tag_name\$': tag,
                r'\$plc_num\$': plc,
                r'\$dbnum\$': db_num,
                r'\$parent_info\$': parent_info,
                r'\$description\$': description,
                r'\$eu\$': eu,
                r'\$decimals\$': decimals,
                r'\$trend_tag_name\$': trend_tag_name
            }
            
            for pattern, replacement in replacements.items():
                unit_df = unit_df.replace(pattern, replacement, regex=True)

            # Add interlock information
            children_df = self.get_children_interlock(unit["Designation"])
            for _, row in children_df.iterrows():
                interlock_df = self.templates_df["interlock"].copy()
                interlock_name = row["Path"].split(":")[-1]
                interlock_replacements = {
                    r'\$tag_name\$': tag,
                    r'\$interlock\$': interlock_name,
                    r'\$plc_num\$': plc,
                    r'\$addr\$': f"{str(row['IOType_3'])[:-2]}",
                    r'\$description\$': row["DefaultText"]
                }
                
                for pattern, replacement in interlock_replacements.items():
                    interlock_df = interlock_df.replace(pattern, replacement, regex=True)
                    
                unit_df = pd.concat([unit_df, interlock_df], axis=0, ignore_index=True)

            return unit_df
        except KeyError as e:
            self.logger.error(f"Missing required unit data: {e}")
            raise ValueError(f"Missing required unit data: {e}")
        except Exception as e:
            self.logger.error(f"Error creating unit DataFrame: {e}")
            raise

    def validate_config(self, config: ConversionConfig) -> None:
        """
        Validate conversion configuration.
        
        Args:
            config (ConversionConfig): Configuration to validate
            
        Raises:
            ValueError: If configuration is invalid
        """
        if not config.input_file:
            raise ValueError("Input file path is required")
            
        if not config.output_file:
            raise ValueError("Output file path is required")
            
        if not config.point_type:
            raise ValueError("Point type is required")
            
        if not config.plc:
            raise ValueError("PLC number is required")
            
        if not config.tag_filter:
            raise ValueError("Tag filter is required")

    def ecs2wincc(self, config: ConversionConfig) -> pd.DataFrame:
        """
        Convert ECS tags to WinCC format.
        
        Args:
            config (ConversionConfig): Conversion configuration
            
        Returns:
            pd.DataFrame: DataFrame containing converted WinCC tags
            
        Raises:
            ValueError: If conversion fails
        """
        self.logger.info(f"Conversion ECS -> WinCC: -> {config.output_file}, Equipment type: {config.point_type}, PLC: {config.plc}, Filter: {config.tag_filter}")
        print(f"{Fore.YELLOW}Converting ECS tags to WINCC: {Fore.MAGENTA}{Style.RESET_ALL}", end="\t")
        
        try:
            # Filter ECS tags based on criteria
            ecs_tags_flt_df = self.ecs_tags_df[
                self.ecs_tags_df['PointType'].str.contains(config.point_type, case=False) &
                self.ecs_tags_df['IOType_0'].str.contains(config.plc, case=False, na=False) &
                self.ecs_tags_df['Designation'].str.contains(config.tag_filter, case=False, na=False)
            ]
            
            if ecs_tags_flt_df.empty:
                self.logger.warning("No tags found matching the criteria")
                return pd.DataFrame()
            
            # Process each tag
            wincc_df = pd.DataFrame()
            qty_rows = 0
            
            for _, row in ecs_tags_flt_df.iterrows():
                if "SIMNONE" not in row["IOType_5"]:
                    unit_df = self.make_unit_df(row)
                    wincc_df = pd.concat([wincc_df, unit_df], axis=0, ignore_index=True)
                    qty_rows += 1

            # Finalize DataFrame
            if not wincc_df.empty:
                wincc_df = wincc_df.astype(str)
                wincc_df["Limit Upper 2 Type"] = "None"
                wincc_df["Limit Lower 2 Type"] = "None"
                
            print(f"{Fore.GREEN}OK {Style.RESET_ALL}")
            print(f"{Fore.YELLOW}Qty tags: {Fore.MAGENTA}{qty_rows}{Style.RESET_ALL}")
            return wincc_df
            
        except Exception as e:
            self.logger.error(f"Error during conversion: {e}")
            raise ValueError(f"Conversion failed: {e}")

    def open_templates(self) -> None:
        """
        Load template Excel files for different point types.
        
        Raises:
            FileNotFoundError: If template files are not found
            ValueError: If template files are invalid
        """
        print(f"{Fore.YELLOW}Opening templates: {Fore.MAGENTA}{self._template_dir}{Style.RESET_ALL}", end="\t")
        self.logger.info(f"Opening templates: {self._template_dir}")
        
        template_files = {
            "motor": "wincc_motor_template.xlsx",
            "valve": "wincc_valve_template.xlsx",
            "analog": "wincc_analog_template.xlsx",
            "interlock": "wincc_interlock_template.xlsx"
        }
        
        try:
            for point_type, filename in template_files.items():
                template_path = self._template_dir / filename
                if not template_path.exists():
                    raise FileNotFoundError(f"Template file not found: {template_path}")
                    
                self.templates_df[point_type] = pd.read_excel(template_path).astype(str)
                if self.templates_df[point_type].empty:
                    raise ValueError(f"Empty template file: {template_path}")
                    
        except Exception as e:
            self.logger.error(f"Error reading templates: {e}")
            raise
        print(f"{Fore.GREEN}OK {Style.RESET_ALL}")

    def open_ecs_tags_xlsx(self, ecs_file: str) -> None:
        """
        Load ECS tags from Excel file.
        
        Args:
            ecs_file (str): Path to ECS tags Excel file
            
        Raises:
            FileNotFoundError: If ECS file is not found
            ValueError: If ECS file is invalid
        """
        print(f"{Fore.YELLOW}Opening ECS tags: {Fore.MAGENTA}{ecs_file}{Style.RESET_ALL}", end="\t")
        self.logger.info(f"Opening ECS tags: {ecs_file}")
        
        try:
            if not Path(ecs_file).exists():
                raise FileNotFoundError(f"ECS file not found: {ecs_file}")
                
            self.ecs_tags_df = pd.read_excel(ecs_file).astype(str)
            if self.ecs_tags_df.empty:
                raise ValueError(f"Empty ECS file: {ecs_file}")
                
            required_columns = ['Designation', 'PointType', 'IOType_0', 'IOType_6', 'FunctionalHierarchy', 'DefaultText', 'Unit', 'Decimals']
            missing_columns = [col for col in required_columns if col not in self.ecs_tags_df.columns]
            if missing_columns:
                raise ValueError(f"Missing required columns in ECS file: {missing_columns}")
                
            self.logger.info(f"Read ECS tags: {ecs_file}")
        except Exception as e:
            self.logger.error(f"Error reading ECS tags: {e}")
            raise
        print(f"{Fore.GREEN}OK {Style.RESET_ALL}")

    def write_wincc_xlsx(self, wincc_df: pd.DataFrame, wincc_file: str) -> None:
        """
        Write converted tags to WinCC Excel file.
        
        Args:
            wincc_df (pd.DataFrame): DataFrame containing WinCC tags
            wincc_file (str): Output file path
            
        Raises:
            ValueError: If DataFrame is empty
            IOError: If file cannot be written
        """
        print(f"{Fore.YELLOW}Saving to: {Fore.MAGENTA}{str(wincc_file)} {Style.RESET_ALL}", end="\t")
        
        if wincc_df.empty:
            raise ValueError("No data to write to WinCC file")
            
        try:
            with pd.ExcelWriter(wincc_file, engine="xlsxwriter") as writer:
                wincc_df.to_excel(writer, index=False, header=True, sheet_name="Hmi Tags")
                
                # Auto-adjust column widths
                worksheet = writer.sheets["Hmi Tags"]
                for idx, col in enumerate(wincc_df.columns):
                    max_length = max(
                        wincc_df[col].astype(str).apply(len).max(),
                        len(str(col))
                    )
                    worksheet.set_column(idx, idx, max_length + 2)
                    
            print(f"{Fore.GREEN}OK {Style.RESET_ALL}")
        except Exception as e:
            print(f"{Fore.RED}FAULT{Style.RESET_ALL}")
            err_msg = f"{Fore.RED}Error: Can't write tags to: {str(wincc_file)} \n\r {e}"
            print(err_msg)
            raise IOError(err_msg)

    def run(self, config: ConversionConfig) -> None:
        """
        Main method to run the conversion process.
        
        Args:
            config (ConversionConfig): Conversion configuration
            
        Raises:
            ValueError: If conversion fails
        """
        try:
            self.validate_config(config)
            
            print(f"{Fore.GREEN}Converting file: {Fore.MAGENTA}{config.input_file} "
                  f"{Fore.GREEN}PLC: {Fore.MAGENTA}{config.plc} "
                  f"{Fore.GREEN}Point type: {Fore.MAGENTA}{config.point_type} "
                  f"{Fore.GREEN}Filter: {Fore.MAGENTA}{config.tag_filter}{Style.RESET_ALL}")
                  
            with alive_bar(4, force_tty=True, length=24, bar='classic', spinner='classic') as bar:
                self.open_templates()
                bar()
                self.open_ecs_tags_xlsx(config.input_file)
                bar()
                wincc_df = self.ecs2wincc(config)
                bar()
                self.write_wincc_xlsx(wincc_df, config.output_file)
                bar()
                
            print(f"{Fore.GREEN}Conversion complete.{Style.RESET_ALL}")
            
        except Exception as e:
            self.logger.error(f"Conversion failed: {e}")
            raise ValueError(f"Conversion failed: {e}")

def main() -> None:
    """
    Main function to parse command line arguments and run the converter.
    """
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

    try:
        args = parser.parse_args()
        config = ConversionConfig(
            input_file=args.file,
            output_file=args.output,
            point_type=args.point_type,
            plc=args.plc,
            tag_filter=args.filter
        )
        
        converter = ECS2WinCCConverter()
        converter.run(config)
        
    except Exception as e:
        print(f"{Fore.RED}Error: {e}{Style.RESET_ALL}")
        sys.exit(1)

if __name__ == "__main__":
    main()
