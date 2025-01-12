# **ecs2wincc** - Point Converter from FLS ECS to SIEMENS WINCC

Console utility **ecs2wincc** creates an Excel file for importing points of Acesys objects to TIA WinCC based on the Excel file for exporting ECS points.

## Features

- Converts FLS ECS Points (export to excel ) to WinCC format.
- Supports different equipment types: motor, valve, and analog.
- Filters tags based on user-defined criteria.
- Logs the conversion process.
- Provides progress updates using `alive_progress`.

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/yourusername/ecs2wincc.git
    ```
2. Navigate to the project directory:
    ```sh
    cd ecs2wincc
    ```
3. Install the required dependencies:
    ```sh
    pip install -r requirements.txt
    ```

## Usage

Run the converter with the following command:
```sh
python ecs2wincc.py [options]
```

## Options
file: Name of the file with ECS tags (Default: Points.xlsx)
-o, --output: Name of the file to save WinCC data (Default: tags_wincc.xlsx)
-t, --point_type: Equipment type (motor|valve|analog) (Default: motor)
-p, --plc: PLC (Default: 994)
-ftr, --filter: Tag filter (Default: .*)


## Example
```sh
python ecs2wincc.py Points.xlsx -o tags_wincc.xlsx -t motor -p 994 -ftr .*
```

## License
This project is licensed under the MIT License - see the LICENSE file for details.