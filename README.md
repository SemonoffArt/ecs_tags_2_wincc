# ECS2WinCC - Point Converter from FLS ECS SCADA to SIEMENS WinCC

A powerful console utility that converts FLS ECS points to SIEMENS WinCC format, making it easy to migrate your automation system data.

## 🚀 Features

- 🔄 Convert FLS ECS Points (Excel export) to WinCC format
- 🏭 Support for multiple equipment types:
  - Motors
  - Valves
  - Analog
- 🔍 Advanced tag filtering capabilities
- 📝 Detailed conversion logging
- 📊 Real-time progress tracking with `alive_progress`

## 📋 Prerequisites

- Python 3.x
- Required Python packages (see requirements.txt)

## 💻 Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/ecs2wincc.git
   cd ecs2wincc
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## 🛠 Usage

Basic command:
```bash
python ecs2wincc.py [options]
```

### Command Line Options

| Option | Description | Default |
|--------|-------------|---------|
| `file` | Input file with ECS tags | Points.xlsx |
| `-o, --output` | Output file for WinCC data | tags_wincc.xlsx |
| `-t, --point_type` | Equipment type (motor\|valve\|analog) | motor |
| `-p, --plc` | PLC number | 994 |
| `-ftr, --filter` | Tag filter pattern | .* |

### Examples

Convert motor points:
```bash
python ecs2wincc.py Points.xlsx -o tags_wincc.xlsx -t motor -p 994 -ftr .*
```

Convert valve points:
```bash
python ecs2wincc.py Points.xlsx -o tags_wincc.xlsx -t valve -p 994 -ftr .*
```

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📫 Support

For support, please open an issue in the GitHub repository.