# SW2 Data Import

A comprehensive suite of tools designed to facilitate data importing into Softwrench through both Maximo and Softwrench APIs. These utilities streamline the process of data integration, enabling seamless transfers and synchronization between systems.

## 🚀 Features

### 1. Maximo Data Importer
A GUI application for importing data directly into IBM Maximo through OSLC REST API.
- Support for CSV and JSON input files
- Multiple operation modes (Create, Bulk Create, Update, Merge Update, Delete)
- Real-time progress tracking
- Detailed operation summaries
- Secure authentication handling

### 2. Softwrench Data Importer (Future Plan)
Tools for importing data directly through Softwrench's API.

## 📋 Prerequisites and Installation

- Python 3.7 or higher
- Access to valid Maximo instance with OSLC API enabled
- Valid Maximo credentials
- Valid Softwrench credentials (for Softwrench imports)
- For more specific details, see the readme inside the folder of each program.

## 🚦 Getting Started

### Maximo Data Importer

1. Navigate to the Maximo importer directory:
```bash
cd src/maximo_data_import/2.\ send\ to\ maximo/
```

2. Run the application:
```bash
python maximo_sender_ui.py
```

3. Follow the GUI prompts to:
   - Select your data file (CSV/JSON)
   - Choose operation type
   - Configure Maximo connection
   - Start the import process

## �� Project Structure

sw2-data-import/
├── src/
│ ├── maximo_data_import/
│ │ ├── 2. send to maximo/
│ │ │ ├── maximo_sender_ui.py # GUI application
│ │ │ ├── maximo_sender.py # Core Maximo operations
│ │ │ ├── csv_to_json.py # CSV conversion utility
│ │ │ └── README.md # Detailed Maximo importer docs
│ │ └── ...
│ └── softwrench_data_import/ # (Coming Soon)
├── requirements.txt
└── README.md

## 📝 Documentation

- Detailed documentation for each tool can be found in their respective directories
- See `src/maximo_data_import/2. send to maximo/README.md` along with `src/maximo_data_import/instructions.md` for complete Maximo importer documentation

## 🤝 Contributing

Contributions are welcome! Please feel free to submit pull requests or create issues for:
- Bug reports
- Feature requests
- Documentation improvements
- Code optimizations

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🆘 Support

For support:
1. Check the tool-specific documentation
2. Review existing issues
3. Create a new issue with detailed information about your problem

## 🔄 Updates

- [x] Maximo Data Importer with GUI
- [ ] Softwrench Data Importer
- [ ] Additional data validation features
- [ ] Enhanced error reporting
- [ ] Batch operation improvements
