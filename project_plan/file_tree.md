ppt_to_h5p/
│── main.py # Application entry point
│── gui.py # Tkinter GUI layout and event handlers
│── converter/
│ ├── **init**.py
│ ├── ppt_exporter.py # PowerPoint COM automation
│ └── h5p_builder.py # H5P JSON schema & ZIP archiving
└── utils/
├── **init**.py
└── logger.py # Centralized logging setup & path resolution
