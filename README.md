# SPAS - Smart Proposal Automation System

A Streamlit-based intelligent document automation platform designed to streamline proposal generation for industrial automation and material handling systems. SPAS transforms technical specifications, DXF files, and engineering data into professional, comprehensive proposal documents in minutes.

## Project Overview

SPAS automates the entire proposal creation workflow for complex industrial automation projects. It processes AutoCAD DXF files, extracts component specifications, performs capacity calculations, and generates publication-ready DOCX documents with consistent formatting, technical specifications, pricing tables, and client branding.

The system supports four distinct product lines:
- CBS (Conveyor Belt Systems)
- Neo (Next-generation Sorting Systems)
- Cubizone (Cubic Storage and Retrieval Systems)
- Robodome 2.0 (Robotic Automation Solutions)

## Key Features

### Intelligent DXF Processing
- Automatic extraction of components from AutoCAD DXF files
- Component mapping and specification analysis
- Dimension and layout identification
- Support for conveyors, diverters, scanners, and automation equipment

### Automated Document Generation
- Professional cover pages with client logo integration
- Executive summaries and personalized cover letters
- Detailed system descriptions and technical specifications
- Component catalogs with manufacturer details
- Process flow diagrams and layout overviews
- Capacity calculations and performance metrics
- Commercial terms and pricing tables
- Warranty configurations and exclusions
- Safety protocols and compliance documentation
- Glossary and technical terminology

### Dynamic Content Management
- Client-specific customization
- Real-time capacity calculations based on system parameters
- Flexible warranty and exclusion configurations
- Manufacturer and component database management
- Reference project integration
- Multi-language support capabilities

### AI-Powered Content Generation
- GROQ API integration for intelligent text generation
- Context-aware system descriptions
- Process flow optimization
- Consistency validation across sections
- Semantic search for reference materials

### User-Friendly Interface
- Intuitive tab-based navigation with persistent state
- Drag-and-drop file upload support
- Interactive configuration panels
- Real-time preview and validation
- Session state management for seamless workflow
- Role-based access control with user authentication

## System Requirements

### Software Requirements
- Python 3.8 or higher
- Windows OS (for docx2pdf conversion)
- Microsoft Word (optional, for DOCX to PDF conversion)

### Hardware Requirements
- Minimum 8GB RAM (16GB recommended for large DXF files)
- 2GB free disk space
- Internet connection for API services

## Technology Stack

### Core Framework
- Streamlit - Web application framework
- Python 3.x - Backend programming language

### Document Processing
- python-docx - DOCX document generation
- docxcompose - Document composition and merging
- ezdxf - DXF file parsing and analysis
- openpyxl - Excel file processing
- pdfplumber - PDF text extraction
- docx2pdf - DOCX to PDF conversion
- convertapi - Document conversion service

### AI and Machine Learning
- groq - AI API for text generation
- sentence-transformers - Semantic similarity analysis
- bert-score - Text quality evaluation
- torch - PyTorch machine learning framework

### Data Processing
- pandas - Data manipulation and analysis
- numpy - Numerical computing
- PIL/Pillow - Image processing

### Utilities
- python-dotenv - Environment variable management
- requests - HTTP library for API calls
- pythoncom - Windows COM interface

## Installation

### Step 1: Clone the Repository
```bash
git clone <repository-url>
cd "V9 - Copy"
```

### Step 2: Create Virtual Environment
```bash
python -m venv venv
```

### Step 3: Activate Virtual Environment
Windows:
```bash
venv\Scripts\activate
```

Linux/Mac:
```bash
source venv/bin/activate
```

### Step 4: Install Dependencies
```bash
pip install -r requirements.txt
```

### Step 5: Configure Environment Variables
Create a `.env` file in the project root directory with the following variables:

```env
GROQ_API_KEY=your_groq_api_key_here
CONVERTAPI_SECRET=your_convertapi_secret_here
```

#### How to Obtain API Keys:

**GROQ API Key:**
1. Visit https://console.groq.com
2. Sign up or log in to your account
3. Navigate to API Keys section
4. Generate a new API key
5. Copy the key to your .env file

**ConvertAPI Secret:**
1. Visit https://www.convertapi.com
2. Sign up for a free account
3. Go to Dashboard > API Secret
4. Copy your secret key to your .env file

### Step 6: Prepare Required Directories
Ensure the following directories exist with appropriate content:

```
FIXED_IMAGE/
├── Chutes/
└── clients/
Static_AboutCompany/
assests/
├── Images/
│   └── clients/
Supportive_Functions/
Cubizone/
└── V1/
```

### Step 7: Verify Installation
```bash
streamlit --version
python -c "import streamlit; print('Streamlit installed successfully')"
```

## Running the Application

### Start the Application
```bash
streamlit run streamlit.py
```

The application will automatically open in your default web browser at:
```
http://localhost:8501
```

### Login Credentials

The system includes pre-configured user accounts for authentication:

**User 1 - Assistant Manager**
- Email: sanyog.singh@falconautotech.com
- Password: 1234

**User 2 - General Manager**
- Email: tuhi@falconautoonline.com
- Password: 1234

**User 3 - Deputy General Manager**
- Email: sandeep.pathak@falconautoonline.com
- Password: 1234

**User 4 - Principal Engineer**
- Email: balmukund.mishra@falconautotech.com
- Password: 1234

**User 5 - Senior Engineer**
- Email: arkajyoti.chakraborty@falconautotech.com
- Password: 1234

## Usage Workflow

### 1. Login
- Open the application in your browser
- Enter your email and password
- Click "Login" to access the main interface

### 2. Select System Type
- Choose from the available tabs: CBS, Neo, Cubizone, or Robodome 2.0
- Each system has specialized configuration options

### 3. Enter Project Details (CBS/Neo)
- Client name and project information
- System specifications and requirements
- Upload DXF files for automatic component extraction
- Upload costing sheets (Excel format)
- Upload capacity calculation documents
- Optional: Add layout images

### 4. Configure Settings
- Parcel spectrum (dimensions and weight ranges)
- Commercial settings (payment terms, delivery)
- Warranty configuration (standard or comprehensive)
- Exclusions and special conditions
- Key components and manufacturers

### 5. Generate Proposal
- Click "Generate Proposal" button
- System processes all inputs and generates document
- Preview appears with PDF view
- Download button available for DOCX format

### 6. Cubizone Workflow
- Select Cubizone type (Manual/Semi-Auto/Automatic)
- Configure zone dimensions and specifications
- Set pricing and commercial terms
- Generate customized Cubizone proposal

## Project Structure

```
V9 - Copy/
├── streamlit.py                    # Main application entry point
├── requirements.txt                # Python dependencies
├── README.md                       # This file
├── .env                           # Environment variables (create this)
├── .gitignore                     # Git ignore rules
├── component_catalog.json         # Component database
├── component_alias_map.json       # Component name mappings
├── DXF_COMPONENT_NAME_MAP.txt     # DXF component mapping
├── CBS_SYSTEM_DESC.txt            # CBS system templates
├── Supportive_Functions/          # Core business logic modules
│   ├── st_sys_desc.py            # System description generation
│   ├── combine_old.py            # Document assembly
│   ├── agentY.py                 # AI-powered process flow
│   ├── dxf_extractor.py          # DXF file processing
│   ├── costing_sheet_mapper.py   # Cost analysis
│   ├── proposal_facts.py         # Fact extraction
│   ├── proposal_facts_extractor.py # Fact processing
│   ├── proposal_context.py       # Context management
│   ├── dynamic_system_description.py # Dynamic content
│   ├── consistency_audit.py      # Quality validation
│   ├── alias_map.py              # Component aliasing
│   ├── bom.py                    # Bill of materials
│   └── feedback_rules.json       # Validation rules
├── Cubizone/                      # Cubizone module
│   └── V1/
│       └── main.py               # Cubizone proposal generator
├── FIXED_IMAGE/                   # Static images and assets
│   ├── Chutes/                   # Component images
│   └── clients/                  # Client logos
├── Static_AboutCompany/           # Company profile content
└── assests/                       # Additional assets
    └── Images/
        └── clients/              # Client-specific assets
```

## Configuration Files

### component_catalog.json
Contains detailed specifications for all supported components including conveyors, scanners, diverters, and automation equipment.

### component_alias_map.json
Maps various naming conventions to standardized component names for consistent identification.

### DXF_COMPONENT_NAME_MAP.txt
Defines patterns for recognizing components in DXF files.

### feedback_rules.json
Validation rules for content consistency and quality assurance.

## Environment Variables Reference

Required environment variables in `.env` file:

| Variable | Description | Required | Example |
|----------|-------------|----------|---------|
| GROQ_API_KEY | API key for GROQ AI service | Yes | gsk_xxxxxxxxxxxxx |
| CONVERTAPI_SECRET | Secret for ConvertAPI service | Yes | secret_xxxxxxxxxx |

## Troubleshooting

### Common Issues

**Issue: GROQ_API_KEY not found**
- Solution: Ensure .env file exists in project root with valid GROQ_API_KEY

**Issue: DXF file not parsing correctly**
- Solution: Verify DXF file is valid AutoCAD format (DXF R12-R2018)
- Check component naming matches patterns in DXF_COMPONENT_NAME_MAP.txt

**Issue: Document generation fails**
- Solution: Ensure all required fields are filled
- Check that uploaded files are in correct format (DXF, XLSX)
- Verify sufficient disk space for temporary files

**Issue: PDF preview not displaying**
- Solution: Check browser compatibility (Chrome/Edge recommended)
- Ensure ConvertAPI credentials are valid

**Issue: Login not working**
- Solution: Use exact email addresses listed in Login Credentials section
- Ensure password is exactly "1234" (case-sensitive)

### Performance Optimization

For large DXF files (>10MB):
- Increase system RAM allocation
- Close unnecessary applications
- Consider splitting into smaller sections

For faster generation:
- Pre-upload all files before clicking Generate
- Use wired internet connection for API calls
- Keep only necessary tabs open

## Development Notes

### Adding New Users
Edit USER_CREDENTIALS dictionary in streamlit.py (around line 4986):
```python
USER_CREDENTIALS = {
    "new.user@domain.com": {
        "name": "User Name",
        "role": "Role Title",
        "password": "password",
        "initials": "UN"
    }
}
```

### Modifying Component Database
Edit component_catalog.json to add or update component specifications.

### Custom System Descriptions
Modify CBS_SYSTEM_DESC.txt or add new templates in Supportive_Functions/.

## API Usage and Limits

### GROQ API
- Used for: AI text generation, system descriptions
- Rate limits: Check GROQ documentation
- Cost: Based on token usage

### ConvertAPI
- Used for: DOCX to PDF conversion
- Free tier: Limited conversions per month
- Upgrade: Available for production use

## Security Considerations

- Store .env file securely, never commit to version control
- Change default passwords for production deployment
- Restrict file upload sizes to prevent DOS attacks
- Validate all user inputs before processing
- Use HTTPS in production environment

## Support and Maintenance

### Logging
Application logs are stored in memory during runtime. Check terminal output for detailed error messages.

### Backup Recommendations
- Regular backup of component_catalog.json
- Backup client logos and assets
- Archive generated proposals periodically

## License

Internal use only. Proprietary software for FALCON Autotech.

## Contact

For support or questions, contact:
- Technical Lead: sanyog.singh@falconautotech.com
- General Manager: tuhi@falconautoonline.com

## Version History

### V9 (Current)
- Multi-system support (CBS, Neo, Cubizone, Robodome 2.0)
- AI-powered content generation
- Enhanced DXF processing
- Improved user interface with persistent navigation
- Role-based authentication
- Real-time document preview

---

Last Updated: January 2026
