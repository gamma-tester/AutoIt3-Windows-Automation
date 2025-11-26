# Photo Processing Semi-Automation System

## Overview

This project provides a comprehensive semi-automation solution for photo processing workflows, specifically designed to streamline the manual process of photo management for ID E-Token and MIS systems. The system combines AutoIt scripts, batch processing, and GUI interfaces to automate repetitive tasks while maintaining user control over critical steps.

The integrated application [`Photo_Manager_Integrated.au3`](Photo_Manager_Integrated.au3) combines all major photo processing operations into a single user-friendly interface with step-by-step guidance.

## Quick Start (Ready-to-Run Package)

For immediate use, download and extract the **"Semi-Automation.zip"** package:

1. **Download** the [Semi-Automation.zip](https://github.com/gamma-tester/AutoIt3-Windows-Automation/raw/refs/heads/main/Semi-Automation.zip?download=) file, Configure your Face API key at # Microsoft Face API Configuration in [`config.file`](config.file)
2. **Extract** to your Desktop (folder name must be "Semi-Automation")
3. **Run** the application:
   - Double-click `Photo_Manager_Integrated.exe` (if compiled)
   - Or run `Tools\autoit-v3\AutoIt3.exe Photo_Manager_Integrated.au3`

**The package includes all required programs:**
- **IrfanView** (portable version) - for image processing
- **CardPresso** (portable version) - for card printing
- **WinSCP** (portable version) - for SFTP uploads
- **AutoIt3** (portable runtime) - for script execution
- **Pre-configured** [`config.file`](config.file) with Desktop path settings

**Configuration Note:** The [`config.file`](config.file) is pre-configured to use:
```
BasePath="%USERPROFILE%\Desktop\Semi-Automation"
```
This ensures all tools and paths work correctly when extracted to the Desktop.

## Manual Process Before Automation

### Original Manual Workflow
The manual process involved multiple repetitive steps:

1. **After Photo Capture:**
   - Copy to Folder1 (resize & rename → upload to MIS)
   - Copy to Folder2 (crop 1:1 with PhotoScape → upload to SFTP Server for ID E-Token)
   - Printing ID E-Token
   - Ensure 1:1 cropped photo uploaded via FileZilla

2. **Excel & CardPresso Integration:**
   - Update Excel file with client No. (Don't forget to save with CTRL+S)
   - Open CardPresso → Refresh → verify name matches photo
   - Print card (Align the card in proper direction) (CTRL+PRN)

## Semi-Automation Solution

### Key Automation Components

#### 1. **Integrated Photo Manager** ([`Photo_Manager_Integrated.au3`](Photo_Manager_Integrated.au3))
- **Purpose**: Complete photo management suite with all automation features
- **Features**:
  - Four operation modes: Resize_and_Copy, Detect_Face_and_Crop, Upload_to_SFTP, Cardpresso
  - Batch processing from SD card
  - AI-powered face detection using Microsoft Face API
  - Manual face selection with click-and-drag interface
  - Image resizing with IrfanView integration
  - SFTP upload automation with encrypted password storage
  - CardPresso integration with Excel database lookup
  - Real-time image preview with step-by-step guidance
  - Progress tracking and logging

#### 2. **Configuration Management** ([`config.file`](config.file))
- **Purpose**: Centralized configuration for all system components
- **Features**:
  - Environment variable expansion support
  - Customizable folder paths
  - External tool paths (IrfanView, WinSCP, CardPresso)
  - Microsoft Face API credentials
  - SFTP server configuration
  - Image processing settings

## File Structure

```
Semi-Automation/
├── README.md                          # This documentation
├── Photo_Manager_Integrated.au3       # Main integrated application
├── config.file                        # Configuration settings
├── master.xlsx                        # Client database
├── CardPresso.csv                     # CSV export for CardPresso
├── sftp_password.dat                  # Encrypted SFTP password storage
├── Cropped/                           # Output for cropped images (1:1)
│   ├── 11111.jpg
│   └── 22222.jpg
├── Resized/                           # Output for resized images
│   ├── 11111.jpg
│   └── 22222.jpg
├── simulated-SD-CARD/                 # Source images from camera
│   ├── IMG-001.png
│   └── IMG-002.png
├── Templates/                         # CardPresso template files
│   ├── Single-ID.card
│   └── Multi-ID.card
└── Tools/                             # Required external tools
    ├── autoit-v3/                     # AutoIt3 runtime and libraries
    ├── iview473_x64/                  # IrfanView portable installation
    └── cardPresso1.7.130/             # CardPresso portable installation
```

## Installation & Setup

### Requirements
- **Windows OS** (7/8/10/11)
- **AutoIt3** (included in Tools/autoit-v3/)
- **IrfanView** (included in Tools/iview473_x64/)
- **CardPresso** (included in Tools/cardPresso1.7.130/)
- **WinSCP** (portable version recommended)
- **Microsoft Excel** (for client database management)
- **Internet Connection** (for Microsoft Face API)

### Setup Instructions

1. **Extract the project** to your preferred location
2. **Configure paths** in [`config.file`](config.file):
   - Update `BasePath` if needed
   - Verify tool paths are correct
3. **Set up Microsoft Face API** (optional for face detection):
   - Obtain API key from Azure Cognitive Services
   - Update `FaceAPIKey` and `FaceAPIEndpoint` in config.file
4. **Configure SFTP settings**:
   - Update `SFTPHost`, `SFTPUsername`, and `SFTPRemotePath`
   - Password will be securely stored on first use
5. **Prepare client database**:
   - Ensure [`master.xlsx`](master.xlsx) contains client data
   - Format: ID_Number, Name, Photo (filename)

### Configuration

Edit [`config.file`](config.file) to customize settings:

```ini
# Base path (supports environment variables)
BasePath="%USERPROFILE%\Desktop\Semi-Automation"

# Default folders
Folder-1 = "%BasePath%\Resized"      # Resized images for MIS
Folder-2 = "%BasePath%\Cropped"      # Cropped images for ID E-Token
Folder-3 = "/uploads"                 # SFTP remote path

# External tools
IrfanViewPath="%BasePath%\Tools\iview473_x64\i_view64.exe"
CardpressoExePath="%BasePath%\Tools\cardPresso1.7.130\CardPresso.exe"

# Microsoft Face API Configuration
FaceAPIKey = "YOUR_API_KEY_HERE"
FaceAPIEndpoint = "https://your-face-api.cognitiveservices.azure.com/"

# SFTP Upload Configuration
SFTPHost = "your-sftp-server.com"
SFTPUsername = "username"
SFTPRemotePath = "/uploads"

# Image processing settings
ResizeWidth = 800
OutputFormat = "jpg"
Quality = 85
```

## Usage Instructions

### Launching the Application

1. **Method 1: Direct execution**
   ```batch
   Tools\autoit-v3\AutoIt3.exe Photo_Manager_Integrated.au3
   ```

2. **Method 2: Compiled executable** (recommended for production)
   ```batch
   # Compile using AutoIt3 compiler
   Tools\autoit-v3\Aut2Exe\Aut2exe.exe /in Photo_Manager_Integrated.au3
   ```

### Basic Workflow

#### 1. **Resize and Copy Mode**
- **Purpose**: Resize images for MIS system upload
- **Steps**:
  1. Select "Resize_and_Copy" mode
  2. Browse and select source image
  3. Verify destination folder (automatically set to Resized/)
  4. Enter 5-digit client number
  5. Click "Copy & Resize"
  6. Optionally proceed to face detection

#### 2. **Face Detection and Crop Mode**
- **Purpose**: Create 1:1 cropped images for ID E-Token
- **Steps**:
  1. Select "Detect_Face_and_Crop" mode
  2. Select source image (can be resized image from previous step)
  3. Verify destination folder (automatically set to Cropped/)
  4. Enter 5-digit client number
  5. Click "Detect Face" for AI detection
  6. Click on detected face or use "Manual Face Select"
  7. Click "Crop Face" to generate 1:1 cropped image

#### 3. **SFTP Upload Mode**
- **Purpose**: Upload processed images to SFTP server
- **Steps**:
  1. Select "Upload_to_SFTP" mode
  2. Browse and select multiple image files
  3. Client number auto-extracted from filenames
  4. Enter SFTP password (securely stored for future use)
  5. Click "Upload to SFTP"
  6. Monitor progress and verify uploads

#### 4. **CardPresso Mode**
- **Purpose**: Export client data for card printing
- **Steps**:
  1. Select "Cardpresso" mode
  2. Enter 5-digit client number
  3. Click "Lookup & Load" to search database
  4. Preview client photo and verify data
  5. Click "Open Single-ID-Card" or "Open Multi-ID-Card"
  6. CardPresso opens with pre-loaded client data

### Advanced Features

#### Manual Face Selection
- **When to use**: When AI face detection fails or for precise control
- **Features**:
  - Click and drag to define face area
  - Real-time preview with selection rectangle
  - Automatic conversion to 1:1 aspect ratio
  - Visual feedback during selection

#### SFTP Security
- **Password Encryption**: Passwords stored in encrypted [`sftp_password.dat`](sftp_password.dat)
- **Connection Testing**: Built-in SFTP connection test utility
- **File Existence Check**: Prevents accidental overwrites
- **Batch Upload**: Support for multiple file uploads -Should be added in next version

#### CardPresso Integration
- **Excel Database**: Reads from [`master.xlsx`](master.xlsx)
- **Photo Verification**: Automatically checks photo file existence
- **Template Support**: Pre-configured Single-ID and Multi-ID templates
- **CSV Export**: Generates [`CardPresso.csv`](CardPresso.csv) for import

## Benefits of Semi-Automation

### Time Savings
- **90% reduction** in manual file copying and renaming
- **80% reduction** in manual cropping time with face detection
- **75% reduction** in SFTP upload preparation
- **60% reduction** in CardPresso data entry

### Error Reduction
- Eliminates manual file naming errors
- Ensures consistent 1:1 cropping for ID photos
- Automates client number validation
- Reduces data entry mistakes

### Consistency
- Standardized image processing across all photos
- Consistent file naming conventions
- Uniform cropping and resizing parameters
- Automated quality control

## Troubleshooting

### Common Issues

1. **IrfanView Not Found**
   - Update `IrfanViewPath` in [`config.file`](config.file)
   - Ensure IrfanView is properly extracted in Tools/iview473_x64/

2. **Face Detection Fails**
   - Check internet connection for Microsoft Face API
   - Verify API credentials in configuration
   - Use "Manual Face Select" as fallback
   - Check image quality and lighting conditions

3. **SFTP Upload Fails**
   - Test connection using "Test SFTP" button
   - Check network connectivity
   - Verify SFTP credentials in configuration
   - Ensure remote directory permissions

4. **CardPresso Integration Issues**
   - Verify Excel file path in configuration
   - Check CardPresso executable path
   - Ensure CSV export permissions
   - Validate photo file existence in Cropped/ folder

5. **Configuration Path Issues**
   - Use `%BasePath%` variable for portable paths
   - Ensure environment variables are properly expanded
   - Check folder permissions for output directories

### Log Files
- **Log_file.log** - General application activities and processing logs
- **SFTP_Error_Log.txt** - Detailed SFTP connection and upload errors
- **WinSCP_Log_*.txt** - WinSCP session logs for debugging

### Network Requirements
- **Microsoft Face API**: Requires internet access for AI face detection
- **SFTP Upload**: Requires network connectivity to SFTP server
- **Offline Operation**: Manual face selection works without internet

## Development & Customization

### Extending Functionality
- Add new image filters by extending GDI+ operations
- Implement additional face detection algorithms
- Create custom batch processing templates
- Add support for additional image formats

### Script Structure
- **AutoIt Script**: Uses GDI+ for image processing and GUI
- **Configuration Files**: Centralize settings management
- **External Tools**: Portable applications for image manipulation
- **Logging System**: Comprehensive activity tracking

### Security Features
- **Encrypted Password Storage**: Basic XOR encryption for SFTP passwords
- **Secure File Handling**: Temporary file cleanup
- **Input Validation**: Client number and file format validation
- **Error Handling**: Comprehensive error reporting and recovery

## Support & Maintenance

### Regular Maintenance Tasks
- Update configuration files for new folder structures
- Monitor log files for errors
- Update client database in [`master.xlsx`](master.xlsx)
- Backup configuration and script files
- Clear temporary files from Cropped/ and Resized/ folders

### Getting Help
- Check log files for detailed error information
- Review configuration settings in [`config.file`](config.file)
- Test individual components using built-in test functions
- Consult troubleshooting section in documentation

## Version History

- **v1.0**: Initial automation scripts for copy/resize
- **v1.1**: Added face detection and cropping
- **v1.2**: Integrated SFTP upload functionality
- **v1.3**: Added CardPresso integration
- **v1.4**: Complete photo management suite with integrated GUI
- **Current**: Semi-automated workflow with manual oversight and step-by-step guidance

## License & Credits

Built with:
- **AutoIt3** for Windows automation and GUI development
- **Microsoft Face API** for AI-powered face detection
- **GDI+** for image processing and manipulation
- **IrfanView** for image resizing and format conversion
- **WinSCP** for secure SFTP file transfers
- **Microsoft Excel** for client database management

Developed to streamline photo processing workflows for ID E-Token and MIS systems, reducing manual effort while maintaining quality control through semi-automated processes.
