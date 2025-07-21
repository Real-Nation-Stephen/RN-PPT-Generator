# RN PowerPoint Genie - Streamlit Version

A modern web application for generating PowerPoint presentations from images, built with Streamlit and styled to match the RN Contrast Checker aesthetic.

## Features

- 🖼️ **Multi-Image Upload**: Upload multiple images at once
- 📐 **Auto-Resize**: Automatically format images to 16:9 aspect ratio
- 🎨 **Modern UI**: Clean, professional interface with Sparkle/Studio mode themes
- 📱 **Responsive**: Works on desktop and mobile devices
- 🚀 **Fast Processing**: Efficient PowerPoint generation
- 📥 **Direct Download**: Download presentations immediately

## Installation

### Prerequisites
- Python 3.7 or higher
- pip package manager

### Setup

1. **Clone or download the files**
   ```bash
   # If you have git
   git clone <repository-url>
   cd ppt-image-app
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements_streamlit.txt
   ```

3. **Run the application**
   ```bash
   streamlit run ppt_genie_streamlit.py
   ```

4. **Open in browser**
   - The app will automatically open in your default browser
   - If not, navigate to `http://localhost:8501`

## Usage

### Step-by-Step Guide

1. **Upload Images**
   - Click "Upload Images" button
   - Select multiple image files (JPG, PNG supported)
   - Images will appear in the preview section

2. **Configure Options**
   - ✅ **Auto Resize to 16:9**: Formats images for widescreen presentations
   - ❌ **Auto Resize to 16:9**: Keeps original image dimensions

3. **Generate Presentation**
   - Click "🚀 Generate Presentation"
   - Wait for processing to complete
   - Download button will appear when ready

4. **Download**
   - Click "📥 Download PowerPoint"
   - File will be saved with timestamp

### Tips for Best Results

- Use high-quality images (1920x1080 or higher recommended)
- Images are processed in the order they appear in the file list
- Each image becomes one slide in the presentation
- Auto-resize maintains aspect ratio while fitting to 16:9 slides

## Styling Modes

The app includes two visual themes matching the RN Contrast Checker:

- **✨ Sparkle Mode**: Colorful gradients and modern styling
- **🏢 Studio Mode**: Professional, minimal design

## Technical Details

### Architecture
- **Frontend**: Streamlit web framework
- **Backend**: Python with python-pptx library
- **Image Processing**: PIL (Python Imaging Library)
- **Styling**: Custom CSS matching RN design system

### File Processing
- Images are temporarily stored during processing
- Automatic cleanup of temporary files
- Error handling for corrupted images
- Memory-efficient processing

### Output Format
- PowerPoint (.pptx) format
- 16:9 widescreen slides (1920x1080)
- Blank layout with full-slide images
- Compatible with PowerPoint 2010+

## Deployment

### Local Development
```bash
streamlit run ppt_genie_streamlit.py
```

### Production Deployment

#### Streamlit Cloud
1. Push code to GitHub repository
2. Connect to Streamlit Cloud
3. Deploy directly from repository

#### Docker
```dockerfile
FROM python:3.9-slim

WORKDIR /app
COPY requirements_streamlit.txt .
RUN pip install -r requirements_streamlit.txt

COPY . .
EXPOSE 8501

CMD ["streamlit", "run", "ppt_genie_streamlit.py", "--server.port=8501", "--server.address=0.0.0.0"]
```

#### Heroku
```bash
# Create Procfile
echo "web: streamlit run ppt_genie_streamlit.py --server.port=$PORT --server.address=0.0.0.0" > Procfile

# Deploy
heroku create your-app-name
git push heroku main
```

## Comparison with Electron Version

| Feature | Electron App | Streamlit App |
|---------|--------------|---------------|
| **Platform** | Desktop (macOS/Windows/Linux) | Web-based (any browser) |
| **Installation** | App bundle/installer | Python environment |
| **File Access** | Native file system | Upload-based |
| **Performance** | Native performance | Web performance |
| **Updates** | App store/manual | Instant (web) |
| **Sharing** | Send app file | Share URL |

## Troubleshooting

### Common Issues

1. **Import Errors**
   ```bash
   # Reinstall dependencies
   pip install --upgrade -r requirements_streamlit.txt
   ```

2. **Memory Issues with Large Images**
   - Reduce image file sizes before upload
   - Process fewer images at once
   - Use image compression tools

3. **PowerPoint Generation Fails**
   - Ensure images are valid formats (JPG, PNG)
   - Check for corrupted image files
   - Verify sufficient disk space

### Performance Optimization

- Use compressed images when possible
- Limit to 50 images per presentation
- Close browser tabs to free memory
- Use modern browser for best performance

## Development

### Project Structure
```
ppt-image-app/
├── ppt_genie_streamlit.py      # Main Streamlit application
├── requirements_streamlit.txt   # Python dependencies
├── README_streamlit.md         # This file
└── assets/                     # Logo and image assets (optional)
```

### Customization

#### Adding New Features
- Edit `ppt_genie_streamlit.py`
- Add new functions for additional processing
- Update UI components as needed

#### Styling Changes
- Modify CSS in the `st.markdown()` sections
- Update color schemes and gradients
- Add new theme modes

#### Logo Customization
- Place logo files in `assets/` directory
- Update logo paths in the application
- Support for PNG, SVG formats

## License

This project maintains the same license as the original RN PowerPoint Genie application.

## Support

For issues and questions:
1. Check the troubleshooting section above
2. Review Streamlit documentation
3. Check python-pptx library documentation

---

**RN PowerPoint Genie Streamlit** - Transform your images into professional presentations with ease! ✨ 