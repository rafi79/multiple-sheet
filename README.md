# Excel to LLM Processor

A powerful web application that processes multiple Excel files and analyzes them using Google's Gemini AI. Optimized for efficient token usage and deployed on Vercel.

## Features

- 📁 Upload multiple Excel files simultaneously
- 📊 Read all sheets from each file automatically
- 🔍 Smart column detection and data type analysis
- 💡 Token-optimized data summarization
- 🤖 AI-powered analysis using Gemini
- 🌐 Web-based interface with drag & drop
- ⚡ Fast processing with Vercel serverless functions

## Deployment Instructions

### 1. Clone and Setup
```bash
git clone <your-repo>
cd excel-llm-processor
```

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```

### 3. Deploy to Vercel
```bash
# Install Vercel CLI
npm install -g vercel

# Deploy
vercel --prod
```

### 4. Usage
1. Visit your deployed Vercel URL
2. Enter your Gemini API key
3. Upload one or more Excel files
4. Optionally add a specific question
5. Click "Process & Analyze"

## Key Optimizations

- **Token Efficiency**: Limits rows per sheet (100 max) and cell content (500 chars max)
- **Smart Summarization**: Creates structured summaries instead of raw data dumps
- **Multi-file Support**: Processes 2-4 Excel files simultaneously
- **Error Handling**: Graceful handling of corrupted or invalid files
- **Data Type Detection**: Automatic column analysis for better insights

## API Key Setup

Get your Gemini API key from [Google AI Studio](https://makersuite.google.com/app/apikey) and enter it in the web interface.

## Supported Formats

- .xlsx (Excel 2007+)
- .xls (Excel 97-2003)

## Limitations

- Maximum 100 rows per sheet (configurable)
- Maximum 500 characters per cell (configurable)
- File size limited by Vercel (10MB)

## Security

- API keys are not stored server-side
- Temporary files are automatically cleaned up
- No data persistence on server
