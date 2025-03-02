# Resume Optimizer Pro - Complete Setup and Usage Guide

This comprehensive guide will help you set up and use the Resume Optimizer Pro application to tailor your resume for specific job descriptions and improve your chances of getting past Applicant Tracking Systems (ATS).

## 1. System Requirements

- **Operating System**: Windows, macOS, or Linux
- **Python**: Version 3.8 or higher
- **Internet Connection**: Required for API calls
- **OpenAI API Key**: Required for resume optimization

## 2. Installation Guide

### Step 1: Install Python

1. Download Python 3.8+ from [python.org](https://python.org)
2. During installation:
   - Windows: Check "Add Python to PATH"
   - macOS/Linux: Installation typically configures PATH automatically
3. Verify installation by opening a terminal/command prompt and typing:
   ```
   python --version
   ```

### Step 2: Create Project Folder

1. Create a new folder named "ResumeOptimizerPro" on your computer
2. This will store all project files

### Step 3: Get OpenAI API Key

1. Visit [platform.openai.com](https://platform.openai.com)
2. Create an account or sign in
3. Navigate to API Keys section
4. Click "Create new secret key"
5. Copy the key (you'll need it later)
6. Note: This requires payment information on file with OpenAI

### Step 4: Create Project Files

1. **Create the requirements.txt file**
   - In your project folder, create a file named `requirements.txt`
   - Copy and paste the following content:
   ```
   beautifulsoup4==4.12.2
   python-docx==0.8.11
   PyPDF2==3.0.1
   openai==1.3.0
   requests==2.31.0
   python-dotenv==1.0.0
   customtkinter==5.2.0
   reportlab==4.0.4
   ```

2. **Create the main script file**
   - Create a file named `resume_optimizer_pro.py`
   - Copy and paste the code from the provided "Final Resume Optimizer with Enhanced Formatting" file

### Step 5: Install Dependencies

1. Open terminal/command prompt
2. Navigate to your project folder:
   ```
   cd path/to/ResumeOptimizerPro
   ```
3. Install required packages:
   ```
   pip install -r requirements.txt
   ```
4. Wait for all packages to install

## 3. Running the Application

1. Open terminal/command prompt
2. Navigate to your project folder:
   ```
   cd path/to/ResumeOptimizerPro
   ```
3. Run the application:
   ```
   python resume_optimizer_pro.py
   ```
4. If this is your first time running the app, you'll be prompted to enter your OpenAI API key

## 4. Using the Application

### Step 1: Prepare Your Materials

1. **Resume**: Have your resume ready in PDF or DOCX format
2. **Job Description**: Either have a job posting URL or the text of the job description

### Step 2: Enter Job Information

1. **Option 1 - URL Method**:
   - Paste the job posting URL in the "Job Posting URL" field
   - The application will automatically extract the job description

2. **Option 2 - Direct Text Method**:
   - Paste the job description text into the "Or paste job description" text area

### Step 3: Select Your Resume

1. Click the "Browse" button
2. Navigate to and select your resume file (PDF or DOCX format)

### Step 4: Choose Output Format

Select your preferred output format:
- **Same as input**: Maintains the original file format (PDF or DOCX)
- **Plain text (.txt)**: Simple text format
- **Word (.docx)**: Microsoft Word document
- **PDF (.pdf)**: PDF document

### Step 5: Process Your Resume

1. Click the "Optimize Resume" button
2. Wait for processing (typically 1-2 minutes)
3. Progress bar will show the current status

### Step 6: Review Results

1. **Analysis Report Tab**:
   - Strengths: Where your resume aligns with the job
   - Weaknesses: Missing elements
   - Areas for Improvement: Specific suggestions
   - Keyword Analysis: Important terms to include
   - ATS Compatibility Score: Estimated match percentage

2. **Optimized Resume Tab**:
   - View your tailored resume text

### Step 7: Save Your Results

1. **Save Analysis Report**: 
   - Click "Save Analysis Report" 
   - Choose location and filename

2. **Save Plain Text Resume**:
   - Click "Save as Plain Text"
   - Choose location and filename

3. **Save Formatted Resume**:
   - Click "Save Formatted Resume"
   - Choose location and filename
   - The application will create a formatted document in your selected output format

## 5. Formatting Features

The Resume Optimizer Pro maintains professional formatting throughout the optimization process:

### Preserved Formatting Elements

1. **Bold Headings and Text**
   - Section headings (like "EXPERIENCE," "EDUCATION") remain bold
   - Important text and company names stay bold
   - Bold formatting is preserved in all output formats

2. **Text Alignment**
   - Left, center, and right-justified text maintains proper alignment
   - Original document indentation patterns are preserved
   - PDF and DOCX outputs respect these alignment settings

3. **Key Metrics Highlighting**
   - Important numbers and achievements (like "increased sales by 45%") are highlighted
   - Highlighted metrics stand out to recruiters while maintaining ATS compatibility
   - Different highlighting styles for each output format (yellow in DOCX, light background in PDF)

4. **Hyperlinks**
   - Email addresses, LinkedIn profiles, and websites remain properly formatted as hyperlinks
   - Links are clickable in the final document formats
   - URLs are properly formatted in plain text output

5. **Bullet Points**
   - Bullet formatting is preserved with appropriate indentation
   - Bullet hierarchy and structure is maintained
   - Consistent formatting across the document

6. **Document Structure**
   - The overall structure, spacing, and flow of the document is maintained
   - Section ordering and hierarchy remains consistent
   - Paragraph breaks and spacing mirror the original document

## 6. Troubleshooting

### Common Issues

1. **API Key Errors**:
   - Ensure your OpenAI API key is valid
   - Check that you have credit available on your OpenAI account
   - The app stores your key in a `.env` file for future use

2. **File Format Issues**:
   - Only PDF and DOCX formats are supported
   - If text extraction fails, try converting your resume to a different format
   - PDF files with complex layouts might lose some formatting

3. **Job Description Extraction Failures**:
   - Some websites block automatic extraction
   - Use the "paste job description" option in these cases
   - Try copying the job description text manually

4. **Formatting in Output**:
   - Complex document layouts may be simplified
   - Some advanced formatting (tables, columns) might be adjusted
   - Review the formatted output before sending to employers

### Getting Help

If you encounter issues not covered here, consider:
1. Checking error messages for specific details
2. Ensuring all dependencies are correctly installed
3. Verifying your original resume file can be opened normally

## 7. Tips for Best Results

1. **Use DOCX format when possible**
   - DOCX files tend to preserve more formatting information
   - Better results for formatted output generation

2. **Keep original formatting relatively simple**
   - Clean, professional layouts work best
   - Avoid overly complex design elements

3. **Review the optimized content**
   - The AI optimizes content while preserving format
   - Verify all information remains accurate

4. **Test with multiple job descriptions**
   - Create different versions optimized for specific roles
   - Compare how optimization changes for different positions

5. **Use the Analysis Report**
   - The detailed analysis helps identify strengths and weaknesses
   - Make additional manual improvements based on feedback

Now you're ready to use Resume Optimizer Pro to create tailored, well-formatted, ATS-friendly resumes that highlight your relevant skills and experiences!
