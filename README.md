# PowerPoint Documentation Updates Extractor

Extract vocabulary, goals, assessments, careers, and materials from PowerPoint module storyboards through both a web interface and command-line tool.

## 🌐 Web Application (Streamlit)

### Live Demo
Access the app at: [Your Render URL here]

### Features
- **Easy file upload**: Drag & drop up to 7 PowerPoint files
- **Real-time processing**: See extraction progress and results
- **Instant download**: Get your Word document immediately
- **Debug mode**: Detailed analysis of color and formatting detection

### What It Extracts

1. **Vocabulary Terms**: Blue/turquoise and bold formatted words with definitions
2. **Session Goals**: Bulleted lists that appear after "In today's session, you will"
3. **Assessment Items**: Numbered lists that appear after instructor evaluation text
4. **Related Careers**: Career listings from slide content
5. **Session Materials**: Required materials and equipment lists

## 🖥️ Command Line Usage

### Installation

```bash
pip install -r requirements.txt
```

### Basic Usage
```bash
python3 doc_updates.py
```
This will:
- Process all `.pptx` files in the current directory
- Extract vocabulary, goals, assessments, careers, and materials
- Generate a Word document named `{ACRONYM}_Doc Updates & Tickets.docx`

### Advanced Usage
```bash
# Specify directory
python3 doc_updates.py -d /path/to/powerpoint/files

# Specify custom acronym
python3 doc_updates.py -a CHEM

# Custom output filename
python3 doc_updates.py -o "My_Custom_Document.docx"

# Enable debug mode to see detailed color and formatting analysis
python3 doc_updates.py --debug
```

## 🚀 Deployment

### Deploy to Render
1. Push this repository to GitHub
2. Connect your GitHub repo to Render
3. Render will automatically detect the `render.yaml` configuration
4. Your app will be live at `https://your-app-name.onrender.com`

### Local Development
```bash
# Install dependencies
pip install -r requirements.txt

# Run Streamlit app
streamlit run streamlit_app.py

# Run command-line version
python3 doc_updates.py
```

## Output

The script generates a Word document with three main sections:

1. **Vocabulary Terms** - Alphabetically sorted with definitions and source files
2. **Session Goals** - Organized by session with bulleted learning objectives
3. **Assessment Items** - Organized by session with numbered evaluation criteria

## File Naming Convention

The output file automatically uses the 4-letter acronym from the PowerPoint filenames (e.g., "MATS" from "MATS_Session_1.pptx") and appends "_Doc Updates & Tickets.docx".

## Examples

From the example MATS sessions, the script extracted:
- 11 vocabulary terms (atom, Elements, Protons, etc.)
- 12 session goals across 6 sessions
- 8 assessment items from 3 sessions

## Troubleshooting

- If no vocabulary is found, try using `--debug` to see color detection details
- The script looks for blue/turquoise colored text that is also bold
- Session goals must follow the specific trigger text pattern
- Assessment items must follow the instructor evaluation text pattern
