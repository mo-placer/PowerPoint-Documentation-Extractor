#!/usr/bin/env python3
"""
Streamlit App for PowerPoint Documentation Updates Extractor

This app provides a web interface for extracting components from PowerPoint files:
- Vocabulary words and definitions
- Session Goals
- Assessment items
- Related Careers
- Session Materials
"""

import streamlit as st
import tempfile
import os
from pathlib import Path
import zipfile
from io import BytesIO
from collections import defaultdict

from doc_updates import DocumentationExtractor


def main():
    st.set_page_config(
        page_title="PowerPoint Documentation Extractor",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 PowerPoint Documentation Extractor")
    st.markdown("Extract vocabulary, goals, assessments, careers, and materials from your PowerPoint module storyboards.")
    
    # Sidebar with instructions
    with st.sidebar:
        st.header("Instructions")
        st.markdown("""
        1. **Upload PowerPoint files** (up to 7 .pptx files)
        2. **Enter module acronym** (4-letter code like MATS, CHEM, etc.)
        3. **Click Process** to extract content
        4. **Download** the generated Word document
        
        ### What gets extracted:
        - 🔤 **Vocabulary**: Blue/turquoise bold text with definitions
        - 🎯 **Session Goals**: Content after "In today's session, you will"
        - 📝 **Assessments**: Items after instructor evaluation text
        - 💼 **Careers**: Related career listings
        - 📋 **Materials**: Required session materials
        """)
        
        st.header("Debug Mode")
        debug_mode = st.checkbox("Enable debug output", help="Shows detailed color and formatting analysis")
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("Upload PowerPoint Files")
        uploaded_files = st.file_uploader(
            "Choose PowerPoint files (.pptx)",
            type=['pptx'],
            accept_multiple_files=True,
            help="Upload up to 7 PowerPoint files from your module storyboards"
        )
        
        if uploaded_files:
            if len(uploaded_files) > 7:
                st.error("⚠️ Please upload no more than 7 files at once.")
                return
            
            st.success(f"✅ {len(uploaded_files)} file(s) uploaded successfully!")
            
            # Display uploaded files
            with st.expander("📁 Uploaded Files", expanded=True):
                for file in uploaded_files:
                    st.write(f"• {file.name} ({file.size / 1024:.1f} KB)")
    
    with col2:
        st.header("Module Settings")
        
        # Auto-detect acronym from first file if available
        default_acronym = ""
        if uploaded_files:
            first_filename = uploaded_files[0].name
            base_name = os.path.splitext(first_filename)[0]
            parts = base_name.split('_')
            if parts:
                default_acronym = parts[0]
        
        module_acronym = st.text_input(
            "Module Acronym",
            value=default_acronym,
            max_chars=4,
            help="4-letter module code (e.g., MATS, CHEM, PHYS)"
        ).upper()
        
        if module_acronym and len(module_acronym) != 4:
            st.warning("⚠️ Acronym should be exactly 4 letters")
    
    # Processing section
    if uploaded_files and module_acronym and len(module_acronym) == 4:
        st.header("Process Files")
        
        if st.button("🚀 Extract Documentation", type="primary"):
            process_files(uploaded_files, module_acronym, debug_mode)


def process_files(uploaded_files, module_acronym, debug_mode):
    """Process the uploaded PowerPoint files and generate Word document."""
    
    # Create progress tracking
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # Initialize extractor
        extractor = DocumentationExtractor()
        
        # Create temporary directory for processing
        with tempfile.TemporaryDirectory() as temp_dir:
            status_text.text("📁 Preparing files...")
            progress_bar.progress(10)
            
            # Save uploaded files to temp directory
            temp_files = []
            for i, uploaded_file in enumerate(uploaded_files):
                temp_path = os.path.join(temp_dir, uploaded_file.name)
                with open(temp_path, 'wb') as f:
                    f.write(uploaded_file.getbuffer())
                temp_files.append(temp_path)
                progress_bar.progress(10 + (i + 1) * 15 // len(uploaded_files))
            
            status_text.text("🔍 Extracting content from PowerPoint files...")
            progress_bar.progress(30)
            
            # Process all files
            all_content = defaultdict(lambda: {
                'vocabulary': [],
                'goals': [],
                'assessments': [],
                'careers': [],
                'materials': []
            })
            
            # Create containers for real-time feedback
            results_container = st.container()
            
            for i, temp_file in enumerate(temp_files):
                filename = os.path.basename(temp_file)
                status_text.text(f"📖 Processing {filename}...")
                
                with results_container:
                    with st.expander(f"📄 {filename}", expanded=False):
                        if debug_mode:
                            st.write("Debug output:")
                            debug_output = st.empty()
                        
                        # Process file
                        content = extractor.extract_all_content_from_ppt(temp_file, debug_mode)
                        
                        if any(content.values()):
                            all_content[filename] = content
                            
                            # Display results
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("Vocabulary", len(content['vocabulary']))
                                if content['vocabulary']:
                                    for word, definition in content['vocabulary'][:3]:  # Show first 3
                                        st.write(f"• {word}")
                                    if len(content['vocabulary']) > 3:
                                        st.write(f"... and {len(content['vocabulary']) - 3} more")
                            
                            with col2:
                                st.metric("Goals", len(content['goals']))
                                if content['goals']:
                                    for goal in content['goals'][:2]:  # Show first 2
                                        st.write(f"• {goal[:50]}{'...' if len(goal) > 50 else ''}")
                                    if len(content['goals']) > 2:
                                        st.write(f"... and {len(content['goals']) - 2} more")
                            
                            with col3:
                                st.metric("Assessments", len(content['assessments']))
                                if content['assessments']:
                                    for assessment in content['assessments'][:2]:  # Show first 2
                                        st.write(f"• {assessment[:50]}{'...' if len(assessment) > 50 else ''}")
                                    if len(content['assessments']) > 2:
                                        st.write(f"... and {len(content['assessments']) - 2} more")
                        else:
                            st.warning("No content extracted from this file")
                
                progress_bar.progress(30 + (i + 1) * 50 // len(temp_files))
            
            if not all_content:
                st.error("❌ No content found in any PowerPoint files.")
                return
            
            status_text.text("📝 Generating Word document...")
            progress_bar.progress(85)
            
            # Generate Word document
            output_filename = f"{module_acronym}_Doc Updates & Tickets.docx"
            output_path = os.path.join(temp_dir, output_filename)
            
            extractor.create_word_document(all_content, output_path, module_acronym)
            
            progress_bar.progress(95)
            status_text.text("✅ Processing complete!")
            
            # Calculate totals for summary
            total_vocab = sum(len(content['vocabulary']) for content in all_content.values())
            total_goals = sum(len(content['goals']) for content in all_content.values())
            total_assessments = sum(len(content['assessments']) for content in all_content.values())
            total_careers = sum(len(content.get('careers', [])) for content in all_content.values())
            total_materials = sum(len(content.get('materials', [])) for content in all_content.values())
            
            # Display summary
            st.success("🎉 Documentation extraction completed successfully!")
            
            summary_col1, summary_col2, summary_col3 = st.columns(3)
            with summary_col1:
                st.metric("Total Vocabulary", total_vocab)
                st.metric("Total Goals", total_goals)
            with summary_col2:
                st.metric("Total Assessments", total_assessments)
                st.metric("Total Careers", total_careers)
            with summary_col3:
                st.metric("Total Materials", total_materials)
                st.metric("Files Processed", len(all_content))
            
            # Provide download
            with open(output_path, 'rb') as f:
                doc_bytes = f.read()
            
            st.download_button(
                label="📥 Download Word Document",
                data=doc_bytes,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary"
            )
            
            progress_bar.progress(100)
            status_text.text("🎯 Ready for download!")
            
    except Exception as e:
        st.error(f"❌ An error occurred: {str(e)}")
        if debug_mode:
            st.exception(e)


if __name__ == "__main__":
    main()