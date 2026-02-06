#!/usr/bin/env python3
"""
Keynote/PPT to Markdown Converter
Convert presentations to Markdown with embedded images and speaker notes.
"""

import sys
import os
import zipfile
import re
import argparse
from pathlib import Path
from xml.etree import ElementTree as ET

def extract_text_from_xml(xml_string):
    """Extract text from PPTX XML elements."""
    text_elements = []
    
    # Pattern for finding text in <a:t> elements
    pattern = r'<a:t[^>]*>([^<]*)</a:t>'
    matches = re.findall(pattern, xml_string)
    
    return '\n\n'.join([m.strip() for m in matches if m.strip()])

def get_slide_rels(zip_file, slide_target):
    """Get relationship mapping for a slide."""
    base_name = os.path.basename(slide_target)
    rels_path = f"ppt/slides/_rels/{base_name}.rels"
    
    if rels_path not in zip_file.namelist():
        return {}
    
    try:
        with zip_file.open(rels_path) as f:
            content = f.read().decode('utf-8')
            root = ET.fromstring(content)
            
            rels = {}
            ns = {'rels': 'http://schemas.openxmlformats.org/package/2006/relationships'}
            
            for elem in root.findall('.//rels:Relationship', ns):
                rid = elem.get('Id')
                target = elem.get('Target')
                rel_type = elem.get('Type')
                
                if 'notesSlide' in str(rel_type):
                    rels['notes'] = target
                elif 'image' in str(rel_type):
                    rels['image'] = target
                    
            return rels
    except Exception as e:
        print(f"Warning: Could not parse rels for {slide_target}: {e}")
        return {}

def get_presentation_rels(zip_file):
    """Get slide relationship mapping from presentation.xml.rels"""
    rels_path = 'ppt/_rels/presentation.xml.rels'
    
    if rels_path not in zip_file.namelist():
        return {}
    
    try:
        with zip_file.open(rels_path) as f:
            content = f.read().decode('utf-8')
            root = ET.fromstring(content)
            
            rels = {}
            ns = {'rels': 'http://schemas.openxmlformats.org/package/2006/relationships'}
            
            for elem in root.findall('.//rels:Relationship', ns):
                rid = elem.get('Id')
                target = elem.get('Target')
                rel_type = elem.get('Type')
                
                if 'slide' in str(rel_type):
                    rels[rid] = target
                    
            return rels
    except Exception as e:
        print(f"Warning: Could not parse presentation rels: {e}")
        return {}

def extract_notes(zip_file, slide_num, slide_path, images_dir):
    """Extract speaker notes from a slide."""
    image_file = f"{str(slide_num).zfill(3)}.png"
    
    if not os.path.exists(os.path.join(images_dir, image_file)):
        return None, f"![](images/{image_file})\n\n(Slide image not found)"
    
    notes_path = slide_path.replace('ppt/slides/', 'ppt/notesSlides/')
    notes_path = notes_path.replace('.xml', '.xml')
    
    if notes_path not in zip_file.namelist():
        return image_file, f"![](images/{image_file})\n\n(No speaker notes)"
    
    try:
        with zip_file.open(notes_path) as f:
            content = f.read().decode('utf-8')
            
            # Extract text from notes
            text = extract_text_from_xml(content)
            
            if text:
                return image_file, f"![](images/{image_file})\n\n{text}"
            else:
                return image_file, f"![](images/{image_file})\n\n(Empty notes)"
    except Exception as e:
        return image_file, f"![](images/{image_file})\n\n(Error reading notes: {e})"

def extract_images(zip_file, output_dir):
    """Extract images from PPTX."""
    images_dir = os.path.join(output_dir, 'images')
    os.makedirs(images_dir, exist_ok=True)
    
    # Get slide relationships
    slide_rels = get_presentation_rels(zip_file)
    
    # Find image references in slides
    image_files = {}
    
    for rid, slide_target in slide_rels.items():
        base_name = os.path.basename(slide_target)
        rels_path = f"ppt/slides/_rels/{base_name}.rels"
        
        if rels_path not in zip_file.namelist():
            continue
            
        try:
            with zip_file.open(rels_path) as f:
                content = f.read().decode('utf-8')
                root = ET.fromstring(content)
                
                ns = {'rels': 'http://schemas.openxmlformats.org/package/2006/relationships'}
                slide_num = 0
                
                for elem in root.findall('.//rels:Relationship', ns):
                    target = elem.get('Target')
                    rel_type = elem.get('Type')
                    
                    if 'image' in str(rel_type) and target:
                        slide_num += 1
                        if target.startswith('../'):
                            target = target[3:]
                        image_files[slide_num] = target
        except Exception:
            continue
    
    # Extract images
    for slide_num, image_path in image_files.items():
        if image_path in zip_file.namelist():
            ext = os.path.splitext(image_path)[1] or '.png'
            output_name = f"{str(slide_num).zfill(3)}{ext}"
            output_path = os.path.join(images_dir, output_name)
            
            try:
                with zip_file.open(image_path) as src, open(output_path, 'wb') as dst:
                    dst.write(src.read())
                print(f"  Extracted: {output_name}")
            except Exception as e:
                print(f"  Warning: Could not extract {image_path}: {e}")

def convert_pptx(pptx_file, output_dir):
    """Convert PPTX to Markdown."""
    if not os.path.exists(pptx_file):
        print(f"Error: File not found: {pptx_file}")
        return
    
    base_name = os.path.splitext(os.path.basename(pptx_file))[0]
    os.makedirs(output_dir, exist_ok=True)
    
    # Create images directory
    images_dir = os.path.join(output_dir, 'images')
    os.makedirs(images_dir, exist_ok=True)
    
    print(f"Converting: {pptx_file}")
    print(f"Output directory: {output_dir}")
    
    try:
        with zipfile.ZipFile(pptx_file, 'r') as zip_file:
            # Extract images first
            print("Extracting images...")
            extract_images(zip_file, output_dir)
            
            # Get presentation relationships
            slide_rels = get_presentation_rels(zip_file)
            
            # Process each slide
            markdown = ""
            slide_num = 0
            
            for rid, slide_target in sorted(slide_rels.items(), key=lambda x: x[0]):
                slide_num += 1
                image_file, notes = extract_notes(zip_file, slide_num, slide_target, images_dir)
                markdown += notes + "\n\n"
            
            # Write markdown
            md_path = os.path.join(output_dir, f"{base_name}.md")
            with open(md_path, 'w', encoding='utf-8') as f:
                f.write(markdown.strip())
            
            print(f"\nDone! Output: {md_path}")
            
    except zipfile.BadZipFile:
        print(f"Error: Not a valid ZIP file: {pptx_file}")
    except Exception as e:
        print(f"Error: {e}")

def main():
    parser = argparse.ArgumentParser(
        description='Convert Keynote/PPT to Markdown'
    )
    parser.add_argument('input', help='Input .key or .pptx file')
    parser.add_argument('-o', '--output', default=None, help='Output directory')
    
    args = parser.parse_args()
    
    input_file = args.input
    if not os.path.exists(input_file):
        print(f"Error: File not found: {input_file}")
        sys.exit(1)
    
    output_dir = args.output or os.path.dirname(os.path.abspath(input_file))
    
    if input_file.endswith('.pptx'):
        convert_pptx(input_file, output_dir)
    elif input_file.endswith('.key'):
        print("Keynote files require macOS with Keynote installed.")
        print("Please convert to PPTX first, or use the CLI on macOS.")
        sys.exit(1)
    else:
        print("Error: Input file must be .pptx or .key")
        sys.exit(1)

if __name__ == '__main__':
    main()
